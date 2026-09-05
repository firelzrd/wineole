//! Waking an STA thread that must serve two masters.
//!
//! A COM apartment-threaded session thread receives event callbacks through
//! the Windows message queue, so it has to pump messages. It also has to
//! service a command channel. `MsgWaitForMultipleObjects` waits on both: a
//! Win32 event for the channel, and QS_ALLINPUT for the message queue.
//!
//! PostThreadMessage would have let one queue carry both, and was rejected.
//! Thread messages are not associated with a window, and a nested modal
//! loop -- an Excel dialog, or COM's own modal loop while an outbound call
//! is in flight -- pumps messages and can discard them. A lost wake means a
//! command is never processed. An event handle is not affected by nested
//! loops.

use windows::Win32::Foundation::{
    CloseHandle, GetLastError, HANDLE, WAIT_EVENT, WAIT_FAILED, WAIT_OBJECT_0, WAIT_TIMEOUT,
};
use windows::Win32::System::Threading::{CreateEventW, SetEvent, INFINITE};
use windows::Win32::UI::WindowsAndMessaging::{
    DispatchMessageW, MsgWaitForMultipleObjects, PeekMessageW, TranslateMessage, MSG, PM_REMOVE,
    QS_ALLINPUT,
};

/// An auto-reset Win32 event used to wake the session thread when a command
/// has been queued.
///
/// Auto-reset rather than manual-reset on purpose: the waiter consumes the
/// signal, so a single `wake` produces exactly one wake-up. The loop then
/// drains the channel until it is empty, which is what makes several
/// commands queued behind one signal safe.
pub struct Waker(HANDLE);

// The handle is owned by this struct and is only ever passed to Win32 calls
// that are themselves thread-safe (SetEvent, MsgWaitForMultipleObjects).
unsafe impl Send for Waker {}
unsafe impl Sync for Waker {}

impl Waker {
    pub fn new() -> windows::core::Result<Self> {
        // (security attrs, manual_reset = false, initial_state = false, name)
        let h = unsafe { CreateEventW(None, false, false, None)? };
        Ok(Waker(h))
    }

    pub fn wake(&self) {
        unsafe {
            let _ = SetEvent(self.0);
        }
    }

    pub fn handle(&self) -> HANDLE {
        self.0
    }
}

impl Drop for Waker {
    fn drop(&mut self) {
        unsafe {
            let _ = CloseHandle(self.0);
        }
    }
}

#[derive(Debug)]
pub enum Wake {
    Commands,
    Messages,
    /// The wait itself failed; the payload is `GetLastError`.
    ///
    /// This is not a wake-up, it is the absence of one, and the caller must
    /// treat it as fatal to the wait loop. Folding it into `Messages` -- which
    /// is what the obvious `else` branch does -- turns a failing wait into an
    /// unthrottled spin: nothing to peek, so the loop comes straight back and
    /// fails again, at 100% CPU, silently and forever.
    Failed(u32),
}

/// The one place a `MsgWaitForMultipleObjects` return value is interpreted, so
/// `wait` and `wait_timeout` cannot drift apart on what a failure means.
///
/// With a single handle the documented returns are `WAIT_OBJECT_0` (the
/// event), `WAIT_OBJECT_0 + 1` (the message queue), `WAIT_TIMEOUT` and
/// `WAIT_FAILED`. `WAIT_TIMEOUT` is handled by the callers, since only one of
/// them can see it.
fn classify(r: WAIT_EVENT) -> Wake {
    if r == WAIT_OBJECT_0 {
        Wake::Commands
    } else if r == WAIT_FAILED {
        Wake::Failed(unsafe { GetLastError().0 })
    } else {
        Wake::Messages
    }
}

/// Block until either a command was queued or the message queue has work.
pub fn wait(h: HANDLE) -> Wake {
    let handles = [h];
    let r = unsafe { MsgWaitForMultipleObjects(Some(&handles), false, INFINITE, QS_ALLINPUT) };
    classify(r)
}

/// Like `wait`, but gives up after `ms` milliseconds. Only the session loop
/// wants an unbounded wait; a test driving the pump by hand needs a way to
/// stop, and a `sleep`-and-check loop would be the polling this whole design
/// exists to avoid.
pub fn wait_timeout(h: HANDLE, ms: u32) -> Option<Wake> {
    let handles = [h];
    let r = unsafe { MsgWaitForMultipleObjects(Some(&handles), false, ms, QS_ALLINPUT) };
    if r == WAIT_TIMEOUT {
        None
    } else {
        Some(classify(r))
    }
}

/// Dispatch everything currently in the message queue. COM event callbacks
/// arrive through here.
pub fn drain_messages() {
    let mut msg = MSG::default();
    unsafe {
        while PeekMessageW(&mut msg, None, 0, 0, PM_REMOVE).as_bool() {
            let _ = TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Arc;
    use windows::Win32::Foundation::{LPARAM, WPARAM};
    use windows::Win32::System::Threading::GetCurrentThreadId;
    use windows::Win32::UI::WindowsAndMessaging::{PostThreadMessageW, WM_USER};

    #[test]
    fn test_wait_timeout_returns_none_when_nothing_happens() {
        let waker = Waker::new().expect("Waker::new");
        let start = std::time::Instant::now();
        assert!(wait_timeout(waker.handle(), 150).is_none(), "nothing signalled it, so it timed out");
        assert!(
            start.elapsed() >= std::time::Duration::from_millis(100),
            "it must actually wait, not return immediately"
        );
    }

    #[test]
    fn test_waker_wakes_the_wait_and_auto_resets() {
        // `Arc<Waker>`, not a raw HANDLE smuggled through a usize: the
        // signalling thread outlives the first assertion below, and if that
        // assertion panicked while it held only an integer, `Waker::drop`
        // would run `CloseHandle` and the thread would then `SetEvent` on a
        // closed -- and, in a single-process test binary, quite possibly
        // recycled -- handle belonging to some other test.
        let waker = Arc::new(Waker::new().expect("Waker::new"));
        waker.wake();
        assert!(matches!(wait(waker.handle()), Wake::Commands));

        // Auto-reset, not manual: a second wait must NOT return Commands
        // immediately off the same signal. If it did, one wake would be
        // consumed twice and the loop would spin.
        //
        // Verified by waking from another thread after a delay -- if the
        // event were manual-reset this returns instantly instead, and the
        // elapsed time gives it away.
        let signaller = waker.clone();
        std::thread::spawn(move || {
            std::thread::sleep(std::time::Duration::from_millis(200));
            signaller.wake();
        });
        let start = std::time::Instant::now();
        assert!(matches!(wait(waker.handle()), Wake::Commands));
        assert!(
            start.elapsed() >= std::time::Duration::from_millis(150),
            "the event must be auto-reset: the first wake was still signalled, \
             so this wait returned without anyone signalling it again"
        );
    }

    #[test]
    fn test_a_queued_message_wakes_the_wait_and_drain_messages_empties_the_queue() {
        // The message half of the pump, which is the whole reason the loop
        // waits on QS_ALLINPUT instead of just recv-ing: a wake that is not a
        // command has to be reported as `Messages`, and `drain_messages` has
        // to actually take the messages out of the queue. If it did not, the
        // session loop would be woken by every undelivered message for as long
        // as it stayed queued -- a spin, not a wait.
        //
        // A posted thread message stands in for the COM event callback that
        // the sink will deliver through this same queue, without needing COM.
        let waker = Waker::new().expect("Waker::new");

        // PostThreadMessageW fails with ERROR_INVALID_THREAD_ID against a
        // thread that has never had a message queue, and a freshly spawned
        // test thread has not. Peeking creates one.
        drain_messages();
        unsafe {
            PostThreadMessageW(GetCurrentThreadId(), WM_USER, WPARAM(0), LPARAM(0))
                .expect("PostThreadMessageW");
        }

        // Nothing signalled the event, so this can only be the queue.
        assert!(
            matches!(wait_timeout(waker.handle(), 1000), Some(Wake::Messages)),
            "a queued message must wake the wait as Messages"
        );

        drain_messages();

        // MsgWaitForMultipleObjects reports the queue as signalled until a
        // Peek/GetMessage has taken the input out of it, so a timeout here is
        // the assertion that drain_messages really drained.
        //
        // A zero timeout is the whole assertion, not a shortcut: an undrained
        // queue is signalled *now* and comes back immediately, so waiting any
        // longer would only make the test slower, and the suite's real-Excel
        // tests are sensitive to how long anything else occupies a test thread.
        assert!(
            wait_timeout(waker.handle(), 0).is_none(),
            "drain_messages must empty the queue: the wait has nothing left to report"
        );
    }

    #[test]
    fn test_a_failed_wait_is_reported_as_failed_not_as_messages() {
        // A failing wait must never be mistaken for a message wake. The
        // session loop answers `Messages` by peeking a queue that has nothing
        // in it and waiting again -- so a persistent failure mapped to
        // `Messages` is an unthrottled spin at 100% CPU, silently, forever.
        // `Failed` is fatal to the loop instead.
        //
        // A NULL handle is the cheapest way to make the wait itself fail.
        let bogus = HANDLE(std::ptr::null_mut());
        match wait_timeout(bogus, 0) {
            Some(Wake::Failed(err)) => assert_ne!(err, 0, "GetLastError should say why"),
            other => panic!("expected Wake::Failed for an invalid handle, got {:?}", other),
        }
    }
}
