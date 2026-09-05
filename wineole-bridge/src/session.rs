use std::collections::HashMap;
use std::sync::atomic::{AtomicU32, Ordering};
use std::sync::mpsc::{channel, Receiver, SendError, Sender};
use std::sync::Arc;
use std::thread::JoinHandle;
use serde_json::Value;
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED, IDispatch};
use crate::dispatch::{ComError, ComResult};
use crate::pump::{self, Wake, Waker};
use crate::registry::{CleanupConfig, InstanceRegistry, JoinResult, LeaveDecision};
use crate::identity::{instance_key, synthetic_key, InstanceKey};
use crate::{dispatch, sink, value};

/// What a `Release` of a root handle resolves to.
///
/// `Done` is the ordinary case: the handle is gone (and, for a participating
/// session's root, any cleanup steps have run). `CleanupPending(seq)` means a
/// live client's closure has been asked to run first, delivered as a
/// `$cleanup` event with this seq; the client answers with `release_event`.
pub enum ReleaseOutcome {
    Done,
    // Sent by the `$cleanup` client-callback path when a live client's closure
    // must run first: the root release is answered with this seq instead of
    // `Done`, the session stays alive, and it completes on the client's
    // matching `release_event` (see the `Release`/`ReleaseEvent` arms).
    CleanupPending(u64),
}

pub enum SessionCommand {
    Invoke {
        handle: u64,
        name: String,
        args: Vec<Value>,
        named: HashMap<String, Value>,
        reply: Sender<Result<(Value, Vec<u64>), (String, String)>>,
    },
    ConstLoad {
        handle: u64,
        reply: Sender<Result<Value, (String, String)>>,
    },
    Release {
        handle: u64,
        /// True when this release came from the wire `release` method (the
        /// client is alive and can run a closure); false on connection
        /// teardown (`release_all`), where there is nobody to consult.
        ///
        /// The `$cleanup` path in the `Release` arm consults this to decide
        /// whether a live client's closure can run first: a dead client (a
        /// disconnect) skips the callback and quits inline instead.
        client_alive: bool,
        reply: Sender<ReleaseOutcome>,
    },
    /// Ask for one event by name on one object.
    ///
    /// There is deliberately no `Advise` command: advising is DERIVED from
    /// this (see the command's arm in the session loop). A separate control
    /// would allow the state where a client has registered interest and the
    /// event never arrives, with nothing to show for it.
    Subscribe {
        handle: u64,
        event: String,
        args: bool,
        reply: Sender<Result<(), (String, String)>>,
    },
    /// Drop one event name. The reply says whether it had been subscribed.
    Unsubscribe {
        handle: u64,
        event: String,
        reply: Sender<bool>,
    },
    /// Release everything one event minted for its arguments. The reply is
    /// the ids that were dropped, so the caller's routing table can be kept
    /// in step without re-reading the JSON that carried them.
    ReleaseEvent {
        seq: u64,
        reply: Sender<Vec<u64>>,
    },
    /// Revoke shutdown permission for this session's instance. Any handle of
    /// the session identifies it (session <-> instance is 1:1). Sent by the
    /// `leave_open` wire method; the arm calls `registry.leave_open`.
    LeaveOpen {
        reply: Sender<()>,
    },
}

/// One event on its way out of a session.
///
/// [`SessionEvent::frame`] is what goes on the wire. `new_handles` is what the
/// session minted for that frame's object arguments, reported structurally for
/// the same reason [`crate::server::Server::handle_tracked`] reports the
/// handles an invoke minted: whoever routes these handles should not have to
/// re-parse the JSON looking for `$ole_ref`, and cannot then drift out of step
/// with how the result happens to be shaped or nested.
pub struct SessionEvent {
    pub frame: crate::protocol::Event,
    pub new_handles: Vec<u64>,
}

pub struct SessionHandle {
    pub route: SessionRoute,
    pub root_handle_id: u64,
    pub created: bool,
    /// The session's worker thread, so that a caller who wants to can wait for
    /// it to finish. Nothing in the server does -- it drops this and lets the
    /// thread run detached -- but a test that asserts the session actually
    /// tears down has no other way to observe the outcome, only the mechanism.
    pub worker: JoinHandle<()>,
}

/// A session's command channel and the waker that belongs to it, welded
/// together.
///
/// A bare `Sender<SessionCommand>` is a loaded gun now that the session thread
/// sleeps in `MsgWaitForMultipleObjects` instead of on the channel. It has two
/// ways to go off, and this type is the answer to both:
///
/// - a send that is not followed by a wake leaves the command in the channel
///   and the thread asleep (see [`SessionRoute::send`]);
/// - dropping the last sender no longer ends the loop by itself the way the
///   blocking `recv` did, so it has to wake the thread too (see the `Drop`
///   impl).
///
/// Every `Sender` for a session lives inside one of these — `SessionHandle`
/// included — so neither guarantee can be lost by a caller holding the sender
/// on its own.
#[derive(Clone)]
pub struct SessionRoute {
    // `Option` only so `Drop` can let go of the sender before it wakes; it is
    // `Some` for the entire life of a live route.
    sender: Option<Sender<SessionCommand>>,
    waker: Arc<Waker>,
}

impl SessionRoute {
    fn new(sender: Sender<SessionCommand>, waker: Arc<Waker>) -> Self {
        SessionRoute { sender: Some(sender), waker }
    }

    /// Queue a command and wake the session thread.
    ///
    /// THE ORDER IS LOAD-BEARING: push first, then wake. Waking first lets the
    /// session thread drain an empty channel and go back to sleep before the
    /// command lands, and the command then waits for an unrelated later wake --
    /// or forever. Never call `sender.send` directly.
    pub fn send(&self, cmd: SessionCommand) -> Result<(), SendError<SessionCommand>> {
        let sender = self.sender.as_ref().expect("a live SessionRoute always holds its sender");
        sender.send(cmd)?;
        self.waker.wake();
        Ok(())
    }
}

impl Drop for SessionRoute {
    /// Dropping the last route is how a session learns that its connection
    /// went away without releasing every handle. The old loop got that for
    /// free: `for cmd in cmd_rx` ended the moment the senders were gone. A
    /// thread parked in `MsgWaitForMultipleObjects` cannot see a channel
    /// disconnect at all, so without this wake it sleeps forever, never runs
    /// its cleanup, and the IDispatch proxies it still holds keep the
    /// automated application (EXCEL.EXE) alive after the bridge is done with
    /// it. Measured: without this, a session dropped rather than released
    /// leaves an EXCEL.EXE behind for good.
    ///
    /// Waking every route's drop rather than only the last one costs a
    /// spurious wake-up: the thread finds the channel merely empty and parks
    /// again.
    fn drop(&mut self) {
        // Same order as `send`, for the same reason: let go of the sender
        // first, so a thread woken by the line below sees `Disconnected`
        // rather than `Empty`.
        self.sender = None;
        self.waker.wake();
    }
}

/// Queue a command on a session and wake its thread. See [`SessionRoute::send`].
pub fn send(handle: &SessionHandle, cmd: SessionCommand) -> Result<(), SendError<SessionCommand>> {
    handle.route.send(cmd)
}

static SESSION_SEQ: AtomicU32 = AtomicU32::new(1);

enum SpawnMode {
    Create,
    Connect,
    ConnectOrCreate,
}

pub fn spawn_create(class_name: String, events: Sender<SessionEvent>, registry: Arc<InstanceRegistry>, cleanup: Option<CleanupConfig>) -> ComResult<SessionHandle> {
    spawn(class_name, SpawnMode::Create, events, registry, cleanup)
}

pub fn spawn_connect(class_name: String, events: Sender<SessionEvent>, registry: Arc<InstanceRegistry>, cleanup: Option<CleanupConfig>) -> ComResult<SessionHandle> {
    spawn(class_name, SpawnMode::Connect, events, registry, cleanup)
}

pub fn spawn_connect_or_create(class_name: String, events: Sender<SessionEvent>, registry: Arc<InstanceRegistry>, cleanup: Option<CleanupConfig>) -> ComResult<SessionHandle> {
    spawn(class_name, SpawnMode::ConnectOrCreate, events, registry, cleanup)
}

/// The session thread's apartment, and everything COM owns back.
///
/// This type exists so that the teardown ORDER is not a sequence of statements
/// somebody can shorten. Every one of these has to be let go of while the
/// apartment that created it still exists, and in this order:
///
/// 1. `advised` — each `Advised`'s `Drop` calls `Unadvise`. A sink that
///    survives is a sink Excel still holds a reference to, and EXCEL.EXE then
///    outlives the bridge.
/// 2. `raw_rx` — callbacks that arrived after the last drain still hold their
///    argument VARIANTs, and an argument VARIANT is an IDispatch.
/// 3. `handles` — every surviving proxy.
/// 4. only then `CoUninitialize`.
///
/// Written out as four statements in the session thread, deleting any one of
/// them left a green suite: this Wine tolerates a release into an apartment
/// that is already gone, so the mistake has no local symptom to test for. As a
/// `Drop` the order cannot be deleted by editing a line — it would take
/// removing a field from a struct — and the early-return paths that used to
/// repeat the sequence by hand collapse into it.
///
/// `raw_rx` is an `Option` for one reason: a struct's fields are dropped AFTER
/// its `Drop::drop` returns, so a plain `Receiver` field would release its
/// queued VARIANTs after step 4 — the exact bug this type prevents. Setting it
/// to `None` inside the body is what puts it at step 2.
struct Apartment {
    handles: HashMap<u64, IDispatch>,
    /// What is actually advised. DERIVED from the session's `subs` and never
    /// set directly: an object is advised exactly while it has at least one
    /// subscription, so there is no way to be subscribed without being advised
    /// (the event would never arrive) or advised without being subscribed
    /// (Excel would keep calling into a sink nobody reads).
    advised: HashMap<u64, sink::Advised>,
    /// `Some` for the whole life of a live apartment; see the type's doc
    /// comment for why it is an `Option` at all.
    raw_rx: Option<Receiver<sink::RawEvent>>,
}

impl Apartment {
    /// Initialise this thread's apartment and take ownership of what will have
    /// to be released inside it.
    ///
    /// Initialising HERE rather than at the call site is the other half of the
    /// guarantee: an initialised apartment and the guard that uninitialises it
    /// come into existence together, so there is no window in which one exists
    /// without the other.
    fn enter(raw_rx: Receiver<sink::RawEvent>) -> ComResult<Apartment> {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok() }.map_err(ComError::from)?;
        Ok(Apartment { handles: HashMap::new(), advised: HashMap::new(), raw_rx: Some(raw_rx) })
    }
}

impl Drop for Apartment {
    fn drop(&mut self) {
        self.advised.clear();
        self.raw_rx = None;
        self.handles.clear();
        unsafe { CoUninitialize() };
    }
}

/// Start a session.
///
/// `events` is where this session's events go, and it is an argument rather
/// than something the handle hands back afterwards on purpose: a session that
/// could be created without saying where its events go is a session whose
/// events can be silently dropped, and the whole point of deriving `Advise`
/// from a subscription is that there is no such state to reach.
fn spawn(progid: String, mode: SpawnMode, events: Sender<SessionEvent>, registry: Arc<InstanceRegistry>, cleanup: Option<CleanupConfig>) -> ComResult<SessionHandle> {
    let session_seq = SESSION_SEQ.fetch_add(1, Ordering::SeqCst) as u64;
    let (cmd_tx, cmd_rx) = channel::<SessionCommand>();
    let (ready_tx, ready_rx) = channel::<ComResult<(u64, bool)>>();
    // Created out here, before the thread exists: a caller that gets a
    // `SessionHandle` back must never be able to send a command it has no way
    // to wake, not even for the instant the thread takes to start.
    let waker = Arc::new(Waker::new().map_err(ComError::from)?);
    let thread_waker = waker.clone();

    let worker = std::thread::spawn(move || {
        // Raw callbacks from the sink. The sink must never block (it runs
        // inside a COM call from Excel), so it pushes VARIANTs here and this
        // thread does the converting, below, once it is out of that call.
        let (raw_tx, raw_rx) = channel::<sink::RawEvent>();

        // The braces below are `com`'s scope, and `Apartment`'s `Drop` is the
        // whole teardown. Every exit from the session — the early returns, the
        // final `Release`, the caller disconnecting, a failed wait — leaves
        // this block, and leaving it runs that teardown in that order. The
        // acknowledgement below is deliberately outside: a caller waiting on
        // it is waiting for the automated process to have been let go, so it
        // must not be sent until the apartment is gone.
        let shutdown_reply: Option<Sender<ReleaseOutcome>> = {
            let mut com = match Apartment::enter(raw_rx) {
                Ok(com) => com,
                Err(e) => {
                    let _ = ready_tx.send(Err(e));
                    return;
                }
            };

            let root = match mode {
                SpawnMode::Create => dispatch::create_instance(&progid).map(|d| (d, true)),
                SpawnMode::Connect => dispatch::get_active_object(&progid).map(|d| (d, false)),
                SpawnMode::ConnectOrCreate => dispatch::connect_or_create(&progid),
            };

            let mut local_seq: u32 = 0;

            // What the client asked for: object -> event name -> the `args`
            // option. This is the record of intent; `com.advised` is what is
            // derived from it.
            let mut subs: HashMap<u64, HashMap<String, bool>> = HashMap::new();
            // Handles minted for event arguments, keyed by the event's seq, so
            // that one event's arguments are released together or not at all --
            // a partial release would leave the client holding ids that no longer
            // resolve, with no way to tell which.
            let mut event_handles: HashMap<u64, Vec<u64>> = HashMap::new();

            let (root_disp, mut created) = match root {
                Ok(pair) => pair,
                Err(e) => {
                    let _ = ready_tx.send(Err(e));
                    return;
                }
            };

            let root_id = (session_seq << 32) | (local_seq as u64);
            local_seq += 1;
            // Cloned before the move into the table: the identity key and the
            // cleanup steps both need the IDispatch, and the table owns the
            // original for the rest of the session.
            let mut root_disp_for_cleanup = root_disp.clone();
            com.handles.insert(root_id, root_disp);

            // Participation: only a session that carried `cleanup` joins the
            // registry. Its key is the cross-apartment identity; if marshaling
            // fails the session falls back to a private synthetic key and is
            // treated as the sole user of its own instance (today's behavior).
            let participation: Option<(InstanceKey, CleanupConfig)> = match &cleanup {
                None => None,
                Some(cfg) => {
                    let key = instance_key(&root_disp_for_cleanup).unwrap_or_else(|e| {
                        eprintln!("wineole: session {} could not read instance identity ({}); \
                                   treating it as its own sole user", session_seq, e);
                        synthetic_key(session_seq)
                    });
                    match registry.join(key, session_seq, created, cfg) {
                        JoinResult::Joined => Some((key, cfg.clone())),
                        JoinResult::Closing => match mode {
                            // connect_or_create makes a fresh instance rather
                            // than failing when the running one is shutting
                            // down (design edge case).
                            SpawnMode::ConnectOrCreate => {
                                com.handles.remove(&root_id);
                                let fresh = match dispatch::create_instance(&progid) {
                                    Ok(d) => d,
                                    Err(e) => {
                                        let _ = ready_tx.send(Err(e));
                                        return;
                                    }
                                };
                                created = true;
                                root_disp_for_cleanup = fresh.clone();
                                com.handles.insert(root_id, fresh);
                                let key2 = instance_key(&root_disp_for_cleanup).unwrap_or_else(|e| {
                                    eprintln!("wineole: session {} could not read instance identity ({}); \
                                               treating it as its own sole user", session_seq, e);
                                    synthetic_key(session_seq)
                                });
                                match registry.join(key2, session_seq, true, cfg) {
                                    JoinResult::Joined => Some((key2, cfg.clone())),
                                    JoinResult::Closing => {
                                        let _ = ready_tx.send(Err(ComError::new(
                                            windows::Win32::Foundation::E_FAIL,
                                            "WineOLE::InstanceClosingError: freshly created instance reported closing",
                                        )));
                                        return;
                                    }
                                }
                            }
                            _ => {
                                // A connect landed on an instance being shut down.
                                let _ = ready_tx.send(Err(ComError::new(
                                    windows::Win32::Foundation::E_FAIL,
                                    "WineOLE::InstanceClosingError: the instance is shutting down; \
                                     connect again to get a fresh one",
                                )));
                                return;
                            }
                        },
                    }
                }
            };

            if ready_tx.send(Ok((root_id, created))).is_err() {
                return;
            }

            // On the release that empties `handles`, don't reply on the spot: the
            // reply channel wakes the caller, who may immediately send another
            // command, and that send only fails once `cmd_rx` is actually
            // dropped. A reply sent from inside the loop would already be in
            // flight before that — a race the caller's thread reliably wins,
            // making the "further sends fail" test assertion flaky/false.
            // Deferring the reply until after the loop, and dropping `cmd_rx`
            // explicitly on the way out, makes the shutdown ordering a real
            // happens-before relationship instead of a race: the caller can't
            // observe the reply until the receiver is already gone.
            let mut shutdown_reply: Option<Sender<ReleaseOutcome>> = None;
            // A live client's closure has been asked to run first, delivered as
            // a `$cleanup` event carrying this seq. Stays `Some` from the moment
            // that event goes out until the session tears down, so that the
            // matching `release_event` (Step 2) and the disconnect catch-all
            // (Step 2b) can both tell they are completing a pending cleanup.
            let mut pending_cleanup: Option<u64> = None;
            'outer: loop {
                match pump::wait(thread_waker.handle()) {
                    Wake::Messages => {
                        // COM event callbacks are delivered here. Without this
                        // the sink in sink.rs is advised and never called.
                        pump::drain_messages();
                    }
                    Wake::Commands => {}
                    // The wait itself broke. There is nothing to peek and nothing
                    // to receive, so going round again would be a 100%-CPU spin on
                    // a call that keeps failing. Tear the session down instead --
                    // loudly, and through the same cleanup as every other exit, so
                    // the IDispatch proxies are still released inside the
                    // apartment rather than left to keep EXCEL.EXE alive.
                    Wake::Failed(err) => {
                        eprintln!(
                            "wineole: session {} wait failed (error {}); shutting the session down",
                            session_seq, err
                        );
                        break 'outer;
                    }
                }

                // Drain the channel every time round, not only on Wake::Commands:
                // one wake can stand for several queued commands, and a message
                // wake may also have raced a send.
                loop {
                    match cmd_rx.try_recv() {
                        Ok(cmd) => match cmd {
                            SessionCommand::Invoke { handle, name, args, named, reply } => {
                                let outcome = run_invoke(&mut com.handles, &mut local_seq, session_seq, handle, &name, &args, &named);
                                let _ = reply.send(outcome);
                            }
                            SessionCommand::ConstLoad { handle, reply } => {
                                let outcome = run_const_load(&com.handles, handle);
                                let _ = reply.send(outcome);
                            }
                            // A released object's subscriptions go with it in
                            // both branches below: leaving them behind keeps an
                            // `IConnectionPoint` on an object the session no
                            // longer holds -- a live COM reference pinning
                            // EXCEL.EXE for the rest of the session -- and keeps
                            // events arriving named by a handle the client has
                            // released and the server has already unrouted.
                            //
                            // `client_alive` is not consulted here: Task 4 runs
                            // the last user's cleanup INLINE. Task 5 rewrites
                            // the `RunCleanup` branch to add the `$cleanup`
                            // client-callback path, which is where it matters.
                            SessionCommand::Release { handle, client_alive, reply } => {
                                if let (Some((key, _cfg)), true) = (participation.as_ref(), handle == root_id) {
                                    match registry.on_root_release(*key, session_seq) {
                                        LeaveDecision::RunCleanup { steps, callback } if callback && client_alive => {
                                            // Ask the live client's closure to run first. Keep the
                                            // root handle so the table stays non-empty and this
                                            // session keeps looping to receive the release_event;
                                            // the steps live on the registry record (not `steps`
                                            // here) so a leave_open before release_event is honored.
                                            let _ = steps;
                                            let seq = (session_seq << 32) | (local_seq as u64);
                                            local_seq += 1;
                                            let _ = events.send(SessionEvent {
                                                frame: crate::protocol::Event {
                                                    event: "$cleanup".to_string(),
                                                    handle: root_id,
                                                    seq,
                                                    args: None,
                                                },
                                                new_handles: vec![],
                                            });
                                            pending_cleanup = Some(seq);
                                            let _ = reply.send(ReleaseOutcome::CleanupPending(seq));
                                            continue; // do NOT tear down; wait for release_event
                                        }
                                        LeaveDecision::RunCleanup { steps, .. } => {
                                            // No callback, or the client is not alive: quit inline.
                                            run_cleanup_steps(&root_disp_for_cleanup, &steps);
                                            registry.finish_cleanup(*key);
                                        }
                                        LeaveDecision::NotLast => {}
                                    }
                                    // NotLast falls straight through: another user still holds
                                    // the instance, so this session tears down without touching
                                    // Excel. Either way, releasing the ROOT ends this session.
                                    com.handles.remove(&handle);
                                    subs.remove(&handle);
                                    com.advised.remove(&handle);
                                    shutdown_reply = Some(reply);
                                    break 'outer;
                                } else {
                                    // Non-root release, or a non-participating session:
                                    // exactly today's behavior.
                                    com.handles.remove(&handle);
                                    subs.remove(&handle);
                                    com.advised.remove(&handle);
                                    if com.handles.is_empty() {
                                        shutdown_reply = Some(reply);
                                        break 'outer;
                                    } else {
                                        let _ = reply.send(ReleaseOutcome::Done);
                                    }
                                }
                            }
                            SessionCommand::Subscribe { handle, event, args, reply } => {
                                let outcome = (|| {
                                    let disp = com.handles.get(&handle).cloned().ok_or_else(|| {
                                        (
                                            "WineOLE::StaleReferenceError".to_string(),
                                            format!("unknown handle {}", handle),
                                        )
                                    })?;
                                    // An event argument is valid for the
                                    // duration of its callback and no longer:
                                    // the client releases the whole seq when
                                    // the callback returns, and this session
                                    // drops the ids with it. A subscription on
                                    // one could therefore never be honoured --
                                    // the Advise would outlive the handle it is
                                    // named by, keeping an IConnectionPoint (and
                                    // so EXCEL.EXE) alive for the rest of the
                                    // session while events went out named by a
                                    // handle the client had released and this
                                    // side no longer routed. Refused rather than
                                    // accepted-and-cleaned-up, because there is
                                    // no moment at which accepting it would have
                                    // been right.
                                    //
                                    // Nothing needs purging in `ReleaseEvent` to
                                    // back this up: ids are unique per session
                                    // (`session_seq << 32 | local_seq`), so an
                                    // id minted for an event argument is never
                                    // also an invoke result's id -- not even for
                                    // the same COM object -- and this check is
                                    // therefore complete on its own. That holds
                                    // short of 2^32 ids minted in ONE session:
                                    // `local_seq` is a u32 incremented without a
                                    // wrap check, and a session that reached the
                                    // wrap would have `handles.insert` colliding
                                    // on its own long before this check was the
                                    // weakest thing about it.
                                    if event_handles.values().any(|ids| ids.contains(&handle)) {
                                        return Err((
                                            "ArgumentError".to_string(),
                                            format!(
                                                "handle {} is an event argument, valid only until its callback \
                                                 returns; subscribe to the object through a call instead (e.g. \
                                                 Worksheets(name))",
                                                handle
                                            ),
                                        ));
                                    }
                                    // The derivation, in one place: this object
                                    // goes from no subscriptions to some, so it
                                    // gets advised. Nothing else ever advises.
                                    if !com.advised.contains_key(&handle) {
                                        let a = sink::advise(&disp, handle, raw_tx.clone())
                                            .map_err(|e| ("WIN32OLERuntimeError".to_string(), e.to_string()))?;
                                        com.advised.insert(handle, a);
                                    }
                                    subs.entry(handle).or_default().insert(event, args);
                                    Ok(())
                                })();
                                let _ = reply.send(outcome);
                            }
                            SessionCommand::Unsubscribe { handle, event, reply } => {
                                let was = subs
                                    .get_mut(&handle)
                                    .map(|m| m.remove(&event).is_some())
                                    .unwrap_or(false);
                                // The other half of the derivation: the last name
                                // to go takes the Advise with it. Dropping the
                                // `Advised` is the Unadvise -- see sink::Advised.
                                if subs.get(&handle).map(|m| m.is_empty()).unwrap_or(false) {
                                    subs.remove(&handle);
                                    com.advised.remove(&handle);
                                }
                                let _ = reply.send(was);
                            }
                            SessionCommand::ReleaseEvent { seq, reply } => {
                                if pending_cleanup == Some(seq) {
                                    // The client's closure has answered: run the
                                    // steps (unless a leave_open in the closure
                                    // cleared `closing`), then tear down.
                                    if let Some((key, cfg)) = participation.as_ref() {
                                        if registry.confirm_cleanup(*key) {
                                            run_cleanup_steps(&root_disp_for_cleanup, &cfg.steps);
                                        }
                                    }
                                    // Return root_id in the "released" vec so
                                    // handle_release_event purges its now-dead route
                                    // (the route was kept alive by the CleanupPending
                                    // reply so this release_event could be routed here).
                                    let _ = reply.send(vec![root_id]);
                                    com.handles.remove(&root_id);
                                    subs.remove(&root_id);
                                    com.advised.remove(&root_id);
                                    break 'outer;
                                } else {
                                    // An unknown seq is not an error: a double send
                                    // costs nothing, and refusing it would make the
                                    // client's ensure-block harder than it needs to be.
                                    let released = event_handles.remove(&seq).unwrap_or_default();
                                    for id in &released {
                                        com.handles.remove(id);
                                    }
                                    let _ = reply.send(released);
                                }
                            }
                            SessionCommand::LeaveOpen { reply } => {
                                if let Some((key, _)) = participation.as_ref() {
                                    registry.leave_open(*key);
                                }
                                let _ = reply.send(());
                            }
                        },
                        Err(std::sync::mpsc::TryRecvError::Empty) => break,
                        // Every sender is gone: the connection went away without
                        // releasing its handles. Same exit as the final Release,
                        // minus anyone to acknowledge to — what matters is the
                        // cleanup below, which both paths share.
                        Err(std::sync::mpsc::TryRecvError::Disconnected) => break 'outer,
                    }
                }

                // Raw callbacks the sink pushed while `drain_messages` ran.
                // Converted here rather than in the sink for two reasons: the
                // sink must not block inside a COM call from Excel, and turning
                // an IDispatch argument into a handle needs the handle table,
                // which only this thread may touch.
                while let Some(raw) = com.raw_rx.as_ref().and_then(|rx| rx.try_recv().ok()) {
                    // The name comes from the source interface's typeinfo, read
                    // once at Advise time. An unnamed DISPID is still delivered
                    // rather than dropped: a client that knows the number can
                    // still subscribe by it.
                    let Some(name) = com.advised.get(&raw.handle).map(|a| {
                        a.names
                            .get(&raw.dispid)
                            .cloned()
                            .unwrap_or_else(|| format!("DISPID_{}", raw.dispid))
                    }) else {
                        continue;
                    };

                    // Not subscribed, not forwarded. Excel has already paid for
                    // the call by the time we get here; this is where everything
                    // after it -- the handle minting, the JSON, the socket write,
                    // and the client's release of it all -- is saved.
                    let Some(want_args) = subs.get(&raw.handle).and_then(|m| m.get(&name)).copied()
                    else {
                        continue;
                    };

                    // An event's seq is minted from the SAME per-session id space
                    // as its handles, so `seq >> 32` names the session that
                    // produced it exactly as a handle id does. That is what lets
                    // `release_event`, which carries nothing but a seq, be routed
                    // back to the right session on a connection that has several.
                    let seq = (session_seq << 32) | (local_seq as u64);
                    local_seq += 1;

                    let mut minted: Vec<u64> = Vec::new();
                    // `args: false` is the branch that decides whether object
                    // handles are minted AT ALL. Not "minted and left out of
                    // the frame": nothing is added to the handle table, so
                    // there is nothing for the client to release and nothing
                    // to leak if it never does.
                    let args = if want_args {
                        let handles = &mut com.handles;
                        let json: Vec<Value> = raw
                            .args
                            .iter()
                            .map(|v| {
                                value::variant_to_json(v, |new_disp| {
                                    let id = (session_seq << 32) | (local_seq as u64);
                                    local_seq += 1;
                                    handles.insert(id, new_disp);
                                    minted.push(id);
                                    id
                                })
                            })
                            .collect();
                        if !minted.is_empty() {
                            event_handles.insert(seq, minted.clone());
                        }
                        Some(json)
                    } else {
                        None
                    };

                    // A closed receiver means the connection that owns this
                    // session has gone. Anything minted just above is then
                    // released by the teardown below, which every exit path
                    // reaches, so there is nothing to unwind here.
                    let _ = events.send(SessionEvent {
                        frame: crate::protocol::Event { event: name, handle: raw.handle, seq, args },
                        new_handles: minted,
                    });
                }
            }

            // A $cleanup was pending and the client vanished (disconnect) rather
            // than answering release_event. The closure is abandoned; the bridge
            // still owns the shutdown (choice B), so run the steps now, on this
            // thread, while the apartment is still up. `confirm_cleanup` returns
            // false if the release_event path already ran them, so this never
            // double-fires; and false too if a closure had called leave_open.
            if pending_cleanup.is_some() {
                if let Some((key, cfg)) = participation.as_ref() {
                    if registry.confirm_cleanup(*key) {
                        run_cleanup_steps(&root_disp_for_cleanup, &cfg.steps);
                    }
                }
            }

            // The receiver is only borrowed by the loop above, where the old
            // `for cmd in cmd_rx` consumed it and dropped it on the way out. That
            // drop is what the deferred `shutdown_reply` below relies on, so it
            // has to be spelled out here rather than left to the end of the
            // closure.
            drop(cmd_rx);

            shutdown_reply
            // `com` is dropped by this closing brace, and its `Drop` is the
            // teardown: every Advise taken back out, then the undrained
            // callbacks, then every surviving proxy, then `CoUninitialize` --
            // in that order, inside the apartment. See `Apartment`.
        };

        if let Some(reply) = shutdown_reply {
            let _ = reply.send(ReleaseOutcome::Done);
        }
    });

    let (root_id, created) = ready_rx
        .recv()
        .map_err(|_| {
            ComError::new(
                windows::Win32::Foundation::E_FAIL,
                "session worker thread died before reporting readiness",
            )
        })??;
    Ok(SessionHandle {
        route: SessionRoute::new(cmd_tx, waker),
        root_handle_id: root_id,
        created,
        worker,
    })
}

fn run_invoke(
    handles: &mut HashMap<u64, IDispatch>,
    local_seq: &mut u32,
    session_seq: u64,
    handle: u64,
    name: &str,
    args: &[Value],
    named: &HashMap<String, Value>,
) -> Result<(Value, Vec<u64>), (String, String)> {
    let disp = handles
        .get(&handle)
        .cloned()
        .ok_or_else(|| ("WineOLE::StaleReferenceError".to_string(), format!("unknown handle {}", handle)))?;

    // `ComError`'s Display always renders the HRESULT, so even the (common
    // under Wine) case of an empty OS message still tells the client which
    // HRESULT failed instead of arriving as `WIN32OLERuntimeError: ""`.
    let positional = args
        .iter()
        .map(|v| value::json_to_variant(v, |id| resolve(handles, id)))
        .collect::<ComResult<Vec<_>>>()
        .map_err(|e| ("WIN32OLERuntimeError".to_string(), e.to_string()))?;

    let named_pairs = named
        .iter()
        .map(|(k, v)| value::json_to_variant(v, |id| resolve(handles, id)).map(|variant| (k.clone(), variant)))
        .collect::<ComResult<Vec<_>>>()
        .map_err(|e| ("WIN32OLERuntimeError".to_string(), e.to_string()))?;

    let result_variant = dispatch::invoke_member(&disp, name, positional, named_pairs)
        .map_err(|e| ("WIN32OLERuntimeError".to_string(), e.to_string()))?;

    let mut new_ids = Vec::new();
    let json_result = value::variant_to_json(&result_variant, |new_disp| {
        let id = (session_seq << 32) | (*local_seq as u64);
        *local_seq += 1;
        handles.insert(id, new_disp);
        new_ids.push(id);
        id
    });

    Ok((json_result, new_ids))
}

fn run_const_load(handles: &HashMap<u64, IDispatch>, handle: u64) -> Result<Value, (String, String)> {
    let disp = handles
        .get(&handle)
        .cloned()
        .ok_or_else(|| ("WineOLE::StaleReferenceError".to_string(), format!("unknown handle {}", handle)))?;

    let pairs = dispatch::const_load(&disp).map_err(|e| ("WIN32OLERuntimeError".to_string(), e.to_string()))?;

    let mut map = serde_json::Map::new();
    for (name, variant) in pairs {
        // A constant is never an object reference — this closure existing to
        // satisfy variant_to_json's signature and panicking if it's ever
        // actually called is intentional: it would mean a type library
        // exposed an enum member holding a live COM object, which is not
        // meaningful and should surface loudly during development rather
        // than silently minting an orphan handle nobody will release.
        let value = value::variant_to_json(&variant, |_| {
            panic!("const_load: enum constant '{}' unexpectedly held an object reference", name)
        });
        map.insert(name, value);
    }
    Ok(Value::Object(map))
}

fn resolve(handles: &HashMap<u64, IDispatch>, id: u64) -> windows::core::Result<IDispatch> {
    handles
        .get(&id)
        .cloned()
        .ok_or_else(|| windows::core::Error::from_hresult(windows::Win32::Foundation::E_INVALIDARG))
}

/// Run the registered cleanup steps against the root, in order, on this
/// session thread while the apartment still exists.
///
/// A failing step is logged and the rest still run: `DisplayAlerts=` failing
/// must not stop `Quit`. No timeout -- a step is an ordinary synchronous COM
/// call, handled as `invoke` is. Object references are impossible here (the
/// args were validated to scalars), so `resolve` is never needed.
fn run_cleanup_steps(root: &IDispatch, steps: &[crate::registry::Step]) {
    for step in steps {
        let positional: Vec<_> = match step
            .args
            .iter()
            .map(|v| value::json_to_variant(v, |_| unreachable!("cleanup args are scalars")))
            .collect::<ComResult<Vec<_>>>()
        {
            Ok(v) => v,
            Err(e) => {
                eprintln!("[cleanup] step {:?} arg conversion failed: {}", step.name, e);
                continue;
            }
        };
        if let Err(e) = dispatch::invoke_member(root, &step.name, positional, vec![]) {
            eprintln!("[cleanup] step {:?} failed: {}", step.name, e);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A registry a session can join. The command-loop tests below pass `None`
    /// cleanup (non-participating), so for them it only has to exist to satisfy
    /// the spawn signature; the instance-lifetime integration tests keep a
    /// clone of it and assert against `record_summary()`.
    #[cfg(test)]
    fn fresh_registry() -> Arc<InstanceRegistry> {
        Arc::new(InstanceRegistry::new())
    }

    #[test]
    fn test_create_invoke_named_arg_and_release_shuts_down_session() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();

        let (ev_tx, _events) = channel::<SessionEvent>();
        let session = spawn_create("Excel.Application".to_string(), ev_tx, fresh_registry(), None).expect("spawn_create");

        let (tx, rx) = channel();
        send(&session, SessionCommand::Invoke {
            handle: session.root_handle_id,
            name: "Visible=".to_string(),
            args: vec![serde_json::json!(false)],
            named: HashMap::new(),
            reply: tx,
        }).unwrap();
        rx.recv().unwrap().expect("Visible=false should succeed");

        let (tx, rx) = channel();
        send(&session, SessionCommand::Invoke {
            handle: session.root_handle_id,
            name: "Workbooks".to_string(),
            args: vec![],
            named: HashMap::new(),
            reply: tx,
        }).unwrap();
        let (workbooks_json, new_ids) = rx.recv().unwrap().expect("Workbooks should succeed");
        assert_eq!(new_ids.len(), 1);
        let workbooks_handle = workbooks_json["$ole_ref"].as_u64().unwrap();

        let (tx, rx) = channel();
        send(&session, SessionCommand::Invoke {
            handle: workbooks_handle,
            name: "Add".to_string(),
            args: vec![],
            named: HashMap::new(),
            reply: tx,
        }).unwrap();
        // Workbooks.Add returns the newly created Workbook, which run_invoke
        // registers as its own handle (new_ids.len() == 1) since the return
        // value is VT_DISPATCH. That handle must be released too, or
        // `handles` never becomes empty and the session never shuts down.
        let (_add_json, add_new_ids) = rx.recv().unwrap().expect("Workbooks.Add should succeed");
        assert_eq!(add_new_ids.len(), 1);
        let new_workbook_handle = add_new_ids[0];

        let (tx, rx) = channel();
        send(&session, SessionCommand::Invoke {
            handle: session.root_handle_id,
            name: "Quit".to_string(),
            args: vec![],
            named: HashMap::new(),
            reply: tx,
        }).unwrap();
        rx.recv().unwrap().expect("Quit should succeed");

        // Releasing every outstanding handle should let the worker thread exit.
        let (tx, rx) = channel();
        send(&session, SessionCommand::Release { handle: new_workbook_handle, client_alive: true, reply: tx }).unwrap();
        rx.recv().unwrap();

        let (tx, rx) = channel();
        send(&session, SessionCommand::Release { handle: workbooks_handle, client_alive: true, reply: tx }).unwrap();
        rx.recv().unwrap();

        let (tx, rx) = channel();
        send(&session, SessionCommand::Release { handle: session.root_handle_id, client_alive: true, reply: tx }).unwrap();
        rx.recv().unwrap();

        // The worker thread has exited; further sends fail because the receiver is dropped.
        let (tx, _rx) = channel();
        let send_result = send(&session, SessionCommand::Release { handle: session.root_handle_id, client_alive: true, reply: tx });
        assert!(send_result.is_err(), "worker thread should have exited after last handle released");
    }

    #[test]
    fn test_const_load_command_returns_excel_constants() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();

        let (ev_tx, _events) = channel::<SessionEvent>();
        let session = spawn_create("Excel.Application".to_string(), ev_tx, fresh_registry(), None).expect("spawn_create");

        let (tx, rx) = channel();
        send(&session, SessionCommand::ConstLoad { handle: session.root_handle_id, reply: tx })
            .unwrap();
        let constants = rx.recv().unwrap().expect("const_load should succeed");
        let obj = constants.as_object().expect("const_load result must be a JSON object");
        assert_eq!(obj.get("xlUp").and_then(|v| v.as_i64()), Some(-4162));

        let (tx, rx) = channel();
        send(&session, SessionCommand::Invoke {
            handle: session.root_handle_id,
            name: "Quit".to_string(),
            args: vec![],
            named: HashMap::new(),
            reply: tx,
        }).unwrap();
        rx.recv().unwrap().expect("Quit should succeed");

        let (tx, rx) = channel();
        send(&session, SessionCommand::Release { handle: session.root_handle_id, client_alive: true, reply: tx }).unwrap();
        rx.recv().unwrap();
    }

    #[test]
    fn test_spawn_create_reports_created_true() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();

        let (ev_tx, _events) = channel::<SessionEvent>();
        let session = spawn_create("Excel.Application".to_string(), ev_tx, fresh_registry(), None).expect("spawn_create");
        assert!(session.created, "spawn_create must report created=true");

        let (tx, rx) = channel();
        send(&session, SessionCommand::Invoke {
            handle: session.root_handle_id,
            name: "Quit".to_string(),
            args: vec![],
            named: HashMap::new(),
            reply: tx,
        }).unwrap();
        rx.recv().unwrap().expect("Quit should succeed");

        let (tx, rx) = channel();
        send(&session, SessionCommand::Release { handle: session.root_handle_id, client_alive: true, reply: tx }).unwrap();
        rx.recv().unwrap();
    }

    #[test]
    fn test_spawn_connect_or_create_attaches_to_an_existing_session() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();

        let (creator_ev_tx, _creator_events) = channel::<SessionEvent>();
        let creator = spawn_create("Excel.Application".to_string(), creator_ev_tx, fresh_registry(), None).expect("spawn_create");

        let (ev_tx, _events) = channel::<SessionEvent>();
        let session = spawn_connect_or_create("Excel.Application".to_string(), ev_tx, fresh_registry(), None)
            .expect("spawn_connect_or_create");
        assert!(!session.created, "spawn_connect_or_create must report created=false when something was already running");

        let (tx, rx) = channel();
        send(&session, SessionCommand::Invoke {
            handle: session.root_handle_id,
            name: "Quit".to_string(),
            args: vec![],
            named: HashMap::new(),
            reply: tx,
        }).unwrap();
        rx.recv().unwrap().expect("Quit should succeed");

        let (tx, rx) = channel();
        send(&session, SessionCommand::Release { handle: session.root_handle_id, client_alive: true, reply: tx }).unwrap();
        rx.recv().unwrap();

        let (tx, rx) = channel();
        send(&creator, SessionCommand::Release { handle: creator.root_handle_id, client_alive: true, reply: tx }).unwrap();
        rx.recv().unwrap();
    }
    #[test]
    fn test_dropping_a_route_wakes_the_session_thread() {
        // The mechanism, in isolation and without COM: `SessionRoute::drop`
        // signals the waker. What that mechanism is *for* is the test below;
        // this one just pins that the signal happens at all, and that it does
        // not happen without a drop.
        //
        // The two `try_recv` assertions this test used to make are gone: both
        // held for any `Drop` impl at all, including an empty one, because
        // Rust drops the `Sender` field either way.
        let waker = Arc::new(Waker::new().expect("Waker::new"));
        let (cmd_tx, cmd_rx) = channel::<SessionCommand>();
        let route = SessionRoute::new(cmd_tx, waker.clone());

        // A fresh Waker is created unsignalled, so any wake below came from
        // the drop and from nothing else. A zero timeout is enough to say so:
        // a signalled event returns immediately.
        assert!(
            pump::wait_timeout(waker.handle(), 0).is_none(),
            "nothing has been sent and nothing dropped, so nothing has woken it"
        );

        drop(route);
        assert!(
            matches!(pump::wait_timeout(waker.handle(), 1000), Some(Wake::Commands)),
            "dropping a route must wake the session thread"
        );

        // The `Drop` impl lets go of the sender before it wakes, so a thread
        // woken by that signal sees Disconnected rather than Empty and exits
        // instead of parking again. A single-threaded test cannot observe the
        // moment between those two lines; the outcome test below observes the
        // consequence instead.
        drop(cmd_rx);
    }

    #[test]
    fn test_the_session_thread_exits_when_its_last_route_is_dropped() {
        // The outcome, not the mechanism: a connection that goes away without
        // releasing its handles must still end the session.
        //
        // This is the path a dropped (rather than closed) client takes. The
        // old `for cmd in cmd_rx` loop ended by itself when the last sender
        // went; a thread parked in MsgWaitForMultipleObjects cannot see a
        // channel disconnect at all, so without both halves of the fix -- the
        // wake in `SessionRoute::drop` and the `break 'outer` the woken loop
        // takes on Disconnected -- this thread sleeps forever, never reaches
        // `handles.clear()`, and its IDispatch proxies keep EXCEL.EXE alive
        // for good. Measured: 2 leftover EXCEL.EXE across the suite.
        //
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();

        let (ev_tx, _events) = channel::<SessionEvent>();
        let session = spawn_create("Excel.Application".to_string(), ev_tx, fresh_registry(), None).expect("spawn_create");

        // Take the worker apart from the route: dropping the whole
        // `SessionHandle` would drop the JoinHandle too and detach the thread,
        // leaving nothing to observe.
        let SessionHandle { route, worker, .. } = session;

        // Deliberately no Release, and deliberately no Quit either: this is
        // the disconnect path, and the point is that the worker still runs its
        // cleanup. Measured running alone: 0 leftover EXCEL.EXE, because
        // `handles.clear()` drops the last proxy and Excel exits on its own.
        drop(route);

        // `JoinHandle::join` has no timeout, and the failure being guarded
        // against is precisely a thread that never ends -- so join on a helper
        // thread and bound the wait with `recv_timeout`, which blocks rather
        // than polls.
        let (done_tx, done_rx) = channel::<()>();
        std::thread::spawn(move || {
            let _ = worker.join();
            let _ = done_tx.send(());
        });
        assert!(
            done_rx.recv_timeout(std::time::Duration::from_secs(30)).is_ok(),
            "dropping the last route must end the session worker thread; it is \
             still parked, holding the COM proxies that keep EXCEL.EXE alive"
        );
    }

    /// One invoke on a session, unwrapped. Every event test below walks a
    /// chain of Excel objects to reach a cell, and spelling the channel out
    /// eight times would bury what each step is for.
    #[cfg(test)]
    fn invoke_for_test(
        session: &SessionHandle,
        handle: u64,
        name: &str,
        args: Vec<Value>,
    ) -> (Value, Vec<u64>) {
        let (tx, rx) = channel();
        send(session, SessionCommand::Invoke {
            handle,
            name: name.to_string(),
            args,
            named: HashMap::new(),
            reply: tx,
        })
        .expect("the session must still be running");
        rx.recv().unwrap().unwrap_or_else(|e| panic!("{} failed: {:?}", name, e))
    }

    /// The object handle an invoke returned.
    #[cfg(test)]
    fn ref_of(value: &Value) -> u64 {
        value["$ole_ref"].as_u64().unwrap_or_else(|| panic!("expected an object reference, got {}", value))
    }

    /// Tear a session down the way a dropped connection does, and wait for it
    /// to finish. Bounded, because the failure being guarded against is a
    /// thread that never ends -- and an unbounded `join` on one hangs the
    /// whole suite rather than failing it.
    #[cfg(test)]
    fn drop_and_join_for_test(session: SessionHandle) {
        let SessionHandle { route, worker, .. } = session;
        drop(route);
        let (done_tx, done_rx) = channel::<()>();
        std::thread::spawn(move || {
            let _ = worker.join();
            let _ = done_tx.send(());
        });
        assert!(
            done_rx.recv_timeout(std::time::Duration::from_secs(30)).is_ok(),
            "the session worker must exit when its last route is dropped; it is still \
             parked, holding the COM proxies that keep EXCEL.EXE alive"
        );
    }

    /// The whole of Task 4's path, on a live Excel: subscribing ADVISES (the
    /// event could not arrive otherwise), the raw callback becomes a frame,
    /// the DISPID becomes a name, the object arguments become handles, and
    /// unsubscribing stops the lot.
    ///
    /// Asserting only that `Subscribe` returned `Ok` would pin nothing at
    /// all: an implementation that records the name and never advises answers
    /// exactly the same way, and the client is then registered for an event
    /// that never comes. Receiving one is the only proof.
    #[test]
    fn test_subscribing_advises_and_the_event_arrives_as_a_named_frame() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();

        let (ev_tx, events) = channel::<SessionEvent>();
        let session = spawn_create("Excel.Application".to_string(), ev_tx, fresh_registry(), None).expect("spawn_create");
        let root = session.root_handle_id;

        invoke_for_test(&session, root, "Visible=", vec![serde_json::json!(false)]);
        // Without this, `Quit` on a workbook with unsaved changes puts up a
        // save prompt no one can answer and leaves EXCEL.EXE behind for every
        // later test to trip over.
        invoke_for_test(&session, root, "DisplayAlerts=", vec![serde_json::json!(false)]);

        let (tx, rx) = channel();
        send(&session, SessionCommand::Subscribe {
            handle: root,
            event: "SheetChange".to_string(),
            args: true,
            reply: tx,
        })
        .unwrap();
        rx.recv().unwrap().expect("subscribe should succeed on Excel.Application");

        let (books, _) = invoke_for_test(&session, root, "Workbooks", vec![]);
        let (book, _) = invoke_for_test(&session, ref_of(&books), "Add", vec![]);
        let (sheet, _) = invoke_for_test(&session, ref_of(&book), "ActiveSheet", vec![]);
        let (cell, _) =
            invoke_for_test(&session, ref_of(&sheet), "Range", vec![serde_json::json!("A1")]);
        // The change that raises SheetChange.
        invoke_for_test(&session, ref_of(&cell), "Value=", vec![serde_json::json!(42)]);

        let ev = events
            .recv_timeout(std::time::Duration::from_secs(30))
            .expect("no event arrived within 30s: subscribing did not advise, or the frame \
                     never left the session");

        assert_eq!(
            ev.frame.event, "SheetChange",
            "the DISPID must be resolved to the name the client subscribed by"
        );
        assert_eq!(ev.frame.handle, root, "the frame names the object the event came from");
        assert_eq!(
            ev.frame.seq >> 32,
            root >> 32,
            "an event seq is minted from the same per-session id space as a handle, which is \
             what lets release_event be routed back to the right session"
        );

        // Two arguments, both objects, both minted as handles -- and reported
        // structurally, not left for someone to find by re-reading the JSON.
        let args = ev.frame.args.clone().expect("args: true means the arguments are there");
        assert_eq!(args.len(), 2, "SheetChange(Sh, Target) has two arguments, got {:?}", args);
        let arg_ids: Vec<u64> = args.iter().map(ref_of).collect();
        assert_eq!(arg_ids, ev.new_handles, "every minted handle must be reported");

        // The handles are live: the session can still be asked about them.
        let (address, _) = invoke_for_test(&session, arg_ids[1], "Address", vec![]);
        assert!(
            address.as_str().unwrap_or_default().contains("$A$1"),
            "SheetChange's second argument is the changed range; got {:?}",
            address
        );

        // One event's arguments go together. Releasing the seq drops both.
        let (tx, rx) = channel();
        send(&session, SessionCommand::ReleaseEvent { seq: ev.frame.seq, reply: tx }).unwrap();
        let released = rx.recv().unwrap();
        assert_eq!(released, arg_ids, "release_event releases the whole seq, not part of it");

        let (tx, rx) = channel();
        send(&session, SessionCommand::Invoke {
            handle: arg_ids[1],
            name: "Address".to_string(),
            args: vec![],
            named: HashMap::new(),
            reply: tx,
        })
        .unwrap();
        let after = rx.recv().unwrap();
        assert_eq!(
            after.unwrap_err().0,
            "WineOLE::StaleReferenceError",
            "a released event argument must no longer resolve"
        );

        // Anything Excel raised while we were busy above, so that the silence
        // asserted after the unsubscribe is about the unsubscribe.
        while events.try_recv().is_ok() {}

        let (tx, rx) = channel();
        send(&session, SessionCommand::Unsubscribe {
            handle: root,
            event: "SheetChange".to_string(),
            reply: tx,
        })
        .unwrap();
        assert!(rx.recv().unwrap(), "unsubscribe reports whether it was subscribed");

        // The same change again. Nothing may come of it: the last name for
        // this object went, so the Advise went with it.
        invoke_for_test(&session, ref_of(&cell), "Value=", vec![serde_json::json!(43)]);
        assert!(
            events.recv_timeout(std::time::Duration::from_secs(5)).is_err(),
            "unsubscribing the only name must stop the events"
        );

        invoke_for_test(&session, root, "Quit", vec![]);
        drop_and_join_for_test(session);
    }

    /// An event argument cannot be subscribed to, and is refused rather than
    /// accepted.
    ///
    /// Reachable from shipped Ruby with nothing exotic:
    /// `sheet.ole_events.on(...)` on the sheet a `SheetChange` callback was
    /// handed. Accepting it advises an object whose handle is released the
    /// moment that callback returns, and the `Advised` outlives it -- a live
    /// `IConnectionPoint` pinning EXCEL.EXE for the rest of the session,
    /// still sending events named by a handle the client has released and
    /// this side no longer routes.
    ///
    /// The second half is not decoration: it is what makes the first half
    /// mean something. Without it, a `subscribe` that was simply broken --
    /// one that refused everything -- would pass this test.
    #[test]
    fn test_subscribing_to_an_event_argument_is_refused() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();

        let (ev_tx, events) = channel::<SessionEvent>();
        let session = spawn_create("Excel.Application".to_string(), ev_tx, fresh_registry(), None).expect("spawn_create");
        let root = session.root_handle_id;

        invoke_for_test(&session, root, "Visible=", vec![serde_json::json!(false)]);
        invoke_for_test(&session, root, "DisplayAlerts=", vec![serde_json::json!(false)]);

        let (tx, rx) = channel();
        send(&session, SessionCommand::Subscribe {
            handle: root,
            event: "SheetChange".to_string(),
            args: true,
            reply: tx,
        })
        .unwrap();
        rx.recv().unwrap().expect("subscribe should succeed on Excel.Application");

        let (books, _) = invoke_for_test(&session, root, "Workbooks", vec![]);
        let (book, _) = invoke_for_test(&session, ref_of(&books), "Add", vec![]);
        let (sheet, _) = invoke_for_test(&session, ref_of(&book), "ActiveSheet", vec![]);
        let (cell, _) =
            invoke_for_test(&session, ref_of(&sheet), "Range", vec![serde_json::json!("A1")]);
        invoke_for_test(&session, ref_of(&cell), "Value=", vec![serde_json::json!(42)]);

        let ev = events
            .recv_timeout(std::time::Duration::from_secs(30))
            .expect("no event arrived within 30s: there is no event argument to subscribe to");
        assert!(
            !ev.new_handles.is_empty(),
            "SheetChange(Sh, Target) with args: true mints handles; without one this test \
             asserts nothing"
        );

        // The Worksheet the event handed us. It IS an event source, so the
        // refusal below is about how the handle was come by and nothing else.
        let (tx, rx) = channel();
        send(&session, SessionCommand::Subscribe {
            handle: ev.new_handles[0],
            event: "Change".to_string(),
            args: true,
            reply: tx,
        })
        .unwrap();
        let (class, message) = rx.recv().unwrap().expect_err(
            "an event argument is valid only until its callback returns, so a subscription on \
             one could never be honoured: accepting it leaves an Advise on a handle the client \
             has already given back",
        );
        assert_eq!(class, "ArgumentError");
        assert!(
            message.contains("is an event argument"),
            "the refusal must say what was wrong with the handle, or the caller is told \
             nothing usable; got {:?}",
            message
        );

        // The same sheet, reached through a call instead, is subscribable --
        // which is exactly what the refusal above tells the caller to do.
        let (worksheet, _) =
            invoke_for_test(&session, ref_of(&book), "Worksheets", vec![serde_json::json!(1)]);
        let (tx, rx) = channel();
        send(&session, SessionCommand::Subscribe {
            handle: ref_of(&worksheet),
            event: "Change".to_string(),
            args: true,
            reply: tx,
        })
        .unwrap();
        rx.recv()
            .unwrap()
            .expect("a Worksheet reached through a call is an ordinary handle and an event source");

        invoke_for_test(&session, root, "Quit", vec![]);
        drop_and_join_for_test(session);
    }

    /// `subscribe` fails on an object that is not an event source, rather
    /// than accepting a subscription that could never fire.
    ///
    /// The Workbook half is not decoration: it is what makes the Range half
    /// mean something. Without it, a `subscribe` that was simply broken --
    /// one that refused everything -- would pass this test.
    #[test]
    fn test_subscribing_to_a_non_event_source_fails_rather_than_pretending() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();

        let (ev_tx, _events) = channel::<SessionEvent>();
        let session = spawn_create("Excel.Application".to_string(), ev_tx, fresh_registry(), None).expect("spawn_create");
        let root = session.root_handle_id;

        invoke_for_test(&session, root, "Visible=", vec![serde_json::json!(false)]);
        invoke_for_test(&session, root, "DisplayAlerts=", vec![serde_json::json!(false)]);

        let (books, _) = invoke_for_test(&session, root, "Workbooks", vec![]);
        let (book, _) = invoke_for_test(&session, ref_of(&books), "Add", vec![]);

        // A Workbook IS an event source, so this must succeed -- which is
        // what shows the failure below is about the object and not about
        // subscribe being broken.
        let (tx, rx) = channel();
        send(&session, SessionCommand::Subscribe {
            handle: ref_of(&book),
            event: "BeforeClose".to_string(),
            args: true,
            reply: tx,
        })
        .unwrap();
        rx.recv().unwrap().expect("a Workbook is an event source");

        // The Workbooks collection is not. MEASURED on this Excel under this
        // Wine: the collections and the formatting objects -- Workbooks,
        // Worksheets, Font, Interior -- report IID_IUnknown as their
        // connection interface and then refuse FindConnectionPoint with
        // CONNECT_E_NOCONNECTION (0x80040200), which is the honest failure
        // this half of the test is about. A Worksheet is a real event source
        // (00024411, nine names), as is a Workbook (00024412, 28) and the
        // Application (00024413, 29).
        let (tx, rx) = channel();
        send(&session, SessionCommand::Subscribe {
            handle: ref_of(&books),
            event: "Change".to_string(),
            args: true,
            reply: tx,
        })
        .unwrap();
        let (class, message) = rx
            .recv()
            .unwrap()
            .expect_err("a Workbooks collection has no events; accepting this would register a \
                         callback that can never fire");
        assert_eq!(class, "WIN32OLERuntimeError");
        assert!(
            message.contains("(0x"),
            "the error must name the HRESULT, or the client is told nothing at all; got {:?}",
            message
        );

        // A Range is the other kind of non-source, and the dangerous one. It
        // offers exactly ONE connection point and it is IPropertyNotifySink
        // (9BFBBC02-EFF1-101A-84ED-00AA00341D07, measured), so `source_iid`
        // takes it -- correctly, since it is the only one -- and this
        // subscribe used to SUCCEED, reporting zero event names because that
        // IID names nothing in Excel's type library. It is a vtable
        // interface, not a dispinterface: slot 3 is `OnChanged(DISPID)` where
        // the sink's IDispatch vtable has `GetTypeInfoCount(*mut u32)`, so the
        // first property notification would write a zero through a DISPID.
        // `sink::advise` refuses it by IID now, and this is that refusal on
        // the real object rather than on a fake.
        let (sheet, _) = invoke_for_test(&session, ref_of(&book), "ActiveSheet", vec![]);
        let (cell, _) =
            invoke_for_test(&session, ref_of(&sheet), "Range", vec![serde_json::json!("A1")]);
        let (tx, rx) = channel();
        send(&session, SessionCommand::Subscribe {
            handle: ref_of(&cell),
            event: "Change".to_string(),
            args: true,
            reply: tx,
        })
        .unwrap();
        let (class, message) = rx.recv().unwrap().expect_err(
            "a Range's only connection point is IPropertyNotifySink, which this sink cannot \
             implement: accepting it hands Excel an IDispatch vtable to call OnChanged through",
        );
        assert_eq!(class, "WIN32OLERuntimeError");
        assert!(
            message.contains("IPropertyNotifySink"),
            "the refusal must name the interface it refused; got {:?}",
            message
        );

        // The Workbook is still subscribed, deliberately: this session tears
        // down with a live Advise, which is the path that has to take it back
        // out before CoUninitialize.
        invoke_for_test(&session, root, "Quit", vec![]);
        drop_and_join_for_test(session);
    }

    /// Releasing an object takes its subscriptions with it -- both halves.
    ///
    /// `Release` used to remove the id from `handles` and leave `subs` and
    /// `advised` alone, which is two failures at once: the `Advised` keeps an
    /// `IConnectionPoint` on an object the session no longer holds, a live COM
    /// reference pinning EXCEL.EXE for the rest of the session, and Excel
    /// keeps calling the sink, so events keep going out named by a handle the
    /// client has released and the server has already unrouted.
    ///
    /// A Worksheet rather than the Application, because releasing the root
    /// empties the handle table and ends the session -- which would tear the
    /// subscription down for an entirely different reason and prove nothing.
    #[test]
    fn test_releasing_a_handle_drops_its_subscription_and_stops_its_events() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();

        let (ev_tx, events) = channel::<SessionEvent>();
        let session = spawn_create("Excel.Application".to_string(), ev_tx, fresh_registry(), None).expect("spawn_create");
        let root = session.root_handle_id;

        invoke_for_test(&session, root, "Visible=", vec![serde_json::json!(false)]);
        invoke_for_test(&session, root, "DisplayAlerts=", vec![serde_json::json!(false)]);

        let (books, _) = invoke_for_test(&session, root, "Workbooks", vec![]);
        let (book, _) = invoke_for_test(&session, ref_of(&books), "Add", vec![]);
        let (sheet, _) = invoke_for_test(&session, ref_of(&book), "ActiveSheet", vec![]);
        let sheet_handle = ref_of(&sheet);
        let (cell, _) =
            invoke_for_test(&session, sheet_handle, "Range", vec![serde_json::json!("A1")]);

        let (tx, rx) = channel();
        send(&session, SessionCommand::Subscribe {
            handle: sheet_handle,
            event: "Change".to_string(),
            args: true,
            reply: tx,
        })
        .unwrap();
        rx.recv().unwrap().expect("a Worksheet is an event source");

        // It really is subscribed and really does deliver, so the silence
        // asserted below is the release doing something rather than the
        // subscription never having worked.
        invoke_for_test(&session, ref_of(&cell), "Value=", vec![serde_json::json!(42)]);
        let ev = events
            .recv_timeout(std::time::Duration::from_secs(30))
            .expect("no event arrived within 30s: subscribing to a Worksheet did not advise");
        assert_eq!(ev.frame.event, "Change");
        assert_eq!(ev.frame.handle, sheet_handle);

        // The client releases the sheet. Its own `Range` handle stays, so
        // Excel still has every reason to raise Worksheet.Change.
        let (tx, rx) = channel();
        send(&session, SessionCommand::Release { handle: sheet_handle, client_alive: true, reply: tx }).unwrap();
        rx.recv().expect("the session must still be running: other handles remain");

        // Anything already in flight, so the silence below is about the release.
        while events.try_recv().is_ok() {}

        invoke_for_test(&session, ref_of(&cell), "Value=", vec![serde_json::json!(43)]);
        assert!(
            events.recv_timeout(std::time::Duration::from_secs(5)).is_err(),
            "releasing an object must take its Advise with it: an event for a released handle \
             names something the client has thrown away and the server no longer routes"
        );

        // ...and the intent is gone too, not merely disconnected: a handle
        // with no subscriptions reports `false` here.
        let (tx, rx) = channel();
        send(&session, SessionCommand::Unsubscribe {
            handle: sheet_handle,
            event: "Change".to_string(),
            reply: tx,
        })
        .unwrap();
        assert!(
            !rx.recv().unwrap(),
            "a released handle has no subscriptions left to remove; a `true` here is a `subs` \
             entry outliving the object it names"
        );

        invoke_for_test(&session, root, "Quit", vec![]);
        drop_and_join_for_test(session);
    }

    // ---- Instance-lifetime integration tests (Task 4) -------------------
    //
    // These drive the `Server` end-to-end against a real Excel, because the
    // root-release ownership decision needs a client that RELEASES the root
    // (the `release` wire method), and the whole point is what happens to the
    // instance afterwards.
    //
    // Liveness is checked over COM (`GetActiveObject`), NOT by polling the
    // Linux process table: the cargo test runner here is `wine`, so these are
    // Windows binaries running inside Wine and a `pgrep` child process finds
    // nothing (a Windows process cannot see Linux PIDs). This is the same
    // in-Wine mechanism identity.rs's real-Excel test uses. Co-use is proved
    // structurally from the registry, which is cross-apartment identity keyed.

    /// Whether an Excel is currently registered as the running (active)
    /// object -- the in-Wine equivalent of "an EXCEL.EXE is up". This is
    /// exactly what `connect` attaches to, so its success means a live
    /// instance a client could reach. A short apartment is entered per call
    /// and balanced, so this is independent of any session thread's apartment.
    #[cfg(test)]
    fn excel_alive() -> bool {
        unsafe {
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            let alive = crate::dispatch::get_active_object("Excel.Application").is_ok();
            // S_OK and S_FALSE both need a balancing uninit; a changed-mode
            // error (already initialised differently) leaves nothing to undo.
            if hr.is_ok() {
                CoUninitialize();
            }
            alive
        }
    }

    /// Poll until no Excel is registered, up to `secs`. Returns whether it
    /// went away in time.
    #[cfg(test)]
    fn wait_until_excel_gone(secs: u64) -> bool {
        use std::time::{Duration, Instant};
        let deadline = Instant::now() + Duration::from_secs(secs);
        loop {
            if !excel_alive() {
                return true;
            }
            if Instant::now() > deadline {
                return false;
            }
            std::thread::sleep(Duration::from_millis(200));
        }
    }

    /// Best-effort teardown from inside Wine: attach to any lingering Excel and
    /// quit it (with prompts suppressed), by COM rather than by PID. A no-op
    /// once none is registered. This is what ends an Excel a `leave_open` test
    /// deliberately kept alive, and a safety net for a test that failed before
    /// its own cleanup ran.
    #[cfg(test)]
    fn quit_any_running_excel() {
        unsafe {
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            for _ in 0..10 {
                match crate::dispatch::get_active_object("Excel.Application") {
                    Ok(disp) => {
                        let _ = dispatch::invoke_member(
                            &disp, "DisplayAlerts=", vec![windows::core::VARIANT::from(false)], vec![],
                        );
                        let _ = dispatch::invoke_member(&disp, "Quit", vec![], vec![]);
                        drop(disp);
                        std::thread::sleep(std::time::Duration::from_millis(300));
                    }
                    Err(_) => break,
                }
            }
            if hr.is_ok() {
                CoUninitialize();
            }
        }
    }

    /// RAII teardown for the real-Excel integration tests: dropped at the end
    /// of the test's scope, whether that end is a normal return or a panicking
    /// assert unwinding through it. Without this, a failing assert skips the
    /// explicit teardown call and orphans a machine-global EXCEL.EXE into
    /// later tests (liveness is ROT-global, not per-test).
    struct ExcelTeardownGuard;
    impl Drop for ExcelTeardownGuard {
        fn drop(&mut self) {
            // Runs on normal exit AND on panic-unwind, so a failing assert
            // cannot orphan a machine-global EXCEL.EXE into later tests.
            quit_any_running_excel();
        }
    }

    /// One RPC against a `Server`, the way a connection makes it. Handles that
    /// the request mints are appended to `owned`.
    #[cfg(test)]
    fn srv_rpc(
        server: &crate::server::Server,
        owned: &mut Vec<u64>,
        events: &Sender<SessionEvent>,
        method: &str,
        params: Value,
    ) -> crate::protocol::Response {
        let mut authenticated = true;
        let req: crate::protocol::Request =
            serde_json::from_value(serde_json::json!({"id": 1, "method": method, "params": params}))
                .expect("a well-formed request");
        server.handle_tracked(req, true, &mut authenticated, owned, events)
    }

    #[cfg(test)]
    fn ole_ref_of(resp: crate::protocol::Response) -> u64 {
        resp.result
            .expect("a successful response carries a result")["$ole_ref"]
            .as_u64()
            .expect("the result is an object reference")
    }

    /// A cleanup that hides the save prompt and quits: the steps a client
    /// would register so its auto-created Excel is ended when the last user
    /// leaves.
    #[cfg(test)]
    fn quit_cleanup() -> Value {
        serde_json::json!({
            "steps": [
                {"name": "DisplayAlerts=", "args": [false]},
                {"name": "Quit", "args": []}
            ]
        })
    }

    /// The same quit steps, but with `callback: true`: the last user leaving via
    /// the wire `release` is asked to run its closure first (the `$cleanup`
    /// path) instead of the bridge quitting inline.
    #[cfg(test)]
    fn quit_cleanup_callback() -> Value {
        serde_json::json!({
            "steps": [
                {"name": "DisplayAlerts=", "args": [false]},
                {"name": "Quit", "args": []}
            ],
            "callback": true
        })
    }

    /// The seq a `release` on the callback path answered with -- the
    /// `{"cleanup": seq}` result. The client answers this seq with
    /// `release_event` once its closure has run.
    #[cfg(test)]
    fn cleanup_seq_of(resp: crate::protocol::Response) -> u64 {
        resp.result
            .expect("a cleanup-pending response carries a result")["cleanup"]
            .as_u64()
            .expect("the result carries a cleanup seq")
    }

    /// The last user of an auto-created Excel runs the cleanup steps inline on
    /// release, ending the instance even though a workbook handle is still
    /// outstanding.
    #[test]
    fn cleanup_steps_quit_an_auto_created_excel_with_an_open_workbook() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();
        let _teardown = ExcelTeardownGuard;

        let server = crate::server::Server::new(None, fresh_registry());
        let mut owned: Vec<u64> = Vec::new();
        let (events, _event_rx) = channel::<SessionEvent>();

        let root = ole_ref_of(srv_rpc(
            &server, &mut owned, &events,
            "create", serde_json::json!({"class_name": "Excel.Application", "cleanup": quit_cleanup()}),
        ));
        srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": root, "name": "Visible=", "args": [false]}));
        // An open workbook gives Excel a reason to stay up, so the
        // disappearance below is the Quit step's doing and not a bare instance
        // exiting on its own.
        let books = ole_ref_of(srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": root, "name": "Workbooks"})));
        srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": books, "name": "Add"}));
        assert!(excel_alive(), "the auto-created Excel must be running before release");

        // Release the ROOT only. The Workbooks/Workbook handles are still
        // outstanding, but a participating session tears down on root release.
        let rel = srv_rpc(&server, &mut owned, &events,
            "release", serde_json::json!({"handle": root}));
        assert!(rel.error.is_none(), "release should succeed: {:?}", rel.error);

        assert!(
            wait_until_excel_gone(15),
            "the last user's Quit cleanup step must end the auto-created Excel"
        );

        drop(events);
    }

    /// Two sessions on the SAME Excel: only the LAST release runs the cleanup.
    /// The first (the creator) leaves without touching Excel; the second (a
    /// connector) is the last user and quits it.
    #[test]
    fn co_users_only_the_last_release_quits() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();
        let _teardown = ExcelTeardownGuard;

        let registry = fresh_registry();
        let server = crate::server::Server::new(None, registry.clone());
        let mut owned: Vec<u64> = Vec::new();
        let (events, _event_rx) = channel::<SessionEvent>();

        // Session 1 auto-creates the Excel and opens a workbook.
        let root1 = ole_ref_of(srv_rpc(
            &server, &mut owned, &events,
            "create", serde_json::json!({"class_name": "Excel.Application", "cleanup": quit_cleanup()}),
        ));
        srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": root1, "name": "Visible=", "args": [false]}));
        let books = ole_ref_of(srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": root1, "name": "Workbooks"})));
        srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": books, "name": "Add"}));
        assert!(excel_alive(), "session 1 must have created a live Excel");

        // Session 2 connects to the SAME Excel (GetActiveObject) and joins as a
        // co-user. The proof they are on ONE Excel is structural: two sessions
        // whose roots resolve to the same cross-apartment (OXID, OID) collapse
        // to a SINGLE registry record with two users. Two Excels -- or a
        // marshaling failure driving either onto a synthetic key -- would show
        // two records instead.
        let root2 = ole_ref_of(srv_rpc(
            &server, &mut owned, &events,
            "connect", serde_json::json!({"class_name": "Excel.Application", "cleanup": quit_cleanup()}),
        ));
        assert_ne!(root1 >> 32, root2 >> 32, "the two sessions must be distinct");
        assert_eq!(
            registry.record_summary(),
            vec![(true, 2, false)],
            "connect must attach to the created Excel: one auto-created record, two users"
        );

        // Session 1 leaves first: NOT the last user, so nothing quits. The
        // release is synchronous, so session 1 has fully torn down by now.
        srv_rpc(&server, &mut owned, &events, "release", serde_json::json!({"handle": root1}));
        assert!(excel_alive(), "releasing a non-last co-user must not end the shared Excel");
        assert_eq!(
            registry.record_summary(),
            vec![(true, 1, false)],
            "the record must survive with the one remaining user"
        );

        // Session 2 leaves last: it runs the cleanup and the Excel goes.
        srv_rpc(&server, &mut owned, &events, "release", serde_json::json!({"handle": root2}));
        assert!(
            wait_until_excel_gone(15),
            "the last co-user's release must run the Quit cleanup step"
        );
        assert!(
            registry.record_summary().is_empty(),
            "the record must be gone once its last user ran cleanup"
        );

        drop(events);
    }

    /// `leave_open` revokes shutdown permission: the auto-created Excel
    /// survives its session's root release, with no cleanup step run.
    #[test]
    fn leave_open_keeps_an_auto_created_excel() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();
        let _teardown = ExcelTeardownGuard;

        let registry = fresh_registry();
        let server = crate::server::Server::new(None, registry.clone());
        let mut owned: Vec<u64> = Vec::new();
        let (events, _event_rx) = channel::<SessionEvent>();

        let root = ole_ref_of(srv_rpc(
            &server, &mut owned, &events,
            "create", serde_json::json!({"class_name": "Excel.Application", "cleanup": quit_cleanup()}),
        ));
        srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": root, "name": "Visible=", "args": [false]}));
        // Suppress the unsaved-workbook prompt for the life of the kept-open
        // Excel: the cleanup Quit that would have set this never runs here.
        srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": root, "name": "DisplayAlerts=", "args": [false]}));
        let books = ole_ref_of(srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": root, "name": "Workbooks"})));
        srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": books, "name": "Add"}));

        // Revoke shutdown permission. The record's auto-created flag clears.
        let lo = srv_rpc(&server, &mut owned, &events,
            "leave_open", serde_json::json!({"handle": root}));
        assert!(lo.error.is_none(), "leave_open should succeed: {:?}", lo.error);
        assert_eq!(
            registry.record_summary(),
            vec![(false, 1, false)],
            "leave_open must revoke the auto-created permission"
        );

        // Release the root: on_root_release now returns NotLast (no cleanup),
        // and the non-auto-created record is dropped as bookkeeping.
        srv_rpc(&server, &mut owned, &events, "release", serde_json::json!({"handle": root}));
        assert!(
            registry.record_summary().is_empty(),
            "a leave_open'd record leaves NotLast and is removed, never closing"
        );

        // The bridge must NOT have quit it: leave_open cleared auto_created, so
        // on_root_release ran no steps. The instance is still registered right
        // after the session that owned it has gone...
        assert!(
            excel_alive(),
            "leave_open must keep the auto-created Excel alive past its session"
        );
        // ...and stays so, rather than merely being caught mid-teardown: still
        // alive two seconds on, where the cleanup path (the test above) has the
        // instance gone well within that window. (Measured: it lingers far
        // longer, kept up by the open workbook.)
        assert!(
            !wait_until_excel_gone(2),
            "a kept-open Excel must not disappear on its own"
        );

        drop(events);
    }

    // ---- The $cleanup client-callback path (Task 5) ----------------------
    //
    // With `callback: true`, the last user leaving via the wire `release` does
    // NOT quit inline: the bridge emits a `$cleanup` event, answers the release
    // with `{"cleanup": seq}`, keeps the session alive, and waits for the
    // client to answer that seq with `release_event`. These drive the `Server`
    // end to end and read the `$cleanup` frame straight off the event channel
    // -- no Ruby client needed. Liveness is COM-based (`GetActiveObject`), the
    // same as the Task 4 tests above, because the test runner is Wine and
    // cannot see the Linux process table.

    /// The whole callback path: `release` answers `{"cleanup": seq}` and emits a
    /// matching `$cleanup` event; the Excel stays up until the client's
    /// `release_event` runs the steps, then it goes.
    #[test]
    fn cleanup_callback_release_event_completes_the_cleanup() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();
        let _teardown = ExcelTeardownGuard;

        let server = crate::server::Server::new(None, fresh_registry());
        let mut owned: Vec<u64> = Vec::new();
        let (events, event_rx) = channel::<SessionEvent>();

        let root = ole_ref_of(srv_rpc(
            &server, &mut owned, &events,
            "create", serde_json::json!({"class_name": "Excel.Application", "cleanup": quit_cleanup_callback()}),
        ));
        srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": root, "name": "Visible=", "args": [false]}));
        // An open workbook gives Excel a reason to stay up, so its later
        // disappearance is the Quit step's doing, not a bare instance exiting.
        let books = ole_ref_of(srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": root, "name": "Workbooks"})));
        srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": books, "name": "Add"}));
        assert!(excel_alive(), "the auto-created Excel must be running before release");

        // Release the ROOT. callback:true + a live client -> the bridge asks
        // the closure to run first: it answers {"cleanup": seq} and stays alive.
        let rel = srv_rpc(&server, &mut owned, &events, "release", serde_json::json!({"handle": root}));
        assert!(rel.error.is_none(), "release should succeed: {:?}", rel.error);
        let seq = cleanup_seq_of(rel);

        // ...and delivers a $cleanup event carrying that seq, which is what the
        // client's closure answers with release_event.
        let ev = event_rx
            .recv_timeout(std::time::Duration::from_secs(10))
            .expect("a $cleanup event must arrive when callback:true and the client is alive");
        assert_eq!(ev.frame.event, "$cleanup", "the callback path emits a $cleanup event");
        assert_eq!(ev.frame.seq, seq, "the $cleanup event carries the same seq as the release reply");
        assert_eq!(ev.frame.handle, root, "the $cleanup event names the root it is closing");
        assert!(ev.frame.args.is_none(), "the $cleanup event carries no arguments");

        // Nothing has run the steps yet: the Excel is deliberately still up.
        assert!(excel_alive(), "the callback path must NOT quit before release_event");

        // The client's closure answers. Now the steps run and the Excel goes.
        let done = srv_rpc(&server, &mut owned, &events, "release_event", serde_json::json!({"seq": seq}));
        assert!(done.error.is_none(), "release_event should succeed: {:?}", done.error);
        assert!(
            wait_until_excel_gone(10),
            "release_event must run the Quit cleanup step and end the Excel"
        );

        drop(events);
    }

    /// Edge 2: the closure calls `leave_open` before answering. `release_event`
    /// then finds the record no longer closing, SKIPS the steps, and the Excel
    /// stays up.
    #[test]
    fn cleanup_callback_leave_open_in_closure_keeps_the_excel() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();
        let _teardown = ExcelTeardownGuard;

        let server = crate::server::Server::new(None, fresh_registry());
        let mut owned: Vec<u64> = Vec::new();
        let (events, event_rx) = channel::<SessionEvent>();

        let root = ole_ref_of(srv_rpc(
            &server, &mut owned, &events,
            "create", serde_json::json!({"class_name": "Excel.Application", "cleanup": quit_cleanup_callback()}),
        ));
        srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": root, "name": "Visible=", "args": [false]}));
        // The Quit step that would have suppressed the save prompt never runs on
        // this path, so set it now for the Excel this test leaves open.
        srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": root, "name": "DisplayAlerts=", "args": [false]}));
        let books = ole_ref_of(srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": root, "name": "Workbooks"})));
        srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": books, "name": "Add"}));
        assert!(excel_alive(), "the auto-created Excel must be running before release");

        let rel = srv_rpc(&server, &mut owned, &events, "release", serde_json::json!({"handle": root}));
        let seq = cleanup_seq_of(rel);
        let ev = event_rx
            .recv_timeout(std::time::Duration::from_secs(10))
            .expect("a $cleanup event must arrive");
        assert_eq!(ev.frame.seq, seq);

        // The closure decides to keep the instance: leave_open BEFORE answering.
        let lo = srv_rpc(&server, &mut owned, &events, "leave_open", serde_json::json!({"handle": root}));
        assert!(lo.error.is_none(), "leave_open should succeed: {:?}", lo.error);

        // Now the closure answers. confirm_cleanup finds the record no longer
        // closing (leave_open cleared it), so the steps are skipped.
        let done = srv_rpc(&server, &mut owned, &events, "release_event", serde_json::json!({"seq": seq}));
        assert!(done.error.is_none(), "release_event should succeed: {:?}", done.error);

        // The Excel stays up -- and not merely caught mid-teardown: still alive
        // three seconds on, where the completing path (the test above) has it
        // gone well within that window.
        assert!(
            !wait_until_excel_gone(3),
            "leave_open in the closure must keep the Excel alive: the steps must be skipped"
        );

        drop(events);
        // ExcelTeardownGuard quits the kept-open Excel by COM.
    }

    /// Edge 6: the client vanishes mid-cleanup -- it never sends `release_event`.
    /// The connection teardown (`release_all`, client_alive:false) is the only
    /// thing left, and the disconnect catch-all must run the steps on the way
    /// out so the Excel is not orphaned.
    #[test]
    fn cleanup_callback_client_disconnect_runs_steps_via_catch_all() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();
        let _teardown = ExcelTeardownGuard;

        let server = crate::server::Server::new(None, fresh_registry());
        let mut owned: Vec<u64> = Vec::new();
        let (events, event_rx) = channel::<SessionEvent>();

        let root = ole_ref_of(srv_rpc(
            &server, &mut owned, &events,
            "create", serde_json::json!({"class_name": "Excel.Application", "cleanup": quit_cleanup_callback()}),
        ));
        srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": root, "name": "Visible=", "args": [false]}));
        let books = ole_ref_of(srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": root, "name": "Workbooks"})));
        srv_rpc(&server, &mut owned, &events,
            "invoke", serde_json::json!({"handle": books, "name": "Add"}));
        assert!(excel_alive(), "the auto-created Excel must be running before release");

        // The root release opens a $cleanup and keeps the session alive...
        let rel = srv_rpc(&server, &mut owned, &events, "release", serde_json::json!({"handle": root}));
        let seq = cleanup_seq_of(rel);
        let ev = event_rx
            .recv_timeout(std::time::Duration::from_secs(10))
            .expect("a $cleanup event must arrive");
        assert_eq!(ev.frame.seq, seq);
        assert!(excel_alive(), "still alive before the client answers");

        // ...but the client vanishes: no release_event. The connection tears
        // down, releasing everything it owned with client_alive:false. The last
        // of those (the root) breaks the loop with the $cleanup still pending.
        server.release_all(&owned);

        // The disconnect catch-all must run the steps as the session exits, or
        // the Excel (with its open workbook) would be orphaned.
        assert!(
            wait_until_excel_gone(10),
            "the disconnect catch-all must run the Quit cleanup step when the client vanishes mid-cleanup"
        );

        drop(events);
    }

    /// Measurement gate: the instance-lifetime design runs cleanup steps with
    /// no timeout and continues past a failing step (see `run_cleanup_steps`
    /// above). That is only safe if a step fired at an already-dead Excel
    /// returns promptly instead of hanging. This fires the worst realistic
    /// case directly at the COM layer -- a property-set on an Excel that has
    /// already been `Quit` -- and records how long it takes.
    #[test]
    fn measure_step_against_quit_excel() {
        use crate::dispatch::{create_instance, invoke_member, lock_excel_for_test};
        use windows::core::VARIANT;
        use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};
        use std::time::Instant;

        let _guard = lock_excel_for_test();
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap() };
        let xl = create_instance("Excel.Application").unwrap();
        invoke_member(&xl, "DisplayAlerts=", vec![VARIANT::from(false)], vec![]).unwrap();
        invoke_member(&xl, "Quit", vec![], vec![]).unwrap();
        let t = Instant::now();
        let r = invoke_member(&xl, "DisplayAlerts=", vec![VARIANT::from(false)], vec![]);
        let elapsed = t.elapsed();
        eprintln!("[measure] DisplayAlerts= after Quit -> {:?} in {:?}", r.map(|_| "ok"), elapsed);
        // No assertion on the value; the assertion is that it returns promptly.
        assert!(elapsed.as_secs() < 30, "a step against a quit Excel must not hang");
        drop(xl);
        unsafe { CoUninitialize() };
    }
}

