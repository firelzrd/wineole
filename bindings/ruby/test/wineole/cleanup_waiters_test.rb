require 'minitest/autorun'
require 'timeout'
require_relative '../../lib/wineole/client'

# The await/signal handshake behind Client#await_cleanup and
# #signal_cleanup_done, tested in isolation from Client itself: building a
# real Client for this would open a socket, and the coordination has nothing
# to do with the wire.
class CleanupWaitersTest < Minitest::Test
  def test_signal_from_another_thread_unblocks_a_waiting_thread
    waiters = WineOLE::Client::CleanupWaiters.new
    returned = false

    waiter = Thread.new do
      waiters.await(7)
      returned = true
    end

    # Widen the race window so the waiter thread is actually parked inside
    # `await` before the signal is sent -- this test is specifically about
    # the broadcast waking an ALREADY-waiting thread, not about the
    # already-signalled case the next test covers.
    sleep 0.05
    refute returned, 'must still be blocked before signal is sent'

    waiters.signal(7)

    assert waiter.join(5), 'await must return once another thread calls signal'
    assert returned
  end

  def test_await_returns_immediately_when_already_signalled
    waiters = WineOLE::Client::CleanupWaiters.new
    waiters.signal(9)

    # Bounded well under CleanupWaiters::TIMEOUT (30s): if `await` actually
    # had to wait for that, this would time out and fail the test instead of
    # merely taking longer than it should.
    Timeout.timeout(1) { waiters.await(9) }
  end

  def test_different_seqs_do_not_share_state
    waiters = WineOLE::Client::CleanupWaiters.new
    waiters.signal(1)

    # Signalling seq 1 must not leak into seq 2: a thread awaiting a
    # different, unsignalled seq blocks until it is signalled on its own.
    returned = false
    waiter = Thread.new do
      waiters.await(2)
      returned = true
    end
    sleep 0.05
    refute returned, 'seq 2 must not be affected by seq 1 having been signalled'

    waiters.signal(2)
    assert waiter.join(5)
  end
end
