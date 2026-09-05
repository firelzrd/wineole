require 'weakref'

# One construct for every "was this actually collected?" assertion in the
# suite, because MRI gives no straight answer to that question.
#
# MRI scans the machine stack conservatively. A value that is dead in the
# source can still be sitting in a register, or in a stack slot belonging to a
# frame the test has already returned from, and the collector -- which cannot
# tell a stale word from a live pointer -- has to assume it is live. So a lone
# `GC.start` after dropping the last reference is not a measurement, it is a
# coin flip; and the more code that ran first, the worse the odds, because
# there are more dead frames below to leave debris in. Measured on this suite:
# `ClientTest#test_finalizer_closes_the_socket_without_capturing_self` passed
# 5 runs of 5 on its own, and failed 4 runs of 10 inside the full unit suite,
# on identical code.
#
# That flakiness is dangerous out of all proportion to the nuisance. A GC
# assertion is the ONLY thing in this suite that catches a reference leak --
# this project has found the same leak three times, at three layers (a reader
# thread whose block captured the Client, a sink list holding it, a dispatcher
# thread rooting its Events), and every other test passed each time. An
# assertion that cries wolf gets dismissed as "that flaky GC test", and the
# next real leak walks straight through behind it. So the answer is to make
# the question answerable, not to soften what is being asked.
#
# Two techniques, because there are two separate causes:
#
#   * The last strong reference is a live local in the test body. `x = nil`
#     does not unbuild the slot the collector already found; the frame has to
#     GO. That is `weak_ref_to`.
#   * A stale pointer is left in a dead frame further down the stack. Nothing
#     will drop it except writing over it. That is `churn_the_stack`, retried
#     on a bounded budget by `collected?`.
#
# What this deliberately does NOT do is weaken the assertion. A genuinely
# pinned object is reachable from a real root and is never collected, however
# many times this asks -- so the retries cannot turn a real leak green. That
# property is not assumed: each of the three leaks above is reintroduced and
# the corresponding assertion is confirmed to still fail.
module GarbageCollectionHelper
  # How many churn-and-collect rounds `collected?` will spend before it
  # concludes the object is pinned. Generous on purpose: every round costs
  # microseconds when the answer is "collected" (the loop exits on the first
  # one that works -- measured: one round has always been enough), and the
  # only case that pays the full budget is a test that is about to fail
  # anyway, where being sure is worth more than being quick.
  COLLECTION_TRIES = 20

  # Builds an object on a thread of its own, waits for that thread to die, and
  # hands back nothing but a weak reference to what it built. The calling
  # thread's stack therefore never holds a pointer to the object at all.
  #
  # A separate THREAD, not merely a separate frame, and that distinction is
  # the whole reason this works. Returning from a frame does not erase it: the
  # words it wrote stay on the machine stack until something else writes over
  # them, and the frames that come next -- `assert_collected`, `collected?` --
  # sit at exactly those addresses without necessarily overwriting every word.
  # Their own live frames then hold the stale pointer, where churning cannot
  # reach it, because recursion only ever overwrites addresses DEEPER than
  # where it started. A thread has its own machine stack, and MRI releases it
  # when the thread dies, so the whole region goes away rather than being
  # overwritten in part.
  #
  # Measured, 2000 build-and-collect cycles each with debris left on the stack
  # beforehand, both with the same churn-and-retry budget behind them:
  # built in a frame of this thread, 1000 of 2000 collected; built on a thread
  # of its own, 2000 of 2000. The whole-suite figure agrees -- the frame
  # version still failed 2 runs in 10.
  #
  # The block's value is what gets weakened. Anything else the test still
  # needs afterwards (a Thread to join, a socket to read to EOF) can be
  # assigned to a local of the enclosing scope from inside the block, as long
  # as it is not the object under test and does not hold one. The block runs
  # on the temporary thread, so anything it raises is re-raised here by
  # `join` -- a failure inside it is reported, not swallowed.
  def weak_ref_to
    builder = Thread.new { WeakRef.new(yield) }
    builder.join
    builder.value
  end

  # Asserts that the object behind `weak` is collectable, spending up to
  # COLLECTION_TRIES rounds establishing it.
  def assert_collected(weak, message = nil)
    assert collected?(weak),
           message || 'the object must be collectible -- something is pinning it'
  end

  # Polled rather than sampled once. Each round overwrites the stack region
  # the dead frames occupied and asks the collector again; the first round
  # that succeeds ends it.
  def collected?(weak, tries: COLLECTION_TRIES)
    tries.times do
      churn_the_stack
      GC.start
      return true unless weak.weakref_alive?
    end
    false
  end

  # The other direction: "this must SURVIVE a collection."
  #
  # There is nothing to poll for here -- the test wants the hardest collection
  # it can get, once, and then asks whether the thing still works. It matters
  # that this is just as thorough as `collected?`: these assertions catch the
  # mirror-image bug (something held weakly that should be held strongly), and
  # a lazy `GC.start` that fails to collect the object is a leak-catcher that
  # does not catch. Same churn, same passes, no early exit.
  def collect_garbage(passes: 3)
    passes.times do
      churn_the_stack
      GC.start
    end
    nil
  end

  # Recurses to overwrite the machine-stack region that frames this test has
  # already returned from used to occupy, so that a stale pointer left down
  # there stops keeping a dead object marked. Allocating a String, an Array
  # and a Hash per frame is what makes sure those words are actually written
  # rather than merely stepped over, and leaves collectable garbage behind as
  # well. The return value is summed only so that nothing here is optimised
  # away as dead.
  def churn_the_stack(depth = 40)
    return 0 if depth.zero?

    a, b, c = depth.to_s, [depth, depth], {depth => depth.to_s}
    a.size + b.size + c.size + churn_the_stack(depth - 1)
  end
end
