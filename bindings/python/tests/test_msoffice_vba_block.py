import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, os.path.dirname(__file__))

from support.fake_com import FakeCodeModule

from wineole.msoffice.vba_block import VBABlock


class MSOfficeVBABlockTest(unittest.TestCase):
    def test_a_block_is_wrapped_in_its_own_markers(self):
        cm = FakeCodeModule()
        VBABlock.write(cm, 'go', 'Sub Go()\nEnd Sub')
        self.assertIn("'<wineole:go>", cm.text)
        self.assertIn("'</wineole:go>", cm.text)
        self.assertIn('Sub Go()', cm.text)

    def test_rewriting_a_block_replaces_it_rather_than_stacking(self):
        cm = FakeCodeModule()
        VBABlock.write(cm, 'go', 'Sub Go()\nEnd Sub')
        VBABlock.write(cm, 'go', "Sub Go()\n  ' second\nEnd Sub")
        self.assertEqual(1, cm.text.count("'<wineole:go>"),
                         'a second write must replace the first, not add another copy')
        self.assertIn("' second", cm.text)

    # VBA identifiers are case-insensitive, and so are the collection lookups
    # that hand a name back ('appform' finds AppForm). Two blocks whose names
    # differ only in case would therefore hold procedures that collide --
    # "Ambiguous name detected", and every Application.Run into the module
    # fails from then on -- so the second write must replace the first.
    def test_a_name_differing_only_in_case_replaces_the_block(self):
        cm = FakeCodeModule()
        VBABlock.write(cm, 'main', 'Sub Go()\nEnd Sub')
        VBABlock.write(cm, 'Main', "Sub Go()\n  ' second\nEnd Sub")
        self.assertEqual(1, cm.text.lower().count("'<wineole:main>"),
                         'a write whose name differs only in case must replace, '
                         'not add another copy')
        self.assertIn("' second", cm.text)
        self.assertNotIn("'<wineole:main>", cm.text, 'the old block went with its markers')

    def test_remove_matches_the_name_case_insensitively(self):
        cm = FakeCodeModule('Sub Handwritten()\nEnd Sub')
        VBABlock.write(cm, 'main', 'Sub Go()')
        self.assertEqual(['Sub Handwritten()', 'End Sub'], VBABlock.remove(cm, 'MAIN'))

    # The whole point: the module is not ours, only the block is.
    def test_other_code_in_the_module_survives(self):
        cm = FakeCodeModule('Sub Handwritten()\nEnd Sub')
        VBABlock.write(cm, 'go', 'Sub Go()\nEnd Sub')
        VBABlock.write(cm, 'other', 'Sub Other()\nEnd Sub')
        VBABlock.remove(cm, 'go')
        self.assertIn('Sub Handwritten()', cm.text)
        self.assertIn('Sub Other()', cm.text)
        self.assertNotIn('Sub Go()', cm.text)

    # False when there was nothing of this name to remove; otherwise the
    # module's remaining lines, so the caller does not have to refetch the
    # body just to find out whether it is now blank.
    def test_remove_reports_whether_there_was_anything_to_remove(self):
        cm = FakeCodeModule()
        self.assertIs(False, VBABlock.remove(cm, 'go'))
        VBABlock.write(cm, 'go', 'Sub Go()')
        self.assertEqual([], VBABlock.remove(cm, 'go'))

    def test_remove_hands_back_the_lines_left_over_after_the_block(self):
        cm = FakeCodeModule('Sub Handwritten()\nEnd Sub')
        VBABlock.write(cm, 'go', 'Sub Go()')
        self.assertEqual(['Sub Handwritten()', 'End Sub'], VBABlock.remove(cm, 'go'))

    # Excel reports 0 lines for a module that has never been written to, and
    # Lines(1, 0) is not a legal call.
    def test_an_empty_module_is_not_read(self):
        cm = FakeCodeModule()
        self.assertIs(False, VBABlock.remove(cm, 'go'))
        self.assertEqual(0, cm.reads)

    def test_the_body_is_fetched_once_per_operation(self):
        cm = FakeCodeModule('Sub A()\nEnd Sub')
        VBABlock.write(cm, 'go', 'Sub Go()')
        self.assertEqual(1, cm.reads, 'one Lines call, not one per line')

    # Measured: a module emptied of every block still reports 2 lines of
    # "\r\n" -- CountOfLines == 2, not 0. Blank means blank after stripping,
    # not CountOfLines == 0.
    def test_blank_sees_through_leftover_newlines(self):
        cm = FakeCodeModule(lines=['', ''])
        self.assertEqual(2, cm.CountOfLines())
        self.assertIs(True, VBABlock.blank(cm))
        self.assertIs(True, VBABlock.blank(FakeCodeModule()))
        self.assertIs(False, VBABlock.blank(FakeCodeModule('Option Explicit')))

    def test_a_name_that_would_break_the_marker_is_refused(self):
        cm = FakeCodeModule()
        for bad in ['a>b', 'a\nb', '', 'a b']:
            with self.assertRaises(ValueError, msg=f"{bad!r} must be refused"):
                VBABlock.write(cm, bad, 'x')

    def test_a_name_that_is_not_a_string_is_a_type_error(self):
        cm = FakeCodeModule()
        with self.assertRaises(TypeError):
            VBABlock.write(cm, 123, 'x')
        self.assertEqual('', cm.text)

    def test_only_one_trailing_newline_is_dropped_like_ruby_chomp(self):
        # Ruby's chomp keeps a deliberate blank line at the end of the
        # payload; stripping every trailing newline would silently reshape
        # the module text against what the Ruby side writes.
        cm = FakeCodeModule()
        VBABlock.write(cm, 'x', 'Sub X()\nEnd Sub\n\n')
        self.assertEqual("'<wineole:x>\nSub X()\nEnd Sub\n\n'</wineole:x>", cm.text)

    def test_an_unclosed_marker_is_reported_rather_than_guessed(self):
        cm = FakeCodeModule("'<wineole:go>\nSub Go()\nEnd Sub")
        with self.assertRaises(ValueError) as caught:
            VBABlock.remove(cm, 'go')
        self.assertIn('go', str(caught.exception))

    # A close marker with no matching open is a corrupted module, not an
    # absence -- silently reporting False here is what let the module stay
    # broken forever in the reproduction the review found.
    def test_a_close_marker_with_no_open_is_reported_as_corruption(self):
        cm = FakeCodeModule("'</wineole:go>\nEnd Sub")
        with self.assertRaises(ValueError) as caught:
            VBABlock.remove(cm, 'go')
        self.assertIn('go', str(caught.exception))

    # write must refuse a payload containing a line that is itself a marker
    # -- for any name, not just the one being written -- because remove
    # cannot tell it apart from a real one afterwards.
    def test_a_marker_shaped_line_in_the_payload_is_refused(self):
        cm = FakeCodeModule()
        with self.assertRaises(ValueError) as caught:
            VBABlock.write(cm, 'go', "Sub Go()\n'</wineole:go>\nEnd Sub")
        self.assertIn('wineole:go', str(caught.exception))
        self.assertEqual('', cm.text, 'a refused write must not touch the module at all')

    def test_a_marker_shaped_line_for_a_different_name_is_also_refused(self):
        cm = FakeCodeModule()
        with self.assertRaises(ValueError):
            VBABlock.write(cm, 'go', "'<wineole:other>\nEnd Sub")
        with self.assertRaises(ValueError):
            VBABlock.write(cm, 'go', "'</wineole:other>\nEnd Sub")

    # Reproduces the review's exact scenario: without the fix, this used to
    # delete up to the caller's accidental marker, leave "End Sub" plus an
    # orphaned close marker behind, and make blank never true again.
    def test_the_reproduction_from_the_review_is_now_refused_up_front(self):
        cm = FakeCodeModule()
        with self.assertRaises(ValueError):
            VBABlock.write(cm, 'go', "Sub Go()\n'</wineole:go>\nEnd Sub")
        self.assertIs(True, VBABlock.blank(cm),
                      'the refused write must never have touched the module')
        self.assertIs(False, VBABlock.remove(cm, 'go'),
                      'nothing was ever added, so there is nothing to remove')


if __name__ == '__main__':
    unittest.main()
