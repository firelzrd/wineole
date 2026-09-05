import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, os.path.dirname(__file__))

from support.fake_com import FakeComRange

from wineole.msoffice.range import Range


def make_range(rows, cols, value=None):
    return Range(FakeComRange(rows=rows, cols=cols, value=value))


class MSOfficeRangeTest(unittest.TestCase):
    # --- to_list ----------------------------------------------------------

    def test_to_list_wraps_a_scalar_from_a_single_cell(self):
        # Excel returns a bare scalar for a one-cell range, whatever the address.
        self.assertEqual([['a1']], make_range(1, 1, 'a1').to_list())

    def test_to_list_passes_a_two_dimensional_value_through(self):
        v = [[1, 2], [3, 4]]
        self.assertEqual(v, make_range(2, 2, v).to_list())

    def test_to_list_wraps_an_empty_single_cell(self):
        self.assertEqual([[None]], make_range(1, 1, None).to_list())

    # --- write: accepted ---------------------------------------------------

    def test_write_broadcasts_a_scalar(self):
        r = make_range(3, 3)
        r.write(7)
        self.assertEqual(7, r.ole.written, 'Excel broadcasts a scalar itself; pass it through')

    def test_write_orients_a_flat_list_down_a_column(self):
        # Excel's own behaviour here replicates the first element down the
        # column; this is the trap `write` exists to close.
        r = make_range(3, 1)
        r.write([1, 2, 3])
        self.assertEqual([[1], [2], [3]], r.ole.written)

    def test_write_orients_a_flat_list_across_a_row(self):
        r = make_range(1, 3)
        r.write([1, 2, 3])
        self.assertEqual([[1, 2, 3]], r.ole.written)

    def test_write_passes_a_matching_two_dimensional_list_through(self):
        r = make_range(2, 3)
        r.write([[1, 2, 3], [4, 5, 6]])
        self.assertEqual([[1, 2, 3], [4, 5, 6]], r.ole.written)

    def test_write_treats_a_string_as_a_scalar(self):
        r = make_range(3, 3)
        r.write('hello')
        self.assertEqual('hello', r.ole.written,
                         'a str is iterable in Python, and must never be read as a row')

    # --- write: rejected ---------------------------------------------------

    def test_write_rejects_a_flat_list_of_the_wrong_length(self):
        with self.assertRaises(ValueError) as ctx:
            make_range(3, 1).write([1, 2])
        self.assertIn('3x1', str(ctx.exception), 'the message must state the range size')
        self.assertIn('2', str(ctx.exception), 'and what it was given')

    def test_write_rejects_a_flat_list_when_the_range_is_not_a_line(self):
        with self.assertRaises(ValueError) as ctx:
            make_range(3, 3).write([1, 2, 3])
        self.assertIn('3x3', str(ctx.exception))

    def test_write_rejects_a_two_dimensional_list_of_the_wrong_size(self):
        with self.assertRaises(ValueError) as ctx:
            make_range(3, 3).write([[1, 2, 3, 4, 5]])
        self.assertIn('3x3', str(ctx.exception))
        self.assertIn('1x5', str(ctx.exception))

    def test_write_rejects_a_ragged_list(self):
        with self.assertRaises(ValueError) as ctx:
            make_range(2, 2).write([[1, 2], [3]])
        self.assertIn('ragged', str(ctx.exception))

    def test_write_rejects_rows_mixed_with_scalars(self):
        with self.assertRaises(ValueError) as ctx:
            make_range(2, 2).write([[1, 2], 3])
        self.assertIn('mixes rows and scalars', str(ctx.exception))

    def test_write_rejects_scalars_mixed_with_rows(self):
        with self.assertRaises(ValueError) as ctx:
            make_range(2, 2).write([1, [2, 3]])
        self.assertIn('mixes scalars and rows', str(ctx.exception))

    def test_write_rejects_a_value_nested_more_than_two_deep(self):
        with self.assertRaises(ValueError) as ctx:
            make_range(1, 1).write([[[1]]])
        self.assertIn('nests more than two deep', str(ctx.exception))

    # --- fill ---------------------------------------------------------------

    def test_fill_replicates_a_scalar_over_the_whole_range(self):
        r = make_range(3, 3)
        r.fill(7)
        self.assertEqual([[7, 7, 7], [7, 7, 7], [7, 7, 7]], r.ole.written)

    def test_fill_treats_a_flat_list_as_a_row_and_replicates_it_down(self):
        r = make_range(3, 3)
        r.fill([1, 2])
        self.assertEqual([[1, 2, None], [1, 2, None], [1, 2, None]], r.ole.written)

    def test_fill_truncates_and_pads_a_two_dimensional_list(self):
        r = make_range(3, 3)
        r.fill([[1, 2, 3, 4, 5]])
        self.assertEqual([[1, 2, 3], [None, None, None], [None, None, None]], r.ole.written)

    # --- passthrough ---------------------------------------------------------

    def test_unknown_methods_go_to_com(self):
        r = make_range(1, 1)
        self.assertIs(r.ole.interior, r.Interior())

    def test_value_is_not_intercepted(self):
        # `Value` must stay exactly Excel's -- scalar for one cell -- so
        # that Excel documentation and VBA knowledge keep applying. The ()
        # is the raw client's own spelling.
        self.assertEqual('a1', make_range(1, 1, 'a1').Value())

    def test_the_range_is_not_iterable(self):
        # to_list is an explicit conversion and is ours; implicit iteration
        # is what `for` and unpacking reach for, and Proxy deliberately
        # refuses it. Allowing it here would undo that.
        with self.assertRaises(TypeError):
            iter(make_range(1, 1))

    # --- format delegates and returns self -----------------------------------

    def test_format_returns_the_range_itself(self):
        r = make_range(1, 1)
        self.assertIs(r, r.format(wrap=True))

    def test_format_reaches_the_com_object(self):
        fake = FakeComRange(rows=1, cols=1)
        Range(fake).format(wrap=True)
        self.assertIs(True, fake.writes['WrapText'])

    def test_format_refuses_an_unknown_key(self):
        with self.assertRaises(ValueError) as ctx:
            make_range(1, 1).format(**{'blod': True})
        self.assertIn('blod', str(ctx.exception))


if __name__ == '__main__':
    unittest.main()
