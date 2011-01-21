# -*- coding: utf-8 -*-

"""
Each test must start with test_

clear() is called between each test.
Sheets Sheet1 and Sheet2 can be used for tests.

Following parameters can be passed to run_tests.py:

  --dev    Only tests with @dev decorator will be executed
  --stop   Errors are raised

If no errors are encountered, all tests will be merged to a test document to be
run as macro.

"""

def clear():
    S('a1:z100').delete()

def test_column_name_vs_index_conversion():
    assert S()._col_index('A') == 0
    assert S()._col_index('B') == 1
    assert S()._col_index('c') == 2 
    assert S()._col_index('Z') == 25
    assert S()._col_index('AA') == 26
    assert S()._col_index('AF') == 31

    assert S()._col_name(0) == 'A'
    assert S()._col_name(1) == 'B'
    assert S()._col_name(25) == 'Z'
    assert S()._col_name(26) == 'AA'
    assert S()._col_name(31) == 'AF'

def test_value():
    S('a1').value = 10

    assert S('a1').value == 10
    assert S('a1').formula == u'10'
    assert S('a1').string == u'10'

def test_string():
    S('a1').string = u'Hello'

    assert S('a1').string == u'Hello'
    assert S('a1').formula == u'Hello'
    assert S('a1').value == 0

    S('a1').string = u'10'
    assert S('a1').formula == u"'10"
    assert S('a1').value == 0

def test_formula():
    S('a1').value = 10
    S('a2').formula = '=a1+5'

    assert S('a2').value == 15
    assert S('a2').string == u'15'
    assert S('a2').formula == '=A1+5'

def test_date():
    S('a1').date = datetime(2010, 12, 17)

    assert S('a1').date == datetime(2010, 12, 17)
    S('a1').date += timedelta(5)
    assert S('a1').date == datetime(2010, 12, 22)
    assert '/' in S('a1').string

def test_data_of_multiple_cells_can_be_changed():
    S('a1:g10').value = 5
    assert S('d5').value == 5
    S('a1:g10').set_value(6)
    assert S('c4').value == 6
    
    S('a1:g10').string = 'hello'
    assert S('e8').string == 'hello'
    S('a1:g10').set_string('world')
    assert S('f2').string == 'world'

    S('a1:g10').date = datetime(2011, 1, 20)
    assert S('e7').date == datetime(2011, 1, 20)
    S('a1:g10').set_date(datetime(2011, 1, 21))
    assert S('f4').date == datetime(2011, 1, 21)

    S('a1').value = 1
    S('a2:g5').formula = '=a1+3'
    assert S('b3').value == 4
    S('a2:g5').set_formula('=a1+5')
    assert S('b3').value == 6

    S('a1:h11').delete()
    S('b2:g10').value = 17
    assert S('a1').value == 0
    assert S('b1').value == 0
    assert S('a2').value == 0
    assert S('b2').value == 17
    assert S('g11').value == 0
    assert S('h10').value == 0
    assert S('h11').value == 0
    assert S('g10').value == 17

def test_cell_contents_can_be_set_by_methods_which_can_be_cascaded():
    S('a1').set_value(1).drag_to('a5')
    assert S('a5').value == 5
    
    S('a1').set_string('hello').drag_to('a5')
    assert S('a5').string == 'hello'

    S('a1').value = 1
    S('a2').set_formula('=a1*2').drag_to('a5')
    assert S('a5').value == 16

    S('a1').set_date(datetime(2011, 1, 13)).drag_to('a5')
    assert S('a5').date == datetime(2011, 1, 17)

def test_drag_to():
    S('a1').value = 10
    S('a2').formula = '=a1+5'
    S('a2').drag_to('a3')

    assert S('a3').value == 20

def test_drag_to_with_cell_range():
    S('a1').value = 10
    S('a2').value = 20
    S('a3').value = 30

    S('a1:a3').drag_to('b3')

    assert S('b1').value == 11
    assert S('b2').value == 21
    assert S('b3').value == 31

def test_drag_calls_can_be_cascaded():
    S('a1').value = 1
    S('a1').drag_to('a5').drag_to('c5')
    assert S('c5').value == 7

def test_selector_handles_sheets():
    """This test requires english OpenOffice"""
    S('a1').value = 2
    S('Sheet2.a1').value = 5

    assert S('a1').value == 2
    assert S('Sheet2.a1').value == 5

    S('Sheet2.a2').value = 3
    S('Sheet2.a1:a2').drag_to('b2')

    assert S('Sheet2.b1').value == 6
    assert S('Sheet2.b2').value == 4

    S('Sheet2.a1:g10').delete()

def test_delete():
    S('a1').value = 1
    S('a1').delete()

    assert S('a1').value == 0
    assert S('a1').string == ''

    S('a1').value = 1
    S('a2').value = 1
    S('b1').value = 1
    S('b2').value = 1
    S('a1:b2').delete()

    assert S('a1').value == 0
    assert S('a1').string == ''
    assert S('a2').value == 0
    assert S('a2').string == ''
    assert S('b1').value == 0
    assert S('b1').string == ''
    assert S('b2').value == 0
    assert S('b2').string == ''


def test_insert_row():
    S('a1').value = 10
    S('b2').formula = '=a1+5'

    S('b2').insert_row()

    assert S('b3').formula.lower() == '=a1+5'
    assert S('b2').value == 0
    assert S('b2').string == ''

def test_insert_row_expands_selector_and_can_be_cascaded():
    S('a1').value = 10
    S('a2').formula = '=a1+5'
    S('b1').value = 12

    S('a2').insert_row().drag_to('b3')

    assert S('b3').value == 17

def test_insert_column():
    S('a1').value = 10
    S('b2').formula = '=a1+5'

    S('b2').insert_column()

    assert S('c2').formula.lower() == '=a1+5'
    assert S('b2').value == 0
    assert S('b2').string == ''

def test_insert_column_expands_selector_and_can_be_cascaded():
    S('a1').value = 10
    S('b1').formula = '=a1+5'
    S('a2').value = 12

    S('b1').insert_column().drag_to('c2')

    assert S('c2').value == 17


def test_delete_rows():
    S('d5').value = 2
    S('a2').delete_rows()

    assert S('d4').value == 2

    S('a1:a3').delete_rows()

    assert S('d1').value == 2

def test_delete_columns():
    S('f5').value = 2
    S('a2').delete_columns()

    assert S('e5').value == 2

    S('a1:d1').delete_columns()

    assert S('a5').value == 2

def test_copy_cut_and_paste():
    S('a1').value = 4
    S('a1').copy()
    S('b2').paste()

    assert S('a1').value == 4
    assert S('b2').value == 4

    S('a1').cut()
    S('c1').paste()

    assert S('a1').value == 0
    assert S('c1').value == 4

def test_copy_cut_and_paste_can_be_cascaded():
    S('a1').set_value(12).copy().set_value(15).shift_right().paste().shift_down().set_value(18).cut().shift_left().paste()
    assert S('a1').value == 15
    assert S('b1').value == 12
    assert S('b2').value == 0
    assert S('a2').value == 18

def test_delete():
    S('a1').value = 10
    S('a1').delete()
    assert S('a1').value == 0

    S('a1').string = 'hello'
    S('a1').delete()
    assert S('a1').string == ''

def test_undo_redo():
    S('a1').value = 1
    S('a1').value = 2
    S('a1').value = 3
    S('a1').value = 4
    S('a1').value = 5

    S().undo()
    assert S('a1').value == 4
    S().undo()
    S().undo()
    assert S('a1').value == 2
    S().redo()
    assert S('a1').value == 3
    S().redo()
    assert S('a1').value == 4
    S().undo()
    assert S('a1').value == 3

def test_save_as():
    filename = '/tmp/test_oosheet.ods'
    assert not os.path.exists(filename)
    S().save_as(filename)
    assert os.path.exists(filename)
    os.remove(filename)

def test_shift_until_works_for_single_cell_with_value_as_parameter():
    S('g10').string = 'total'

    assert str(S('g1').shift_down_until('total')).endswith('G10')
    assert str(S('g20').shift_up_until('total')).endswith('G10')
    assert str(S('a10').shift_right_until('total')).endswith('G10')
    assert str(S('z10').shift_left_until('total')).endswith('G10')

    S('g10').value = 18
    assert str(S('g1').shift_down_until(18)).endswith('G10')

    S('g10').value = 18.5
    assert str(S('g1').shift_down_until(18.5)).endswith('G10')

    date = datetime(2011, 1, 20)
    S('g10').date = date
    assert str(S('g1').shift_down_until(date)).endswith('G10')

def test_shift_until_works_with_conditions_for_one_dimension_selectors():
    date = datetime(2011, 1, 20)

    S('c10').string = 'total'
    S('d11').value = 19
    S('e12').value = 19.5
    S('f13').date = date
    S('c14').value = 20

    assert str(S('a1:z1').shift_down_until(column_c = 'total')).endswith('.A10:Z10')
    assert str(S('a1:z1').shift_down_until(column_d = 19)).endswith('.A11:Z11')
    assert str(S('a1:z1').shift_down_until(column_e = 19.5)).endswith('.A12:Z12')
    assert str(S('a1:z1').shift_down_until(column_f = date)).endswith('.A13:Z13')
    assert str(S('a1:z1').shift_down_until(column_c = 20)).endswith('.A14:Z14')

    assert str(S('a30:z30').shift_up_until(column_c = 'total')).endswith('.A10:Z10')
    assert str(S('a1:a30').shift_right_until(row_11 = 19)).endswith('.D1:D30')
    assert str(S('z1:z30').shift_left_until(row_12 = 19.5)).endswith('.E1:E30')

def test_shift_until_works_with_conditions_for_two_dimension_selectors():
    date = datetime(2011, 1, 20)

    S('c10').string = 'total'
    S('d11').value = 19
    S('e12').value = 19.5
    S('f13').date = date
    S('c14').value = 20

    assert str(S('a1:z2').shift_down_until(column_c = 'total')).endswith('.A9:Z10')
    assert str(S('a1:z2').shift_down_until(column_d = 19)).endswith('.A10:Z11')
    assert str(S('a1:z2').shift_down_until(column_e = 19.5)).endswith('.A11:Z12')
    assert str(S('a1:z2').shift_down_until(column_f = date)).endswith('.A12:Z13')
    assert str(S('a1:z4').shift_down_until(column_c = 20)).endswith('.A11:Z14')

    assert str(S('a20:z30').shift_up_until(column_c = 'total')).endswith('.A10:Z20')
    assert str(S('a1:c30').shift_right_until(row_11 = 19)).endswith('.B1:D30')
    assert str(S('x1:z30').shift_left_until(row_12 = 19.5)).endswith('.E1:G30')

def test_shift_until_accepts_lambda_to_test_condition():
    S('f10').string = 'some stuff'
    S('g10').string = 'one string'
    S('h11').string = 'another string'
    S('h12').string = 'another stuff'

    assert str(S('a1:z1').shift_down_until(column_g_satisfies = lambda c: c.string.endswith('string'))).endswith('.A10:Z10')
    assert str(S('a1:z2').shift_down_until(column_h_satisfies = lambda c: c.string.startswith('another'))).endswith('.A10:Z11')
    assert str(S('a1:z2').shift_down_until(column_h_satisfies = lambda c: c.string.endswith('stuff'))).endswith('.A11:Z12')
    assert str(S('a1:a20').shift_right_until(row_10_satisfies = lambda c: c.string.endswith('string'))).endswith('.G1:G20')

def test_shift_until_accepts_none_for_empty_cell():
    S('a1').set_value(1).drag_to('g1').drag_to('g10')
    S('g10').delete()

    assert str(S('b1').shift_right_until(row_1 = None)).endswith('.H1')
    assert str(S('b1').shift_down_until(column_b = None)).endswith('.B11')
    assert str(S('b1:5').shift_down_until(column_c = None)).endswith('.B7:B11')
    assert str(S('b1:5').shift_right_until(row_2 = None)).endswith('.H1:H5')
    assert str(S('a2:z2').shift_down_until(column_f = None)).endswith('.A11:Z11')
    assert str(S('a2:z2').shift_down_until(column_g = None)).endswith('.A10:Z10')

def test_shift_right():
    S('a1').set_value(1).drag_to('a10').drag_to('f10')
    S('c1').set_value(100).drag_to('c10')

    S('b1:b5').shift_right().drag_to('d5')
    assert S('d3').value == 103

    S('a6:a10').shift_right(2).drag_to('d10')
    assert S('d7').value == 107

def test_shift_left():
    S('a1').set_value(1).drag_to('a10').drag_to('f10')
    S('c1').set_value(102).drag_to('c10')

    S('d1:d5').shift_left().drag_to('b5')
    assert S('b3').value == 103

    S('e6:e10').shift_left(2).drag_to('b10')
    assert S('b7').value == 107

def test_shift_down():
    S('a1').set_value(1).drag_to('a10').drag_to('g10')
    S('a4').set_value(100).drag_to('g4')

    S('a3:d3').shift_down().drag_to('d5')
    assert S('c5').value == 103

    S('e2:g2').shift_down(2).drag_to('g5')
    assert S('f5').value == 106

def test_shift_up():
    S('a1').set_value(1).drag_to('a10').drag_to('g10')
    S('a4').set_value(102).drag_to('g4')

    S('a5:d5').shift_up().drag_to('d3')
    assert S('c3').value == 103

    S('e6:g6').shift_up(2).drag_to('g3')
    assert S('f3').value == 106

def test_shifting_works_with_cell_contents():
    S('a1').set_value(10).shift_right().set_value(12).shift_down().set_value(15).shift_left().set_value(17)
    assert S('a2').value == 17

def test_selector_can_be_expanded():
    assert str(S('d4').grow_right()).endswith('.D4:E4')
    assert str(S('d4').grow_right(2)).endswith('.D4:F4')
    assert str(S('d4').grow_left()).endswith('.C4:D4')
    assert str(S('d4').grow_left(2)).endswith('.B4:D4')
    assert str(S('d4').grow_down()).endswith('.D4:D5')
    assert str(S('d4').grow_down(2)).endswith('.D4:D6')
    assert str(S('d4').grow_up()).endswith('.D3:D4')
    assert str(S('d4').grow_up(2)).endswith('.D2:D4')

    assert str(S('d4:e5').grow_right(2).grow_left(2).grow_down(2).grow_up(2)).endswith('.B2:G7')

def test_selector_can_be_reduced():
    assert str(S('b2:g7').shrink_right()).endswith('.B2:F7')
    assert str(S('b2:g7').shrink_right(2)).endswith('.B2:E7')
    assert str(S('b2:g7').shrink_left()).endswith('.C2:G7')
    assert str(S('b2:g7').shrink_left(2)).endswith('.D2:G7')
    assert str(S('b2:g7').shrink_down()).endswith('.B2:G6')
    assert str(S('b2:g7').shrink_down(2)).endswith('.B2:G5')
    assert str(S('b2:g7').shrink_up()).endswith('.B3:G7')
    assert str(S('b2:g7').shrink_up(2)).endswith('.B4:G7')

    assert str(S('B2:G7').shrink_right(2).shrink_left(2).shrink_down(2).shrink_up(2)).endswith('.D4:E5')

def test_object_can_be_cloned():
    start = S('a1')
    end = S('a1').clone().shift_right()

    assert str(start).endswith('.A1')
    assert str(end).endswith('.B1')
