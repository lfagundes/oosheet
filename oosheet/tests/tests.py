# -*- coding: utf-8 -*-

"""
Each test must start with test_

clear() is called between each test

Following parameters can be passed to run_tests.py:

  --dev    Only tests with @dev decorator will be executed
  --stop   Errors are raised

If no errors are encountered, all tests will be merged to a test document to be
run as macro.

"""

def clear():
    S('a1:z100').delete()
    S('Sheet2.a1:g10').delete()


def test_column_name_vs_index_conversion():
    assert S()._col_index('A') == 0
    assert S()._col_index('B') == 1
    assert S()._col_index('Z') == 25
    assert S()._col_index('AA') == 26

    assert S()._col_name(0) == 'A'
    assert S()._col_name(1) == 'B'
    assert S()._col_name(25) == 'Z'
    assert S()._col_name(26) == 'AA'

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

def test_drag_calls_can_be_cascaded():
    S('a1').value = 1
    S('a1').drag_to('a5').drag_to('c5')
    assert S('c5').value == 7

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

def test_selector_handles_sheets():
    S('a1').value = 2
    S('Sheet2.a1').value = 5

    assert S('a1').value == 2
    assert S('Sheet2.a1').value == 5

    S('Sheet2.a2').value = 3
    S('Sheet2.a1:a2').drag_to('b2')

    assert S('Sheet2.b1').value == 6
    assert S('Sheet2.b2').value == 4

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

def test_find_last_column():
    S('a1').set_value(1).drag_to('g1')

    S('b1').find_last_column().value = 100
    assert S('g1').value == 100

def test_find_last_column_works_with_ranges():
    S('g1').set_value(100).drag_to('g3')
    S('a1').set_value(1).drag_to('a3').drag_to('f3')
    
    S('b1:3').find_last_column().drag_to('i3')
    assert S('i2').value == 103

def test_find_last_column_may_consider_specific_row():
    S('a1').set_value(1).drag_to('a5').drag_to('g5')
    S('g3').delete()
    S('f1').set_value(100).drag_to('f5')

    S('a1:5').find_last_column(3).drag_to('g5')
    assert S('g1').value == 101
    assert S('g3').value == 103
    assert S('g5').value == 105
    

def test_find_last_row():
    S('a1').set_value(1).drag_to('a10')

    S('a2').find_last_row().value = 100
    assert S('a10').value == 100

def test_find_last_row_works_with_ranges():
    S('a10').set_value(100).drag_to('c10')
    S('a1').set_value(1).drag_to('c1').drag_to('c9')

    S('a2:c2').find_last_row().drag_to('c12')
    assert S('b12').value == 103

def test_find_last_row_may_consider_specific_column():
    S('a1').set_value(1).drag_to('e1').drag_to('e5')
    S('c5').delete()
    S('a4').set_value(100).drag_to('e4')

    S('a1:e1').find_last_row('c').drag_to('e6')
    assert S('a6').value == 102
    assert S('c6').value == 104
    assert S('e6').value == 106

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
    assert S('b1').value == 17
