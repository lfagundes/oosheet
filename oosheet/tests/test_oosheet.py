#!/usr/bin/python

import subprocess, os, time, random, types, shutil, sys
from datetime import datetime, timedelta

import unittest

from oosheet import OOSheet as S, OOMerger

def dev(func):
    func.dev = True
    return func

def getarg(argname):
    try:
        return ('--%s' % argname) in sys.argv
    except AttributeError:
        return False
    
dev_only = getarg('dev')
stop_on_error = getarg('stop')

class OOCalcLauncher(object):

    TIMEOUT = 10

    def __init__(self, path = None):
        assert not self.running

        if path is None:
            os.system('oocalc -accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"')
        else:
            os.system('oocalc %s' % path)
                      
        now = time.time()
        while time.time() < now + self.TIMEOUT:
            try:
                S().model
                return
            except Exception:
                time.sleep(0.1)

    def quit(self):
        filename = '/tmp/%s.ods' % ''.join([ random.choice('abcdefghijklmnopqrstuvwxyz') for i in range(32) ])
        S().save_as(filename) #avoid the saving question
        S().quit()
        os.remove(filename)
        
    @property
    def pid(self):
        sub = subprocess.Popen('ps aux'.split(), stdout=subprocess.PIPE)
        sub.wait()
        processes = [ line for line in sub.stdout if 'soffice' in line ]
        try:
            return int(processes[0].split()[1])
        except IndexError:
            return None
        
    @property
    def running(self):
        if self.pid is None:
            return False

        
        return self.pid is not None


def clear():
    S('a1:z100').delete()
    S('Sheet2.a1:g10').delete()

def test_internal_routines():
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

def test_insert_column():
    S('a1').value = 10
    S('b2').formula = '=a1+5'

    S('b2').insert_column()

    assert S('c2').formula.lower() == '=a1+5'
    assert S('b2').value == 0
    assert S('b2').string == ''


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

def test_copy_and_paste():
    S('a1').value = 4
    S('a1').copy()
    S('b2').paste()

    assert S('a1').value == 4
    assert S('b2').value == 4

    S('a1').cut()
    S('c1').paste()

    assert S('a1').value == 0
    assert S('c1').value == 4

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


def tests():
    tests = []
    for name, method in globals().items():
        if type(method) is types.FunctionType and name.startswith('test_'):
            tests.append(method)

    return tests
            
def run_tests(event = None):
    ok = True
    for i, test in enumerate(tests()):
        try:
            dev = test.dev
        except AttributeError:
            dev = False

        if dev_only and not dev:
            continue
            
        if event:
            S('Tests.b%d' % (i+10)).string = test.__name__
        else:
            sys.stdout.write('%s... ' % test.__name__)

        clear()
        if stop_on_error:
            test.__call__()
            print 'OK'
        else:
            try:
                if event:
                    S('Tests.c%d' % (i+10)).string = 'OK'
                else:
                    print 'OK'
            except Exception, e:
                ok = False
                if event:
                    S('Tests.d%d' % (i+10)).string = e
                else:
                    print '%s: %s' % (type(e).__name__, e)

    if event:
        S('Tests.a1').focus()

    return ok
            
            
if __name__ == '__main__':
    calc = OOCalcLauncher()
    try:
        result = run_tests()
    finally:
        calc.quit()

    if result:
        testmodel = os.path.join(os.path.dirname(__file__), 'testing_sheet.ods')
        testsheet = os.path.join(os.path.dirname(__file__), 'test.ods')

        shutil.copy(testmodel, testsheet)

        OOMerger(testsheet, __file__).merge()

        time.sleep(1)
        calc = OOCalcLauncher(testsheet)
    

    
    
    



