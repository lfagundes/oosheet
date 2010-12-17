#!/usr/bin/python

import subprocess, os, time
from datetime import datetime, timedelta

from oosheet import OOSheet as S

class OOCalcLauncher(object):

    TIMEOUT = 10

    def __init__(self):
        assert not self.running

        os.system('oocalc -accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"')
                      
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


def run_tests():
    test_value()
    test_string()
    test_formula()

if __name__ == '__main__':
    calc = OOCalcLauncher()
    run_tests()
    calc.quit()

    calc = OOCalcLauncher()
    S('a1').string = 'Please run run_tests macro'
    S().save_as('/tmp/oosheet_test.ods')
    calc.quit()

    # TODO insert this script in oosheet_test.ods and open oocalc so tests can be ran in macro context
    
    



