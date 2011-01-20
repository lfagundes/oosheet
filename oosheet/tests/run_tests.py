#!/usr/bin/python

"""
This is a custom test_runner for OOSheet, designed to run same tests both by connecting to
OpenOffice.org by socket and as macro.

Check tests.py for instructions.
"""

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

try:
    tests_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tests.py')
except NameError:
    pass # we're inside macro

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

execfile(tests_file)

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
        time.sleep(0.5)
        calc.quit()

    if result:
        testmodel = os.path.join(os.path.dirname(__file__), 'testing_sheet.ods')
        testsheet = os.path.join(os.path.dirname(__file__), 'test.ods')

        shutil.copy(testmodel, testsheet)

        script_path = '/tmp/test_oosheet.py'
        script = open(script_path, 'w')
        for line in open(__file__):
            if not line.startswith('execfile'):
                script.write(line)
            else:
                script.write(open(tests_file).read())
        script.close()

        OOMerger(testsheet, script_path).merge()

        os.remove(script_path)

        time.sleep(1)
        calc = OOCalcLauncher(testsheet)
    

    
    
    



