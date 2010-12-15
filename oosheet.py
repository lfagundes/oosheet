# -*- coding: utf-8 -*-

import uno, re
from datetime import datetime, timedelta

# http://codesnippets.services.openoffice.org/Office/Office.MessageBoxWithTheUNOBasedToolkit.snip
from com.sun.star.awt import WindowDescriptor
from com.sun.star.awt.WindowClass import MODALTOP
from com.sun.star.awt.VclWindowPeerAttribute import OK

class OODoc(object):

    @property
    def model(self):
        try:
            return XSCRIPTCONTEXT.getDocument()
        except NameError:
            localContext = uno.getComponentContext()
            resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
            ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
            smgr = ctx.ServiceManager
            desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
            
            return desktop.getCurrentComponent()

    @property
    def dispatcher(self):
        try:
            ctx = XSCRIPTCONTEXT.getComponentContext()
        except NameError:
            localContext = uno.getComponentContext()
            resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
            ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )

        smgr = ctx.ServiceManager
        return smgr.createInstanceWithContext( "com.sun.star.frame.DispatchHelper", ctx)

    def args(self, *args):
        uno_struct = (uno.createUnoStruct('com.sun.star.beans.PropertyValue'),
                      uno.createUnoStruct('com.sun.star.beans.PropertyValue'),
                      uno.createUnoStruct('com.sun.star.beans.PropertyValue'))

        for i, arg in enumerate(args):
            uno_struct[i].Name = arg[0]
            uno_struct[i].Value = arg[1]

        return uno_struct

    def dispatch(self, cmd, *args):
        if args:
            args = self.args(*args)

        self.dispatcher.executeDispatch(self.model.getCurrentController(),
                                        cmd, '', 0, args)
        

    def alert(self, msg, title = u'Atenção'):
        parentWin = self.model.CurrentController.Frame.ContainerWindow

        aDescriptor = WindowDescriptor()
	aDescriptor.Type = MODALTOP
	aDescriptor.WindowServiceName = 'messbox'
	aDescriptor.ParentIndex = -1
	aDescriptor.Parent = parentWin
	aDescriptor.WindowAttributes = OK

        tk = parentWin.getToolkit()
        box = tk.createWindow(aDescriptor)

        box.setMessageText(msg)
        
        if title:
            box.setCaptionText(title)

        box.execute()

    
class OOSheet(OODoc):

    def __init__(self, selector):
        if not selector:
            return
        
        try:
            sheet_name, cells = selector.split('.')
            self.sheet = self.model.Sheets.getByName(sheet_name)
        except ValueError:
            self.sheet = self.model.Sheets.getByIndex(0)
            cells = selector
        cells.replace('$', '')
        cells = cells.upper()

        if ':' in cells:
            self.range = cells
            self.cell = None
        else:
            col, row = self._position(cells)
            self.cell = self.sheet.getCellByPosition(col, row)
            self.range = None

        self.selector = '.'.join([self.sheet.Name, cells])

    def _position(self, descriptor):
        col = re.findall('^([A-Z]+)', descriptor)[0]
        row = descriptor[len(col):]
            
        col = self.col_index(col)
        row = int(row) - 1

        return col, row
        

    def col_index(self, name):
        letters = [ l for l in name ]
        letters.reverse()
        index = 0
        power = 0
        for letter in letters:
            index += (1 + ord(letter) - ord('A')) * pow(ord('Z') - ord('A') + 1, power)
            power += 1
        return index - 1

    @property
    def basedate(self):
        return datetime(1899, 12, 30)

    def focus(self):
        self.dispatch('.uno:GoToCell', ('ToPoint', self.selector))

    def drag_to(self, destiny):

        if '.' in destiny:
            sheet_name, destiny = destiny.split('.')
            assert sheet_name == self.sheet.Name
            
        self.focus()

        self.dispatch('.uno:AutoFill', ('EndCell', '%s.%s' % (self.sheet.Name, destiny)))

    @property
    def value(self):
        assert self.cell is not None
        return self.cell.getValue()

    @value.setter
    def value(self, value):
        assert self.cell is not None
        self.cell.setValue(value)

    @property
    def formula(self):
        assert self.cell is not None
        return self.cell.getFormula()

    @formula.setter
    def formula(self, formula):
        assert self.cell is not None
        if not formula.startswith('='):
            formula = '=%s' % formula
        self.cell.setFormula(formula)

    @property
    def string(self):
        assert self.cell is not None
        return self.cell.getString()

    @string.setter
    def string(self, string):
        assert self.cell is not None
        self.cell.setString(string)

    @property
    def date(self):
        assert self.cell is not None
        return self.basedate + timedelta(self.value)

    @date.setter
    def date(self, date):
        assert self.cell is not None
        delta = date - self.basedate
        self.value = delta.days

    def delete_rows(self):
        self.focus()
        self.dispatch('.uno:DeleteRows')

    def delete_columns(self):
        self.focus()
        self.dispatch('.uno:DeleteColumns')

    def insert_row(self):
        self.focus()
        self.dispatch('.uno:InsertRows')

    def insert_column(self):
        self.focus()
        self.dispatch('.uno:InsertColumns')

    def copy(self):
        self.focus()
        self.dispatch('.uno:Copy')

    def cut(self):
        self.focus()
        self.dispatch('.uno:Cut')

    def paste(self):
        self.focus()
        self.dispatch('.uno:Paste')        
