# -*- coding: utf-8 -*-

import uno, re, sys, os, zipfile
from datetime import datetime, timedelta

# http://codesnippets.services.openoffice.org/Office/Office.MessageBoxWithTheUNOBasedToolkit.snip
from com.sun.star.awt import WindowDescriptor
from com.sun.star.awt.WindowClass import MODALTOP
from com.sun.star.awt.VclWindowPeerAttribute import OK

class OODoc(object):

    @property
    def model(self):
        localContext = uno.getComponentContext()
        if sys.modules.get('pythonscript'):
            # We're inside openoffice macro
            ctx = localContext
        else:
            # We have to connect by socket
            resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
            ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
            
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
            
        return desktop.getCurrentComponent()

    @property
    def dispatcher(self):
        localContext = uno.getComponentContext()
        if sys.modules.get('pythonscript'):
            ctx = localContext
        else:
            resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
            ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )

        smgr = ctx.ServiceManager
        return smgr.createInstanceWithContext( "com.sun.star.frame.DispatchHelper", ctx)

    def args(self, *args):
        uno_struct = []

        for i, arg in enumerate(args):
            struct = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
            struct.Name = arg[0]
            struct.Value = arg[1]
            uno_struct.append(struct)

        return tuple(uno_struct)

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

    def __init__(self, selector = None):
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
            (start, end) = cells.split(':')
            if not re.match('^[A-Z]', end):
                col, row = self._position(start)
                end = ''.join([self._col_name(col), end])
            self.start_col, self.start_row = self._position(start)
            self.end_col, self.end_row = self._position(end)
        else:
            col, row = self._position(cells)
            self.start_col, self.end_col = col, col
            self.start_row, self.end_row = row, row

    @property
    def selector(self):
        start = '%s%d' % (self._col_name(self.start_col), self.start_row + 1)
        end = '%s%d' % (self._col_name(self.end_col), self.end_row + 1)
        return '%s.%s:%s' % (self.sheet.Name, start, end)

    @property
    def cell(self):
        assert self.start_col == self.end_col
        assert self.start_row == self.end_row
        return self.sheet.getCellByPosition(self.start_col, self.start_row)

    def __repr__(self):
        return self.selector

    def _position(self, descriptor):
        col = re.findall('^([A-Z]+)', descriptor)[0]
        row = descriptor[len(col):]
            
        col = self._col_index(col)
        row = int(row) - 1

        return col, row

    def _col_index(self, name):
        letters = [ l for l in name ]
        letters.reverse()
        index = 0
        power = 0
        for letter in letters:
            index += (1 + ord(letter) - ord('A')) * pow(ord('Z') - ord('A') + 1, power)
            power += 1
        return index - 1

    def _col_name(self, index):
        name = []
        letters = [ chr(ord('A')+i) for i in range(26) ]
        
        while index > 0:
            i = index % 26
            index = int(index/26) - 1
            name.append(letters[i])

        if index == 0:
            name.append('A')

        name.reverse()
        return ''.join(name)            

    @property
    def basedate(self):
        return datetime(1899, 12, 30)

    @property
    def value(self):
        assert self.cell is not None
        return self.cell.getValue()

    @value.setter
    def value(self, value):
        assert self.cell is not None
        self.cell.setValue(value)

    def set_value(self, value):
        self.value = value
        return self

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

    def set_formula(self, formula):
        self.formula = formula
        return self

    @property
    def string(self):
        assert self.cell is not None
        return self.cell.getString()

    @string.setter
    def string(self, string):
        assert self.cell is not None
        self.cell.setString(string)

    def set_string(self, string):
        self.string = string
        return self

    @property
    def date(self):
        assert self.cell is not None
        return self.basedate + timedelta(self.value)

    @date.setter
    def date(self, date):
        assert self.cell is not None
        delta = date - self.basedate
        self.value = delta.days

    def set_date(self, date):
        self.date = date
        return self

    def focus(self):
        self.dispatch('.uno:GoToCell', ('ToPoint', self.selector))

    def drag_to(self, destiny):

        if '.' in destiny:
            sheet_name, destiny = destiny.split('.')
            assert sheet_name == self.sheet.Name
            
        self.focus()
        self.dispatch('.uno:AutoFill', ('EndCell', '%s.%s' % (self.sheet.Name, destiny)))

        if '.' not in destiny:
            destiny = '.'.join([self.sheet.Name, destiny])

        destiny = OOSheet(destiny)
        self.start_col = min(self.start_col, destiny.start_col)
        self.start_row = min(self.start_row, destiny.start_row)
        self.end_col = max(self.end_col, destiny.end_col)
        self.end_row = max(self.end_row, destiny.end_row)

        return self

    def delete_rows(self):
        self.focus()
        self.dispatch('.uno:DeleteRows')

    def delete_columns(self):
        self.focus()
        self.dispatch('.uno:DeleteColumns')

    def insert_row(self):
        self.focus()
        self.dispatch('.uno:InsertRows')
        self.end_row += 1
        return self

    def insert_column(self):
        self.focus()
        self.dispatch('.uno:InsertColumns')
        self.end_col += 1
        return self

    def shift_right(self, num = 1):
        self.start_col += num
        self.end_col += num
        return self
    
    def shift_left(self, num = 1):
        self.start_col -= num
        self.end_col -= num
        assert self.start_col >= 0
        return self

    def shift_down(self, num = 1):
        self.start_row += num
        self.end_row += num
        return self
    
    def shift_up(self, num = 1):
        self.start_row -= num
        self.end_row -= num
        assert self.start_row >= 0
        return self
    
    def find_last_column(self, row = None):
        assert self.start_col == self.end_col

        col = self.start_col
        if row is None:
            row = self.start_row
        else:
            row -= 1
            
        assert row >= self.start_row and row <= self.end_row
        
        while True:
            col += 1
            cell = self.sheet.getCellByPosition(col, row)
            if cell.getValue() == 0 and cell.getString() == '' and cell.getFormula() == '':
                col -= 1
                break
            
        cells = '%s%d' % (self._col_name(col), self.start_row+1)
        if self.end_row != self.start_row:
            cells += ':%d' % (self.end_row+1)
        selector = '.'.join([self.sheet.Name, cells])
        return OOSheet(selector)

    def find_last_row(self, col = None):
        assert self.start_row == self.end_row

        row = self.start_row
        if col is None:
            col = self.start_col
        else:
            col = self._col_index(col.upper())

        assert col >= self.start_col and col <= self.end_col
        while True:
            row += 1
            cell = self.sheet.getCellByPosition(col, row)
            if cell.getValue() == 0 and cell.getString() == '' and cell.getFormula() == '':
                row -= 1
                break

        cells = '%s%d' % (self._col_name(self.start_col), row+1)
        if self.end_col != self.start_col:
            cells += ':%s%d' % (self._col_name(self.end_col), row+1)
        selector = '.'.join([self.sheet.Name, cells])
        return OOSheet(selector)

    def copy(self):
        self.focus()
        self.dispatch('.uno:Copy')
        return self

    def cut(self):
        self.focus()
        self.dispatch('.uno:Cut')
        return self

    def paste(self):
        self.focus()
        self.dispatch('.uno:Paste')
        return self

    def delete(self):
        self.focus()
        self.dispatch('.uno:Delete', ('Flags', 'A'))

    def format_as(self, selector):
        OOSheet(selector).copy()
        self.focus()
        self.dispatch('.uno:InsertContents',
                      ('Flags', 'T'),
                      ('FormulaCommand', 0),
                      ('SkipEmptyCells', False),
                      ('Transpose', False),
                      ('AsLink', False),
                      ('MoveMode', 4),
                      )

        self.dispatch('.uno:TerminateInplaceActivation')
        self.dispatch('.uno:Cancel')

    def undo(self):
        self.dispatch('.uno:Undo')

    def redo(self):
        self.dispatch('.uno:Redo')

    def save_as(self, filename):
        if not filename.startswith('/'):
            filename = os.path.join(os.environ['PWD'], filename)
            
        self.dispatch('.uno:SaveAs', ('URL', 'file://%s' % filename), ('FilterName', 'calc8'))
        
    def quit(self):
        self.dispatch('.uno:Quit')


class OOMerger():

    def __init__(self, ods, script):
        self.ods = zipfile.ZipFile(ods, 'a')
        self.script = script

        assert os.path.exists(script)

    @property
    def script_name(self):
        return self.script.rpartition('/')[2]

    def manifest_add(self, path):
        manifest = []
        for line in self.ods.open('META-INF/manifest.xml'):
            if '</manifest:manifest>' in line:
                manifest.append(' <manifest:file-entry manifest:media-type="application/binary" manifest:full-path="%s"/>' % path)
            elif ('full-path:"%s"' % path) in line:
                return
            
            manifest.append(line)

        self.ods.writestr('META-INF/manifest.xml', ''.join(manifest))
        

    def merge(self):
        self.ods.write(self.script, 'Scripts/python/%s' % self.script_name)
        
        self.manifest_add('Scripts/')
        self.manifest_add('Scripts/python/')
        self.manifest_add('Scripts/python/%s' % self.script_name)

        self.ods.close()

def merge():
    print "Hello"
        
        

        
        



