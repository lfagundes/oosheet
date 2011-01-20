# -*- coding: utf-8 -*-

import uno, re, sys, os, zipfile, types
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
        

    def alert(self, msg, title = u'Alert'):
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
        return self._generate_selector(self.start_col, self.end_col,
                                       self.start_row, self.end_row)
    
    def _generate_selector(self, start_col, end_col, start_row, end_row):
        start = '%s%d' % (self._col_name(start_col), start_row + 1)
        end = '%s%d' % (self._col_name(end_col), end_row + 1)
        if start != end:
            return '%s.%s:%s' % (self.sheet.Name, start, end)
        else:
            return '%s.%s' % (self.sheet.Name, start)

    @property
    def cell(self):
        assert self.start_col == self.end_col
        assert self.start_row == self.end_row
        return self.sheet.getCellByPosition(self.start_col, self.start_row)

    @property
    def cells(self):
        for col in range(self.start_col, self.end_col+1):
            for row in range(self.start_row, self.end_row+1):
                yield self.sheet.getCellByPosition(col, row)

    def __repr__(self):
        return self.selector

    def _position(self, descriptor):
        col = re.findall('^([A-Z]+)', descriptor)[0]
        row = descriptor[len(col):]
            
        col = self._col_index(col)
        row = int(row) - 1

        return col, row

    def _col_index(self, name):
        letters = [ l for l in name.upper() ]
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
        for cell in self.cells:
            cell.setValue(value)

    def set_value(self, value):
        self.value = value
        return self

    @property
    def formula(self):
        assert self.cell is not None
        return self.cell.getFormula()

    @formula.setter
    def formula(self, formula):
        if not formula.startswith('='):
            formula = '=%s' % formula
        for cell in self.cells:
            cell.setFormula(formula)

    def set_formula(self, formula):
        self.formula = formula
        return self

    @property
    def string(self):
        assert self.cell is not None
        return self.cell.getString()

    @string.setter
    def string(self, string):
        for cell in self.cells:
            cell.setString(string)

    def set_string(self, string):
        self.string = string
        return self

    @property
    def date(self):
        assert self.cell is not None
        return self.basedate + timedelta(self.value)

    @date.setter
    def date(self, date):
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
        
    def shift(self, col, row):
        self.start_col += col
        self.end_col += col
        self.start_row += row
        self.end_row += row

        assert self.start_col >= 0
        assert self.start_row >= 0

        return self

    def shift_right(self, num = 1):
        return self.shift(num, 0)
    def shift_left(self, num = 1):
        return self.shift(-num, 0)
    def shift_down(self, num = 1):
        return self.shift(0, num)
    def shift_up(self, num = 1):
        return self.shift(0, -num)

    def _cell_matches(self, cell, value):
        assert type(value) in (types.NoneType, types.StringType, types.FloatType, types.IntType, datetime)

        if type(value) is types.StringType:
            return cell.getString() == value
        if type(value) in (types.FloatType, types.IntType):
            return cell.getValue() == value
        if type(value) is datetime:
            return cell.getValue() == (value - self.basedate).days

        # value is None
        return cell.getValue() == 0 and cell.getString() == '' and cell.getFormula() == ''
        
    def shift_until(self, col, row, *args, **kwargs):
        assert col != 0 or row != 0
        
        try:
            value = args[0]
            assert self.cell is not None
            while not self._cell_matches(self.cell, value):
                self.shift(col, row)
            return self
        except IndexError:
            pass

        assert len(kwargs.keys()) == 1
        ref = kwargs.keys()[0]
        value = kwargs[ref]

        reftype, position = ref.split('_')[:2]

        if ref.endswith('_satisfies'):
            condition = value
        else:
            condition = lambda s: s._cell_matches(s.cell, value)

        assert reftype in ('row', 'column')

        if reftype == 'row':
            assert row == 0
            ref_row = int(position) - 1
            if col > 0:
                ref_col = self.end_col
            else:
                ref_col = self.start_col
        else:
            assert col == 0
            ref_col = self._col_index(position)
            if row > 0:
                ref_row = self.end_row
            else:
                ref_row = self.start_row

        cell = OOSheet(self._generate_selector(ref_col, ref_col, ref_row, ref_row))
        while not condition(cell):
            self.shift(col, row)
            cell.shift(col, row)

        return self            

    def shift_right_until(self, *args, **kwargs):
        return self.shift_until(1, 0, *args, **kwargs)
    def shift_left_until(self, *args, **kwargs):
        return self.shift_until(-1, 0, *args, **kwargs)
    def shift_down_until(self, *args, **kwargs):
        return self.shift_until(0, 1, *args, **kwargs)
    def shift_up_until(self, *args, **kwargs):
        return self.shift_until(0, -1, *args, **kwargs)

    def grow(self, col, row):
        if col < 0:
            self.start_col += col
        else:
            self.end_col += col
        if row < 0:
            self.start_row += row
        else:
            self.end_row += row

        return self

    def grow_right(self, num = 1):
        return self.grow(num, 0)
    def grow_left(self, num = 1):
        return self.grow(-num, 0)
    def grow_up(self, num = 1):
        return self.grow(0, -num)
    def grow_down(self, num = 1):
        return self.grow(0, num)

    def shrink(self, col, row):
        if col < 0:
            self.start_col -= col
        else:
            self.end_col -= col
        if row < 0:
            self.start_row -= row
        else:
            self.end_row -= row

        assert self.start_row <= self.end_row
        assert self.start_col <= self.end_col

        return self

    def shrink_right(self, num = 1):
        return self.shrink(num, 0)
    def shrink_left(self, num = 1):
        return self.shrink(-num, 0)
    def shrink_up(self, num = 1):
        return self.shrink(0, -num)
    def shrink_down(self, num = 1):
        return self.shrink(0, num)

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
    try:
        document = sys.argv[1]
        script = sys.argv[2]
    except IndexError:
        print_help()

    if not os.path.exists(document):
        sys.stderr.write("%s not found" % document)
        print_help()

    if not os.path.exists(script):
        sys.stderr.write("%s not found" % script)
        print_help()

    OOMerger(document, script).merge()

def print_help():
    script_name = sys.argv[0].split('/')[-1]
    print "Usage: %s document.ods script.py" % script_name
    sys.exit(1)

    


        
        

        
        



