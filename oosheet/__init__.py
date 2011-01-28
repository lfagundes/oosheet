# -*- coding: utf-8 -*-

import uno, re, sys, os, zipfile, types
from datetime import datetime, timedelta

# http://codesnippets.services.openoffice.org/Office/Office.MessageBoxWithTheUNOBasedToolkit.snip
from com.sun.star.awt import WindowDescriptor
from com.sun.star.awt.WindowClass import MODALTOP
from com.sun.star.awt.VclWindowPeerAttribute import OK

class OODoc(object):
    """Interacts with any OpenOffice.org instance, not necessarily a Spreadsheet.
    This is the actual wrapper around python-uno.
    """
    @property
    def model(self):
        """Desktop's current component, a pyuno object of type com.sun.star.lang.XComponent.
        From this the document data can be manipulated.

        For example, to manipulate Sheet1.A1 cell through this:
        
        >>> OODoc().model.Sheets.getByIndex(0).getCellByPosition(0, 0)

        The current environment is detected to decide to connect either via socket or directly.
        
        """
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
        """A python-uno dispatcher object, of type com.sun.star.uno.XInterface
        From this user events can be simulated.

        For example, to focus on Sheet1.A1 through this:
        
        >>> doc = OODoc()
        >>> doc.dispatcher.executeDispatch(doc.model.getCurrentController(), '.uno:GoToCell', '', 0, doc.args(('ToPoint', 'Sheet1.A1')))
        
        The current environment is detected to decide to connect either via socket or directly.
        """
        localContext = uno.getComponentContext()
        if sys.modules.get('pythonscript'):
            ctx = localContext
        else:
            resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
            ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )

        smgr = ctx.ServiceManager
        return smgr.createInstanceWithContext( "com.sun.star.frame.DispatchHelper", ctx)

    def args(self, *args):
        """Receives a list of tupples and returns a list of com.sun.star.beans.PropertyValue objects corresponding to those tupples.
        This result can be passed to OODoc.dispatcher.
        """
        uno_struct = []

        for i, arg in enumerate(args):
            struct = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
            struct.Name = arg[0]
            struct.Value = arg[1]
            uno_struct.append(struct)

        return tuple(uno_struct)

    def dispatch(self, cmd, *args):
        """Combines OODoc.dispatcher and OODoc.args to dispatch a event.
        For example, to focus on Sheet1.A1:

        >>> OODoc().dispatch('.uno:GoToCell', ('ToPoint', 'Sheet1.A1'))
        
        """
        if args:
            args = self.args(*args)

        self.dispatcher.executeDispatch(self.model.getCurrentController(),
                                        cmd, '', 0, args)
        

    def alert(self, msg, title = u'Alert'):
        """Opens an alert window with a message and title, and requires user to click 'Ok'
        """
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
    """Interacts with an OpenOffice.org Spreadsheet instance.
    This high-level library works with a group of cells defined by a selector.
    """

    def __init__(self, selector = None):
        """Constructor gets a selector as parameter. Selector can be one of the following forms:
        a10
        a1:10
        a1:b3
        Sheet2.a10
        Sheet3.a1:10
        SheetX.a1:g10

        Selector is case-insensitive
        """
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
        """The selector used by an instance, in complete form. It will always in one of two forms:
        Sheet1.A1
        Sheet1.A1:A10

        Column labels will always be uppercase.        
        """
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
        """A python-uno com.sun.star.table.XCell object, representing one cell.
        Only works if selector is a single cell, otherwise raises AssertionError"""
        assert self.start_col == self.end_col
        assert self.start_row == self.end_row
        return self.sheet.getCellByPosition(self.start_col, self.start_row)

    @property
    def cells(self):
        """An generator of all cells of this selector. Each cell returned will be a
        python-uno com.sun.star.table.XCell object.
        """
        for col in range(self.start_col, self.end_col+1):
            for row in range(self.start_row, self.end_row+1):
                yield self.sheet.getCellByPosition(col, row)

    def __repr__(self):
        try:
            return self.selector
        except AttributeError:
            return 'empty OOSheet() object'

    @property
    def width(self):
        return self.end_col - self.start_col + 1

    @property
    def height(self):
        return self.end_row - self.start_row + 1

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
        """Hard-coded datetime.datetime object representing the date that corresponds to value 0"""
        return datetime(1899, 12, 30)

    @property
    def value(self):
        """The float value of a cell. Only works for single-cell selectors"""
        assert self.cell is not None
        return self.cell.getValue()

    @value.setter
    def value(self, value):
        """Sets the float value of all cells affected by this selector. Expects a float."""
        for cell in self.cells:
            cell.setValue(value)

    def set_value(self, value):
        """Sets the float value of all cells affected by this selector. Expects a float."""
        self.value = value
        return self

    @property
    def formula(self):
        """The formula of a cell. Only works for single-cell selectors"""
        assert self.cell is not None
        return self.cell.getFormula()

    @formula.setter
    def formula(self, formula):
        """Sets the formula of all cells affected by this selector. Expects a string"""
        if not formula.startswith('='):
            formula = '=%s' % formula
        for cell in self.cells:
            cell.setFormula(formula)

    def set_formula(self, formula):
        """Sets the formula of all cells affected by this selector. Expects a string"""
        self.formula = formula
        return self

    @property
    def string(self):
        """The string representation of a cell. Only works for single-cell selectors"""
        assert self.cell is not None
        return self.cell.getString()

    @string.setter
    def string(self, string):
        """Sets the string of all cells affected by this selector. Expects a string."""
        for cell in self.cells:
            cell.setString(string)

    def set_string(self, string):
        """Sets the string of all cells affected by this selector. Expects a string."""
        self.string = string
        return self

    @property
    def date(self):
        """The date representation of a cell. Only works for single-cell selectors"""
        assert self.cell is not None
        return self.basedate + timedelta(self.value)

    @date.setter
    def date(self, date):
        """Sets the date of all cells affected by this selector. Expects a datetime.datetime object."""
        delta = date - self.basedate
        self.value = delta.days

        date_format = uno.getConstantByName( "com.sun.star.util.NumberFormat.DATE" )
        formats = self.model.getNumberFormats()
        locale = uno.createUnoStruct( "com.sun.star.lang.Locale" )
        cells = self.sheet.getCellRangeByName(self.selector)
        cells.NumberFormat = formats.getStandardFormat( date_format, locale )


    def set_date(self, date):
        """Sets the date of all cells affected by this selector. Expects a datetime.datetime object."""
        self.date = date
        return self

    def focus(self):
        """Focuses on all cells of this selector"""
        self.dispatch('.uno:GoToCell', ('ToPoint', self.selector))

    def drag_to(self, destiny):
        """Focuses on cells and drag to the destiny specified by given selector, doing an AutoFill.
        Destiny is a selector string"""

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
        """Delete all rows that intersect with this selector"""
        self.focus()
        self.dispatch('.uno:DeleteRows')

    def delete_columns(self):
        """Delete all columns that intersect with this selector"""
        self.focus()
        self.dispatch('.uno:DeleteColumns')

    def insert_row(self):
        """Insert rows before this selector. The current selector is shift down, and expanded
        by one row, so the inserted row gets included in the resulting selector"""
        return self.insert_rows(1)

    def insert_rows(self, num):
        """Works as insert_row(), but inserts several rows"""
        self.focus()
        for i in range(num):
            self.dispatch('.uno:InsertRows')
        self.end_row += num
        return self

    def insert_column(self):
        """Insert one column before this selector. The current selector is shift right and expanded
        by one column, so the inserted column gets included in the resulting selector"""
        return self.insert_columns(1)

    def insert_columns(self, num):
        """Works as insert_column(), but inserts several columns"""
        self.focus()
        for i in range(num):
            self.dispatch('.uno:InsertColumns')
        self.end_col += num
        return self

    def copy(self):
        """Focuses and copies the contents, so it can be pasted somewhere else"""
        self.focus()
        self.dispatch('.uno:Copy')
        return self

    def cut(self):
        """Focuses and cuts the contents, they'll disappear and can be pasted somewhere"""
        self.focus()
        self.dispatch('.uno:Cut')
        return self

    def paste(self):
        """Focuses and pastes what's been copied or cut"""
        self.focus()
        self.dispatch('.uno:Paste')
        return self

    def delete(self):
        """Deletes the contents of cells in this selector"""
        self.focus()
        self.dispatch('.uno:Delete', ('Flags', 'A'))

    def format_as(self, selector):
        """Copies to the current selector the formmating of the given selector.
        Internally, copies the other selector and does a "paste special" in the current cells,
        pasting everything but data. No success has been achieved while trying to use the "brush" tool.
        """
        
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
        """Undo the last action"""
        self.dispatch('.uno:Undo')

    def redo(self):
        """Redo the last undo"""
        self.dispatch('.uno:Redo')

    def save_as(self, filename):
        """Saves the current doc to filename. Expects a string representing a path in filesystem.
        Path can be absolute or relative to PWD environment variable.
        """
        if not filename.startswith('/'):
            filename = os.path.join(os.environ['PWD'], filename)
            
        self.dispatch('.uno:SaveAs', ('URL', 'file://%s' % filename), ('FilterName', 'calc8'))
        
    def shift(self, col, row):
        """Moves the selector horizontally by a cell number given by "col" parameter and vertically
        by "row". Parameters can be negative to determine the direction of shift. Used internally
        by shift_right, shift_left, shift_up and shift_down methods.
        """
        self.start_col += col
        self.end_col += col
        self.start_row += row
        self.end_row += row

        assert self.start_col >= 0
        assert self.start_row >= 0

        return self

    def __add__(self, tup):
        assert type(tup) is types.TupleType
        assert len(tup) == 2
        assert type(tup[0]) in (types.IntType, types.FloatType)
        assert type(tup[1]) in (types.IntType, types.FloatType)
        return self.clone().shift(int(tup[0]), int(tup[1]))

    def __sub__(self, tup):
        assert type(tup) in (types.TupleType, type(self))
        if type(tup) is types.TupleType:
            assert len(tup) == 2
            assert type(tup[0]) in (types.IntType, types.FloatType)
            assert type(tup[1]) in (types.IntType, types.FloatType)
            return self + (-tup[0], -tup[1])
        else:
            assert self.width == tup.width
            assert self.height == tup.height
            return (self.start_col - tup.start_col, self.start_row - tup.start_row)

    def shift_right(self, num = 1):
        """Moves the selector to right, but number of columns given by "num" parameter."""
        return self.shift(num, 0)
    def shift_left(self, num = 1):
        """Moves the selector to left, but number of columns given by "num" parameter."""
        return self.shift(-num, 0)
    def shift_down(self, num = 1):
        """Moves the selector down, but number of rows given by "num" parameter."""
        return self.shift(0, num)
    def shift_up(self, num = 1):
        """Moves the selector up, but number of rows given by "num" parameter."""
        return self.shift(0, -num)

    def _cell_matches(self, cell, value):
        assert type(value) in (types.NoneType, types.StringType, types.UnicodeType, types.FloatType, types.IntType, datetime)

        if type(value) in (types.StringType, types.UnicodeType):
            return cell.getString() == value
        if type(value) in (types.FloatType, types.IntType):
            return cell.getValue() == value
        if type(value) is datetime:
            return cell.getValue() == (value - self.basedate).days

        # value is None
        return cell.getValue() == 0 and cell.getString() == '' and cell.getFormula() == ''
        
    def shift_until(self, col, row, *args, **kwargs):
        """Moves the selector in direction given by "col" and "row" parameters, until a condition is satisfied.
        If selector is a single cell, than a value can be given as parameter and shift will be done until
        that exact value is found.
        For multiple cells selectors, the parameters can be in one of the following forms:
           column_LABEL = value
           row_NUMBER = value
           column_LABEL_satisfies = lambda
           row_NUMBER_satisfies = lambda

        If column is given as condition, then shift must be horizontal, and vice-versa.
        If matching against a value, the type of the value given will be checked and either "value", "string"
        or "date" property of cell will be used.
        If matching against a lambda function, a python-uno com.sun.star.table.XCell object will be given
        as parameter to the lambda function.
        """
        
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
        """Moves selector to right until condition is matched. See shift_until()"""
        return self.shift_until(1, 0, *args, **kwargs)
    def shift_left_until(self, *args, **kwargs):
        """Moves selector to left until condition is matched. See shift_until()"""
        return self.shift_until(-1, 0, *args, **kwargs)
    def shift_down_until(self, *args, **kwargs):
        """Moves selector down until condition is matched. See shift_until()"""
        return self.shift_until(0, 1, *args, **kwargs)
    def shift_up_until(self, *args, **kwargs):
        """Moves selector up until condition is matched. See shift_until()"""
        return self.shift_until(0, -1, *args, **kwargs)

    def grow(self, col, row):
        """Expands the selector by sizes given by "col" and "row" parameter.
        If col is a positive number, columns will be added to right, if negative to left. Same for row.
        """        
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
        """Add columns to right of selector"""
        return self.grow(num, 0)
    def grow_left(self, num = 1):
        """Add columns to left of selector"""
        return self.grow(-num, 0)
    def grow_up(self, num = 1):
        """Add rows before selector"""
        return self.grow(0, -num)
    def grow_down(self, num = 1):
        """Add rows after selector"""
        return self.grow(0, num)

    def shrink(self, col, row):
        """Reduces the size of the selector, in same logic as grow()"""
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
        """Removes columns from right of selector. Does not afect data, only the selector."""
        return self.shrink(num, 0)
    def shrink_left(self, num = 1):
        """Removes columns from left of selector. Does not afect data, only the selector."""
        return self.shrink(-num, 0)
    def shrink_up(self, num = 1):
        """Removes rows from beginning of selector. Does not afect data, only the selector."""
        return self.shrink(0, -num)
    def shrink_down(self, num = 1):
        """Removes rows from end of selector. Does not afect data, only the selector."""
        return self.shrink(0, num)

    def clone(self):
        """Returns a clone of this selector.
        Useful to preserve a state before calls that modify the selector.
        """
        return OOSheet(self.selector)

    def quit(self):
        """Closes the OpenOffice.org instance"""
        self.dispatch('.uno:Quit')


class OOPacker():
    """This class manipulates a document in OpenDocument format (the one used by OpenOffice.org)
    to pack python scripts inside it. This is necessary because OpenOffice.org does not offer a way to
    edit Python scripts.
    """
    def __init__(self, ods, script):
        """"ods" and "script" parameters are strings containing the filename of the OpenDocument document and
        Python script, respectively.
        For now, expects the document to have .ods format.
        """
        self.ods = zipfile.ZipFile(ods, 'a')
        self.script = script

        assert os.path.exists(script)

    @property
    def script_name(self):
        """Gets the script name, ignoring the path of the file"""
        return self.script.rpartition('/')[2]

    def manifest_add(self, path):
        """Parses the META-INF/manifest.xml file inside the document and adds lines to include the
        Python script.
        """        
        manifest = []
        for line in self.ods.open('META-INF/manifest.xml'):
            if '</manifest:manifest>' in line:
                manifest.append(' <manifest:file-entry manifest:media-type="application/binary" manifest:full-path="%s"/>' % path)
            elif ('full-path:"%s"' % path) in line:
                return
            
            manifest.append(line)

        self.ods.writestr('META-INF/manifest.xml', ''.join(manifest))
        

    def pack(self):
        """Packs the Python script inside the document"""
        self.ods.write(self.script, 'Scripts/python/%s' % self.script_name)
        
        self.manifest_add('Scripts/')
        self.manifest_add('Scripts/python/')
        self.manifest_add('Scripts/python/%s' % self.script_name)

        self.ods.close()

def pack():
    """Command line to pack the script in a document. Acessed as "oosheet-pack"."""
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

    OOPacker(document, script).pack()

def print_help():
    """Prints help message for pack()"""
    script_name = sys.argv[0].split('/')[-1]
    print "Usage: %s document.ods script.py" % script_name
    sys.exit(1)

    


        
        

        
        



