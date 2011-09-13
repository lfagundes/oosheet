
.. _working-with-any-document-type:

===================================================
Working with documents types other than spreadsheet
===================================================

Although OOSheet high-level API is developed for spreadsheets, it's base class **OODoc** will make
macro development job much easier for other types of document. The facility of testing your code via socket
and then packing with oosheet's packing tool is the same. And OODoc's API for dispatching events is much 
simpler than OpenOffice.org Basic code.

.. _recording-macros:

Recording macros
================

First thing you want is to know what kind of OpenOffice.org events will do what you need to do. For that,
you can record a macro in OpenOffice.org Basic.

For recording a macro, go to **Tools -> Macros -> Record Macro**. A recording dialog will open, and every
action will do will be recorded. Do some actions you'd like to reproduce later, then click the "stop recording"
button in the recording dialog. You'll need to select a name for saving this macro. After saving, go to
**Tools -> Macros -> Organize Macros... -> OpenOffice.org Basic**, find the macro you just saved and click
in "Edit". Check the code you just recorded.

The following code is an example of OpenOffice.org Basic macro for typing "Hello World" and then centering
it in Writer:

::

    document   = ThisComponent.CurrentController.Frame 
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

    rem ----------------------------------------------------------------------
    dim args1(0) as new com.sun.star.beans.PropertyValue
    args1(0).Name = "Text"
    args1(0).Value = "Hello World"
    dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args1())

    rem ----------------------------------------------------------------------
    dim args2(0) as new com.sun.star.beans.PropertyValue
    args2(0).Name = "CenterPara"
    args2(0).Value = true
    dispatcher.executeDispatch(document, ".uno:CenterPara", "", 0, args2())


The following Python code would do exactly the same thing, with OODoc:

    >>> from oosheet import OODoc
    >>> doc = OODoc()
    >>> doc.dispatch('InsertText', ('Text', 'Hello World'))
    >>> doc.dispatch('CenterPara', True)

Note that the args1, passed in the first dispatcher.executeDispatch() call, is substituted for a
tupple (Name, Value) in Python code, while the args2 value can be represented by just a boolean.
The tuple can be substituted by a single value when the Name of the property is the same as the command
being executed.

The same will work with OOSheet:

    >>> from oosheet import OOSheet as S
    >>> S().dispatch('AutomaticCalculation', False)








