
.. _macros:

=============
Python Macros
=============

To make a macro, create a python script and create a function in global scope for each routine you want to call from your document. Since the same code can run outside OpenOffice.org, you're advised to make a script that will work in standalone context. The template below might be a good start:

  #!/usr/bin/ipython

  from oosheet import OOSheet as S

  def my_macro():
      # do something
      pass

  def my_other_macro(): 
      pass

Using ipython will give you a python shell when you run you script from command line. Do this to test your script. When you're done, you have three options to run the macro from OpenOffice.org:

- Put your script in the global python scripts path. This is some directory like /usr/lib/openoffice/basis3.2/share/Scripts/python/. The macro you be available to all users in this computer.
- Put your script in the local python scripts path of your user. The path is something like ~/.openoffice.org/3/user/Scripts/python
- Pack the macro in the document. Details below.

If you choose one of the first two methods, which are simpler, you won't need any security configurations and will be able to run the macro for several documents. The third method makes more sense if your script logic is tied to an specific document structure and/or if you want to have stuff like buttons that trigger your macros. 

In any of these methods, you can run you macro from Tools -> Macros -> Run macro menu.

Packing your script in document
===============================

If you go to Tools -> Macros -> Organize macros -> Python, you'll notice that the "Create", "Edit", "Rename" and "Remove" options are disabled. This is because OO.org does not support managing your macros yet. To solve this, OOSheet comes with a command-line tool to pack your script in document. Just type:

  oosheet-pack my_document.ods my_script.py

When you open the document, you'll be warned that the document contains macros and that this is a security issue. So, you have to go to Tools -> Options -> Security -> Macro Security and configure it properly. It's a smarty thing to leave the security level at least "High".


