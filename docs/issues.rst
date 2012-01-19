======
Issues
======

Breakpoint issue
================

It's worth noticing that *ipdb.set_trace() does not work* when you use OOSheet. This is not an issue from this module, it happens in deeper and darker layers of python-uno. If you see an error like this:

  SystemError: 'pyuno runtime is not initialized, (the pyuno.bootstrap needs to be called before using any uno classes)'

it's probably because you have an ipdb breakpoint. Use *pdb* instead.

Connection issues
=================

In version 1.2 performance got much better when using sockets, by caching the conection. The drawback is that when connection is broken, script must be restarted.

Sometimes the initial connection takes a long time. This was not reported in OpenOffice.org, but with LibreOffice this is a bit common. Any clue on why this helps is very welcome.
