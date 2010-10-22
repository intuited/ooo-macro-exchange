`ooo-macro-exchange`
====================

Routines and CLI to facilitate injection/extraction of OpenOffice.org macros.

Works by connecting to a live instance of OOo.

It does not appear to be possible to use such a system to access Python macros
in an OpenOffice document.  This module only deals with Basic macros.


Usage
-----

Basic functions are

-   ``pull``: output a module's code or save it to a file
-   ``push``: replace a module's code with lines from a file or stdin
-   ``invoke``: run a macro.

Options etc. can be discovered by running ``python __main__ -h``,
or by installing and running ``ooo-macro-exchange -h``.


Relationship to other modules
-----------------------------

The structure of the code is pretty similar
to that used by `openoffice-python`_.

There are vague plans to integrate this module's functionality with it.

The code in this module was originally derived from this `forum post`_.

.. _openoffice-python: http://pypi.python.org/pypi/openoffice-python/
.. _forum post: http://www.oooforum.org/forum/viewtopic.phtml?t=94349#356461
