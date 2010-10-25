``ooo-macro-exchange``
======================

Routines and CLI to facilitate injection/extraction of OpenOffice.org macros.

Works by connecting to a live instance of OOo.

It does not appear to be possible to use such a system to access Python macros
in an OpenOffice document.  This module only deals with Basic macros.

There has also been some effort made to "Pythonify" the uno interface,
for example, by adding sequence and mapping proxy classes
for some of the uno container interfaces.


Dependencies
------------

You will need to have the py-uno bridge set up.

On Debian-ish systems like Ubuntu and Linux Mint,
this can generally be installed via ::

    $ sudo aptitude install python-uno


Installation
------------

Installation can be done via `PyPI`_,
or by fetching and installing the source code from the `github repo`_.

.. _PyPI: http://pypi.python.org/ooo-macro-exchange
.. _github repo: http://github.com/intuited/ooo-macro-exchange


Environment
-----------

The ``uno`` Python module will normally be imported without problems
on systems which have the ``python-uno`` module properly installed.

If the ``uno`` Python module cannot be imported,
``ooo-macro-exchange`` will try to add it to the path.

If the environment variable ``PY_UNO_PATH`` is defined,
``ooo-macro-exchange`` will add it to the python path.

This can also be accomplished by setting ``PYTHONPATH``.

If ``PY_UNO_PATH`` is unset,
``ooo-macro-exchange`` will try the path
``/usr/lib/openoffice/basis3.2/program/``.


Usage
-----

Basic functions are

-   ``pull``: output a module's code or save it to a file
-   ``push``: replace a module's code with lines from a file or stdin
-   ``invoke``: run a macro.

The ``oomax`` script provides command-line access to these actions.  E.G.::

    $ oomax pull 'Document 1' Standard.Module1
    sub Main()
        ' Content of main macro
    end sub

    sub AnotherMacro()
        ' Content of some other macro
    end sub

Note that ``oomax push`` will not save the document
unless the ``-s`` option is passed.
Instead, it will mark both the document and its libraries as modified.
This will cause the "Save" icon on the main document toolbar to activate.

Additional options can be discovered by running ``oomax -h``.

`pull`, `push`, and `invoke` are also available as methods
of the class `oomax.Exchange`.


Relationship to other modules
-----------------------------

The structure of the code is pretty similar
to that used by `openoffice-python`_.
There are vague plans to integrate this module's functionality with it.

The code in this module was originally derived from this `forum post`_.

.. _openoffice-python: http://pypi.python.org/pypi/openoffice-python/
.. _forum post: http://www.oooforum.org/forum/viewtopic.phtml?t=94349#356461


License
-------

``ooo-macro-exchange`` is licensed under the `FreeBSD License`_.
See the file COPYING for details.

.. _FreeBSD License: http://www.freebsd.org/copyright/freebsd-license.html
