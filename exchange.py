"""Library to facilitate moving code between files and open OpenOffice docs.

The main functions -- `push`, `pull`, `run` --
can be invoked from the CLI.
"""
import sys

import find_ooo
import connect, context, document, libraries, library

class Exchange:
    """Class of the main exchange object.

    Its initialization finds a local OpenOffice installation
    and attempts to connect to an instance of OOo on the given host/port.

    It retains the context, service manager, and desktop acquired
    from this connection for use by one or more of the exchange methods.
    """
    def __init__(self, host='localhost', port='2002', find_uno=find_ooo.find_uno):
        self.context = connect.get_context(host=host, port=port,
                                           find_uno=find_uno)
        self.smgr = self.context.getServiceManager()
        self.desktop = context.get_desktop(self.context, self.smgr)

    def invoke(self, doc_name, macro_name):
        doc, libs = context.resolve_doc_name(self.context, self.smgr,
                                             self.desktop, doc_name)
        document.invoke_macro(doc, macro_name)

    def push(self, doc_name, macro_name, source):
        """Pushes the module code for `macro_name` from `source`.

        Returns the updated document.

        >>> Exchange().push({doc_name: 'Untitled 1',
        ...                  macro_name: 'Standard.Fraggle.main',
        ...                  source: open('project/src/basic/fraggle.bas')})
        ... # doctest: +SKIP
        """
        # TODO: sort out the fact that the routine name is irrelevant here.
        #       The fact that it's required by parse_macro_name is vestigial.
        #       This routine should take the library and module names
        #       instead of the unparsed macro_name;
        #       parsing should happen at the command line layer.
        lib_name, mod_name, procedure = library.parse_macro_name(macro_name)
        doc, libs = context.resolve_doc_name(self.context, self.smgr,
                                             self.desktop, doc_name)
        lib = libraries.get_lib_by_name(libs, lib_name, 'write')
        library.update_module(source, lib, mod_name)
        return doc

    def pull(self, doc_name, macro_name):
        """Gets the module code for `macro_name` from `source`.

        Yields the lines of source.

        Example::
            {doc_name: 'Untitled 1', macro_name: 'Standard.Module1.main'}
        """
        # TODO: see TODO for `Exchange.push`.
        lib_name, mod_name, procedure = library.parse_macro_name(macro_name)
        doc, libs = context.resolve_doc_name(self.context, self.smgr,
                                             self.desktop, doc_name)
        lib = libraries.get_lib_by_name(libs, lib_name, 'read')
        return library.get_module_source(lib, mod_name)

    def update_and_run(self, doc_name, macro_name, source):
        """ Update module from filename and run `macro_name` in `doc_name`.

            If doc_name is 'application',
            the macro is updated in the application library
            and executed on the active document.
            Otherwise, it is stored in the document and invoked on that document.
        """
        doc = self.push(doc_name, macro_name, source)
        self.run(doc, macro_name)
