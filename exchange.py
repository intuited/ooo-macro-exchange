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
    def __init__(self, host='localhost', port='2002',
                       find_uno=find_ooo.find_uno):
        self.context = connect.get_context(host=host, port=port,
                                           find_uno=find_uno)
        self.smgr = self.context.getServiceManager()
        self.desktop = context.get_desktop(self.context, self.smgr)

    def invoke(self, doc_name, macro_name):
        """Invoke the macro in the running OOo instance.

        `macro_name` should be a fully-qualified macro name,
        for example 'Standard.Module1.main'.
        """
        doc, libs = context.resolve_doc_name(self.context, self.smgr,
                                             self.desktop, doc_name)
        document.invoke_macro(doc, macro_name)

    def push(self, doc_name, library_name, module_name, source):
        """Pushes the module code for `macro_name` from `source`.

        Returns the updated document.

        >>> Exchange().push(doc_name='Untitled 1',
        ...                 library_name='Standard', module_name='Fraggle',
        ...                 source=open('project/src/basic/fraggle.bas'))
        ... # doctest: +SKIP
        """
        # TODO: sort out the fact that the routine name is irrelevant here.
        #       The fact that it's required by parse_macro_name is vestigial.
        #       This routine should take the library and module names
        #       instead of the unparsed macro_name;
        #       parsing should happen at the command line layer.
        doc, libs = context.resolve_doc_name(self.context, self.smgr,
                                             self.desktop, doc_name)
        lib = libraries.get_lib_by_name(libs, library_name, 'write')
        library.update_module(source, lib, module_name)
        return doc

    def pull(self, doc_name, library_name, module_name):
        """Gets the module code for `macro_name` from `source`.

        Yields the lines of source.

        >>> Exchange().pull(doc_name: 'Untitled 1',
        ...                 library_name: 'Standard',
        ...                 module_name: 'Module1')
        ... # doctest: +SKIP
        ['sub main',
         '    MsgBox("This is the main macro in Standard.Module1.")',
         'end sub']
        """
        # TODO: see TODO for `Exchange.push`.
        doc, libs = context.resolve_doc_name(self.context, self.smgr,
                                             self.desktop, doc_name)
        lib = libraries.get_lib_by_name(libs, library_name, 'read')
        return library.get_module_source(lib, module_name)
