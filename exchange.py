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

    def push(self, doc_name, library_name, module_name, source, save=False):
        """Pushes the module code for `macro_name` from `source`.

        If the library does not exist, it is created.

        If `save` is truthy, the updated document will be saved to disk.

        Returns the updated document.

        >>> Exchange().push(doc_name='Untitled 1',
        ...                 library_name='Standard', module_name='Fraggle',
        ...                 source=open('project/src/basic/fraggle.bas'))
        ... # doctest: +SKIP
        """
        # TODO: There should probably be some exception-handling in here.
        doc, libs = context.resolve_doc_name(self.context, self.smgr,
                                             self.desktop, doc_name)
        joined_source = '\n'.join(line.rstrip('\n') for line in source)
        try:
            libs[library_name][module_name] = joined_source
        except KeyError:
            libs.createLibrary(library_name)
            libs[library_name][module_name] = joined_source

        if save:
            doc.store()
        else:
            # There doesn't appear to be any way to cause an open Basic editor
            # to change its save icon to active.
            # That functionality seems to be semi-broken anyway,
            # so probably people will be expecting to save the document
            # in order to be sure.
            libs.setModified(True)
            doc.setModified(True)
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
        doc, libs = context.resolve_doc_name(self.context, self.smgr,
                                             self.desktop, doc_name)
        lib = libs[library_name][module_name]
        return (line + "\n" for line in lib.rstrip('\n').split('\n'))
