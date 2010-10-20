# -*- encoding: utf-8 -*-

import sys
import find_ooo
sys.path.append(find_ooo.find_ooo())
import uno

class DocLibLookupError(Exception):
    """Raised if document name lookup fails."""
    pass

class ReadonlyLibraryError(Exception):
    """Raised if an attempt is made to update a readonly library."""
    pass

class PasswordProtectionError(Exception):
    """Raised if a password check prevents a library update."""
    pass

class IllegalMacroNameError(Exception):
    """Raised if a macro name with less or more than three parts is given."""


def connect(uno=uno, host='localhost', port=2002):
    """Returns a resolved connection context."""
    localctx = uno.getComponentContext()
    create_instance = localctx.getServiceManager().createInstanceWithContext

    resolver_class = "com.sun.star.bridge.UnoUrlResolver"
    resolver = create_instance(resolver_class, localctx)

    context_url = ("uno:socket,host={host},port={port};"
                   "urp;StarOffice.ComponentContext".format)

    return resolver.resolve(context_url(host=host, port=port))


def get_desktop(context, service_manager):
    create_instance = service_manager.createInstanceWithContext
    desktop_class = "com.sun.star.frame.Desktop"
    return create_instance(desktop_class, context)

def get_app_lib(context, service_manager):
    # TODO: find out if it's possible for an exception to be thrown
    #       by create_instance in a situation where
    #       there is legitimately no library for the application.
    create_instance = service_manager.createInstanceWithContext
    library_class = "com.sun.star.script.ApplicationScriptLibraryContainer"
    return create_instance(library_class, context)

def get_current_doc(desktop):
    return desktop.getCurrentComponent()

def get_lib_by_name(libs, library_name, mode='read'):
    """Get the library named by ``library_name``.

    ``libraries`` is a sequence of libraries,
    as returned by ``get_doc_lib``.

    If ``mode`` is 'write',
    the library is checked for write access.
    """
    if not libs.hasByName(lib_name):
        libs.createLibrary(lib_name)

    if mode == 'write':
        if libs.isLibraryReadOnly(lib_name):
            raise ReadonlyLibraryError(lib_name)

    if (libs.isLibraryPasswordProtected(lib_name)
        and not libs.isLibraryPasswordVerified(lib_name):
        raise PasswordProtectionError(lib_name)

    if not libs.isLibraryLoaded(lib_name):
        libs.loadLibrary(lib_name)

    return libs.getByName(lib_name)


def get_doc_lib(desktop, doc_name):
    """Returns (document, library) for ``doc_name``."""
    frames = desktop.getFrames()
    for i in range(frames.getCount()):
        controller = frames.getByIndex(i).getController()
        if controller and controller.getTitle() == doc_name:
            doc = controller.getModel()
            return doc, doc.BasicLibraries

    raise DocLibLookupError(doc_name)

def resolve_doc_name(context, service_manager, desktop, doc_name):
    """Returns (doc, libraries).

    If ``doc_name`` is "application",
    returns the current document and the app library.

    Raises DocLibLookupError if a document with the specified name
    is not found.
    """
    if doc_name == 'application':
        return (get_current_doc(desktop),
                get_app_lib(context, service_manager))
    return get_doc_lib(desktop, doc_name)


script_name_url = (
    'vnd.sun.star.script:{0}?language=Basic&location=document'.format)


class Basic:
    """ Update Basic module and excute subroutine. """
    def __init__(self, ctx):
        self.ctx = ctx
        self.smgr = ctx.getServiceManager()
        self.desktop = get_desktop(self.ctx, self.smgr)


    def update_module(self, lines, libs, lib_name, mod_name):
        """Update the named module in the named library."""
        lib = get_lib_by_name(libs, lib_name, 'write')

        contents = ''.join(lines)

        if not lib.hasByName(mod_name):
            lib.insertByName(mod_name, contents)
        else:
            lib.replaceByName(mod_name, contents)



    def run(self, doc, script_name):
        provider = doc.getScriptProvider()
        script = provider.getScript(script_name_url(script_name))
        return script.invoke((), (), ())



    def parse_name(self, name):
        """Parse macro name."""
        parts = name.split('.')
        if len(parts) != 3:
            raise IllegalMacroNameError(name)
        return parts


    def push(self, doc_name, macro_name, source):
        """Pushes the module code for ``macro_name`` from ``source``.

        Returns the updated document.
        """
        lib_name, mod_name, procedure = self.parse_name(macro_name)
        doc, libraries = resolve_doc_name(self.ctx, self.smgr, self.desktop, doc_name)
        self.update_module(source, libraries, lib_name, mod_name)
        return doc

    def update_and_run(self, doc_name, macro_name, source):
        """ Update module from filename and run ``macro_name`` in ``doc_name``.

            If doc_name is 'application',
            the macro is updated in the application library
            and executed on the active document.
            Otherwise, it is stored in the document and invoked on that document.
        """
        doc = self.push(doc_name, macro_name, source)
        self.run(doc, macro_name)


def parse_arg(args):
    if len(args) == 4:
        return args[1:]

if __name__ == '__main__':
    ctx = connect()
    file_name, macro_name, bas_name = parse_arg(sys.argv)
    with open(file_name, 'r') as source_file:
        Basic(ctx).update_and_run(source_file, macro_name, bas_name)
    #Basic(ctx).update_and_run('Untitled 1', 'Standard.Module1.main',
    #    '/home/asuka/Desktop/python/moduleA.bas')
