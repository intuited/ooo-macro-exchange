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
    pass


# http://udk.openoffice.org/common/man
#       /concept/uno_contexts.html#current_context
def get_context(uno=uno, host='localhost', port=2002):
    """Returns a resolved connection context."""
    localctx = uno.getComponentContext()
    create_instance = localctx.getServiceManager().createInstanceWithContext

    resolver_class = "com.sun.star.bridge.UnoUrlResolver"
    resolver = create_instance(resolver_class, localctx)

    # http://udk.openoffice.org/common/man/spec/uno-url.html
    # http://wiki.services.openoffice.org/wiki
    #       /Documentation/DevGuide/ProUNO/Component_Context
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


def update_module(lines, lib, mod_name):
    """Update the named module in the given library with ``lines``.

    ``lines`` should be an iterator over strings.
    """
    contents = ''.join(lines)

    if not lib.hasByName(mod_name):
        lib.insertByName(mod_name, contents)
    else:
        lib.replaceByName(mod_name, contents)


def parse_macro_name(name):
    """Splits a macro name into its requisite 3 components.

    The components are (library, module, function).
    For example: 'Standard.Module1.main'.
    """
    parts = name.split('.')
    if len(parts) != 3:
        raise IllegalMacroNameError(name)
    return parts


# Not sure if there's any point in parametrizing these functions.
# The idea is to allow their results to be memoized in some manner,
# perhaps by virtue of being properties of an object.
# I'm mildly concerned that they could return inconsistent results
# if the OOo app changes document while this macro is running.
# Memoizing would allow consistent results to be returned
# over the scope of a "transaction".
def resolve_doc_name(context, service_manager, desktop, doc_name,
                     get_current_doc=get_current_doc, get_app_lib=get_app_lib,
                     get_doc_lib=get_doc_lib):
    """Returns (doc, libraries).

    If ``doc_name`` is "application",
    returns the current document and the app library.

    Raises DocLibLookupError if a document with the specified name
    is not found.

    The various get_* functions can be overridden.
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




    def run(self, doc, script_name):
        provider = doc.getScriptProvider()
        script = provider.getScript(script_name_url(script_name))
        return script.invoke((), (), ())




    def push(self, doc_name, macro_name, source):
        """Pushes the module code for ``macro_name`` from ``source``.

        Returns the updated document.
        """
        lib_name, mod_name, procedure = parse_macro_name(macro_name)
        doc, libraries = resolve_doc_name(self.ctx, self.smgr, self.desktop, doc_name)
        lib = get_lib_by_name(libraries, lib_name, 'write')
        update_module(source, lib, mod_name)
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
    ctx = get_context()
    file_name, macro_name, bas_name = parse_arg(sys.argv)
    with open(file_name, 'r') as source_file:
        Basic(ctx).update_and_run(source_file, macro_name, bas_name)
    #Basic(ctx).update_and_run('Untitled 1', 'Standard.Module1.main',
    #    '/home/asuka/Desktop/python/moduleA.bas')
