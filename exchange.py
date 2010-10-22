"""Library to facilitate moving code between files and open OpenOffice docs.

The main functions -- `push`, `pull`, `run` --
can be invoked from the CLI.
"""
import sys
import find_ooo

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


def get_context(host='localhost', port=2002, find_uno=find_ooo.find_uno):
    """Returns a resolved connection context.

    `find_uno` is called to acquire an `uno`:module: object.
    All subsequent calls to the UNO API will go through this object.

    Returns a `component context`, accessed via the `current context`_.

    .. _component context:
       http://wiki.services.openoffice.org/wiki
             /Documentation/DevGuide/ProUNO/Component_Context
    .. _current context:
       http://udk.openoffice.org
             /common/man/concept/uno_contexts.html#current_context
    """

    uno = find_uno()
    localctx = uno.getComponentContext()
    create_instance = localctx.getServiceManager().createInstanceWithContext

    resolver_class = "com.sun.star.bridge.UnoUrlResolver"
    resolver = create_instance(resolver_class, localctx)

    # http://udk.openoffice.org/common/man/spec/uno-url.html
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
    """Get the library named by `library_name`.

    `libraries` is a sequence of libraries,
    as returned by `get_libraries`.

    If `mode` is 'write',
    the library is checked for write access.
    """
    if not libs.hasByName(library_name):
        libs.createLibrary(library_name)

    if mode == 'write':
        if libs.isLibraryReadOnly(library_name):
            raise ReadonlyLibraryError(library_name)

    if (libs.isLibraryPasswordProtected(library_name)
        and not libs.isLibraryPasswordVerified(library_name)):
        raise PasswordProtectionError(library_name)

    if not libs.isLibraryLoaded(library_name):
        libs.loadLibrary(library_name)

    return libs.getByName(library_name)


def get_document(desktop, doc_name):
    """Returns the document corresponding to `doc_name`."""
    frames = desktop.getFrames()
    for i in range(frames.getCount()):
        controller = frames.getByIndex(i).getController()
        if controller and controller.getTitle() == doc_name:
            return controller.getModel()

    raise DocLibLookupError(doc_name)

def get_libraries(document):
    """Returns the libraries for the `document`."""
    return document.BasicLibraries


def get_module_source(lib, module_name):
    """Return a list of the source lines for the module `module_name` in `lib`.

    Lines do not have terminating newlines.
    """
    return lib.getByName(module_name).split("\n")

def update_module(lines, lib, mod_name):
    """Update the named module in the given library with `lines`.

    `lines` should be an iterator over strings.
    """
    contents = '\n'.join(line.rstrip("\n") for line in iter(lines))

    if not lib.hasByName(mod_name):
        lib.insertByName(mod_name, contents)
    else:
        lib.replaceByName(mod_name, contents)

def invoke_macro(doc, script_name):
    provider = doc.getScriptProvider()
    script = provider.getScript(script_name_url(script_name))
    return script.invoke((), (), ())


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
                     get_document=get_document):
    """Returns (doc, libraries).

    If `doc_name` is "application",
    returns the current document and the app library.

    Raises DocLibLookupError if a document with the specified name
    is not found.

    The various get_* functions can be overridden.
    """
    if doc_name == 'application':
        return (get_current_doc(desktop),
                get_app_lib(context, service_manager))
    doc = get_document(desktop, doc_name)
    return doc, get_libraries(doc)


def script_name_url(script_name):
    """Returns a URL for a script name like '
    """
    fmtstr = 'vnd.sun.star.script:{0}?language=Basic&location=document'
    return fmtstr.format(script_name)


class Exchange:
    """Class of the main exchange object.

    Its initialization finds a local OpenOffice installation
    and attempts to connect to an instance of OOo on the given host/port.

    It retains the context, service manager, and desktop acquired
    from this connection for use by one or more of the exchange methods.
    """
    def __init__(self, host='localhost', port='2002', find_uno=find_ooo.find_uno):
        self.context = get_context(host=host, port=port, find_uno=find_uno)
        self.smgr = self.context.getServiceManager()
        self.desktop = get_desktop(self.context, self.smgr)

    def invoke(self, doc_name, macro_name):
        doc, libraries = resolve_doc_name(self.context, self.smgr, self.desktop, doc_name)
        invoke_macro(doc, macro_name)

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
        lib_name, mod_name, procedure = parse_macro_name(macro_name)
        doc, libraries = resolve_doc_name(self.context, self.smgr, self.desktop, doc_name)
        lib = get_lib_by_name(libraries, lib_name, 'write')
        update_module(source, lib, mod_name)
        return doc

    def pull(self, doc_name, macro_name):
        """Gets the module code for `macro_name` from `source`.

        Yields the lines of source.

        Example::
            {doc_name: 'Untitled 1', macro_name: 'Standard.Module1.main'}
        """
        # TODO: see TODO for `Exchange.push`.
        lib_name, mod_name, procedure = parse_macro_name(macro_name)
        doc, libraries = resolve_doc_name(self.context, self.smgr, self.desktop, doc_name)
        lib = get_lib_by_name(libraries, lib_name, 'read')
        return get_module_source(lib, mod_name)

    def update_and_run(self, doc_name, macro_name, source):
        """ Update module from filename and run `macro_name` in `doc_name`.

            If doc_name is 'application',
            the macro is updated in the application library
            and executed on the active document.
            Otherwise, it is stored in the document and invoked on that document.
        """
        doc = self.push(doc_name, macro_name, source)
        self.run(doc, macro_name)
