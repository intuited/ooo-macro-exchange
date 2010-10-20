# -*- encoding: utf-8 -*-

import sys
import find_ooo
sys.path.append(find_ooo.find_ooo())
import uno

class DocLibLookupError(Exception):
    """Raised if document name lookup fails."""
    pass

def connect():
    localctx = uno.getComponentContext()
    resolver = localctx.getServiceManager().createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localctx)
    return resolver.resolve(
        "uno:socket,host=localhost,port=2083;urp;StarOffice.ComponentContext")


class Basic:
    """ Update Basic module and excute subroutine. """
    def __init__(self, ctx):
        self.ctx = ctx
        self.smgr = ctx.getServiceManager()
        self.desktop = self.get_desktop()

    def get_desktop(self):
        create_instance = self.smgr.createInstanceWithContext
        desktop_class = "com.sun.star.frame.Desktop"
        return create_instance(desktop_class, self.ctx)

    def get_current_doc(self):
        return self.desktop.getCurrentComponent()

    def get_app_lib(self):
        # TODO: find out if it's possible for an exception to be thrown
        #       by create_instance in a situation where
        #       there is legitimately no library for the application.
        create_instance = self.smgr.createInstanceWithContext
        library_class = "com.sun.star.script.ApplicationScriptLibraryContainer"
        return create_instance(library_class, self.ctx)

    def get_doc_lib(self, doc_name):
        """Returns (doc, libraries).

        If ``doc_name`` is "application",
        returns the current document and the app library.

        Raises DocLibLookupError if the a document with the specified name
        is not found.
        """
        if doc_name == 'application':
            return (self.get_current_doc(),
                    self.get_app_lib(ctx))

        frames = self.desktop.getFrames()
        for i in range(frames.getCount()):
            controller = frames.getByIndex(i).getController()
            if controller and controller.getTitle() == doc_name:
                doc = controller.getModel()
                return doc, doc.BasicLibraries

        raise DocLibLookupError


    def update_module(self, file_name, libs, lib_name, mod_name):
        if not libs:
            raise RuntimeError("illegal library.")
        if not libs.hasByName(lib_name):
            libs.createLibrary(lib_name)
        if libs.isLibraryReadOnly(lib_name):
            raise RuntimeError("readonly library (%s)" % lib_name)
        if libs.isLibraryPasswordProtected(lib_name) and not libs.isLibraryPasswordVerified(lib_name):
            raise RuntimeError("library is protected by the password. (%s)" % lib_name)
        lib = libs.getByName(lib_name)
        if not libs.isLibraryLoaded(lib_name):
            libs.loadLibrary(lib_name)

        with open(file_name, 'r') as module_file:
            contents = module_file.read()

        if not lib.hasByName(mod_name):
            lib.insertByName(mod_name, contents)
        else:
            lib.replaceByName(mod_name, contents)
        return True

    def run(self, doc, name):
        ret = None
        sp = doc.getScriptProvider()
        try:
            script = sp.getScript('vnd.sun.star.script:%s?language=Basic&location=document' % name)
            if script:
                ret = script.invoke((), (), ())
        except Exception as e:
            print(e)
        return ret

    def parse_name(self, name):
        """parse macro name."""
        parts = name.split('.')
        if len(parts) != 3:
            raise RuntimeError("Illegal name: %s" % name)
        return parts


    def update_and_run(self, doc_name, macro_name, filename):
        """ Update module from filename and run ``macro_name`` in ``doc_name``.

            If doc_name is 'application',
            the macro is updated in the application library
            and executed on the active document.
            Otherwise, it is stored in the document and invoked on that document.
        """
        lib_name, mod_name, procedure = self.parse_name(macro_name)
        doc, libraries = self.get_doc_lib(doc_name)
        self.update_module(file_path, libraries, lib_name, mod_name):
        self.run(doc, macro_name)


def parse_arg(args):
    if len(args) == 4:
        return args[1:]

if __name__ == '__main__':
    ctx = connect()
    file_name, macro_name, bas_name = parse_arg(sys.argv)
    Basic(ctx).update_and_run(file_name, macro_name, bas_name)
    #Basic(ctx).update_and_run('Untitled 1', 'Standard.Module1.main',
    #    '/home/asuka/Desktop/python/moduleA.bas')
