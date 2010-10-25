"""Functions to get data or objects from a context."""

import desktop, document, libraries

def get_desktop(context, service_manager):
    create_instance = service_manager.createInstanceWithContext
    desktop_class = "com.sun.star.frame.Desktop"
    return create_instance(desktop_class, context)


def get_application_libraries(context, service_manager):
    # TODO: find out if it's possible for an exception to be thrown
    #       by create_instance in a situation where
    #       there is legitimately no library for the application.
    create_instance = service_manager.createInstanceWithContext
    libraries_class = "com.sun.star.script.ApplicationScriptLibraryContainer"
    return create_instance(library_class, context)


# Not sure if there's any point in parametrizing these functions.
# The idea is to allow their results to be memoized in some manner,
# perhaps by virtue of being properties of an object.
# I'm mildly concerned that they could return inconsistent results
# if the OOo app changes document while this macro is running.
# Memoizing would allow consistent results to be returned
# over the scope of a "transaction".
def resolve_doc_name(context, service_manager, desktop, doc_name,
                     get_current_doc=desktop.get_current_doc,
                     get_application_libraries=get_application_libraries,
                     get_libraries=document.get_libraries,
                     get_document=desktop.get_document):
    """Returns (doc, libraries).

    If `doc_name` is "application",
    returns the current document and the app library.

    Raises DocLibLookupError if a document with the specified name
    is not found.

    The various get_* functions can be overridden.
    """
    if doc_name == 'application':
        return (get_current_doc(desktop),
                get_application_libraries(context, service_manager))
    doc = get_document(desktop, doc_name)
    libs = libraries.Libraries(get_libraries(doc))
    return doc, libs
