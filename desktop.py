"""Functions to work with the desktop object."""
from container import index_access

class DocLibLookupError(Exception):
    """Raised if document name lookup fails."""
    pass


def get_current_doc(desktop):
    """Returns the currently open document."""
    return desktop.getCurrentComponent()

def get_document(desktop, doc_name):
    """Returns the first document corresponding to `doc_name`."""
    # TODO: Consider implementing a check for duplicate document names.
    frames = index_access(desktop.getFrames())
    controllers = (frame.getController() for frame in frames)
    for controller in controllers:
        if controller and controller.getTitle() == doc_name:
            return controller.getModel()

    raise DocLibLookupError(doc_name)
