"""Functions to work with the desktop object."""

class DocLibLookupError(Exception):
    """Raised if document name lookup fails."""
    pass


def get_current_doc(desktop):
    """Returns the currently open document."""
    return desktop.getCurrentComponent()

def get_document(desktop, doc_name):
    """Returns the document corresponding to `doc_name`."""
    frames = desktop.getFrames()
    for i in range(frames.getCount()):
        controller = frames.getByIndex(i).getController()
        if controller and controller.getTitle() == doc_name:
            return controller.getModel()

    raise DocLibLookupError(doc_name)
