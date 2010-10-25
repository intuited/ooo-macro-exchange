"""Routines to update and access macros in libraries."""
from collections import Mapping

from container import name_access, name_container

class ReadonlyLibraryError(Exception):
    """Raised if an attempt is made to update a readonly library."""
    pass


class WriteableLibrary(name_container):
    pass

class ReadOnlyLibrary(WriteableLibrary):
    def __setitem__(self, name, value):
        raise ReadonlyLibraryError(
            "Tried to set module '{0}' on a read-only library.".format(name))
    def __delitem__(self, name):
        raise ReadonlyLibraryError(
            "Tried to remove module '{0}' of a read-only library.".format(name))
