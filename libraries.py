"""Functionality related to UNO sequences of libraries."""
import library, container

class PasswordProtectionError(Exception):
    """Raised if a password check prevents a library update."""
    pass


class Libraries(container.name_access):
    """Represent a libraries container.

    `libraries` is an UNO collection of libraries,
    as returned by `document.get_libraries`.
    """
    # TODO: add __setitem__, __delitem__, etc.
    def __init__(self, libraries):
        self._proxied = libraries

    def __getitem__(self, library_name):
        """Get the library named by `library_name`.

        If the library is readonly, a library.ReadOnlyLibrary is returned.
        Otherwise, the return is a library.WriteableLibrary.

        PasswordProtectionError will be raised
        if an attempt is made to access a password-protected library
        and the password has not been verified.
        """
        libs = self._proxied

        if not libs.hasByName(library_name):
            raise KeyError(library_name)

        if (libs.isLibraryPasswordProtected(library_name)
            and not libs.isLibraryPasswordVerified(library_name)):
            raise PasswordProtectionError(library_name)

        if not libs.isLibraryLoaded(library_name):
            libs.loadLibrary(library_name)

        lib = libs.getByName(library_name)

        if libs.isLibraryReadOnly(library_name):
            return library.ReadOnlyLibrary(lib)
        else:
            return library.WriteableLibrary(lib)
