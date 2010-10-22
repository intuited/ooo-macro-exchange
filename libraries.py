"""Functionality related to UNO sequences of libraries."""

class ReadonlyLibraryError(Exception):
    """Raised if an attempt is made to update a readonly library."""
    pass

class PasswordProtectionError(Exception):
    """Raised if a password check prevents a library update."""
    pass


def get_lib_by_name(libs, library_name, mode='read'):
    """Get the library named by `library_name`.

    `libraries` is an UNO sequence of libraries,
    as returned by `document.get_libraries`.

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
