"""Routines to update and access macros in libraries."""

class IllegalMacroNameError(Exception):
    """Raised if a macro name with less or more than three parts is given."""
    pass


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
