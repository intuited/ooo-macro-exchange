"""Serves to find the OOo installation's libraries.

This is mostly provided as an abstraction
to make dependency injection more convenient.
"""

def find_ooo():
    """Returns the path to OpenOffice's ``program`` directory.

    May be useful for systems where the python uno bridge
    is not on the system PYTHONPATH.
    """
    import os
    try:
        return os.environ['PY_UNO_PATH']
    except KeyError:
        # Just return the standard debian-ish location for 3.2.
        # This is unlikely to ever be useful,
        # since the debian ``python-uno`` module puts `uno` on the PYTHONPATH,
        # but it serves as an example.
        return "/usr/lib/openoffice/basis3.2/program/"

def find_uno():
    try:
        import uno
    except ImportError:
        import sys
        sys.path.insert(0, find_ooo())
        import uno
    return uno
