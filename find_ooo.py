"""Serves to find the OOo installation's libraries."""

def find_ooo():
    """Returns the path to OpenOffice's ``program`` directory."""
    # Just return the standard debian-ish location for 3.2 for now.
    return "/usr/lib/openoffice/basis3.2/program/"
