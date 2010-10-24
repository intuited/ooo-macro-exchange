"""Access functions for documents."""

def script_name_url(script_name):
    """Returns a URL for a script name like '
    """
    fmtstr = 'vnd.sun.star.script:{0}?language=Basic&location=document'
    return fmtstr.format(script_name)


def get_libraries(document):
    """Returns the libraries for the `document`."""
    return document.BasicLibraries

def invoke_macro(doc, script_name):
    provider = doc.getScriptProvider()
    script = provider.getScript(script_name_url(script_name))
    return script.invoke((), (), ())

def save(doc):
    """Save the document to its location."""
    doc.store()
