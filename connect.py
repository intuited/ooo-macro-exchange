"""Code to establish connection to an OOo instance."""
import find_ooo

# Equivalent to compose(openoffice.connect, openoffice._build_connect_string)
# when using the python-openoffice module.. except for `find_uno`.
def get_context(host='localhost', port=2002, find_uno=find_ooo.find_uno):
    """Returns a resolved connection context.

    `find_uno` is called to acquire an `uno`:module: object.
    All subsequent calls to the UNO API will go through this object.

    Returns a `component context`, accessed via the `current context`_.

    .. _component context:
       http://wiki.services.openoffice.org/wiki
             /Documentation/DevGuide/ProUNO/Component_Context
    .. _current context:
       http://udk.openoffice.org
             /common/man/concept/uno_contexts.html#current_context
    """

    uno = find_uno()
    localctx = uno.getComponentContext()
    create_instance = localctx.getServiceManager().createInstanceWithContext

    resolver_class = "com.sun.star.bridge.UnoUrlResolver"
    resolver = create_instance(resolver_class, localctx)

    # http://udk.openoffice.org/common/man/spec/uno-url.html
    context_url = ("uno:socket,host={host},port={port};"
                   "urp;StarOffice.ComponentContext".format)

    return resolver.resolve(context_url(host=host, port=port))
