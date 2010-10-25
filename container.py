from collections import Sequence, Mapping
class index_access(Sequence):
    """Useful for wrapping uno sequences which implement `XIndexAccess`_.

    For example, `desktop.getFrames()` returns such an object.

    .. _XIndexAccess:
       http://api.openoffice.org
             /docs/common/ref/com/sun/star/container/XIndexAccess.html
    """
    def __init__(self, proxied):
        self._proxied = proxied
    def __len__(self):
        return self._proxied.getCount()
    def __iter__(self):
        for i in range(len(self)):
            yield self[i]
    def __contains__(self, value):
        return value in iter(self)
    def __getitem__(self, item):
        if item >= len(self):
            raise IndexError("{0} outside of index range.".format(name))
        return self._proxied.getByIndex(item)
    def __getattr__(self, name):
        return getattr(self._proxied, name)

class name_access(Mapping):
    """Proxies `XNameAccess`_ implementors with Pythonic dict-ness.

    .. _XNameAccess:
       http://api.openoffice.org
             /docs/common/ref/com/sun/star/container/XNameAccess.html
    """
    def __init__(self, proxied):
        self._proxied = proxied
    def __len__(self):
        return len(self._proxied.getElementNames())
    def __iter__(self):
        return iter(self._proxied.getElementNames())
    def __contains__(self, name):
        return self._proxied.hasByName(name)
    def __getitem__(self, name):
        if self._proxied.hasByName(name):
            return self._proxied.getByName(name)
        else:
            raise KeyError("'{0}' not in XNameAccess object.".format(name))
    def __getattr__(self, name):
        return getattr(self._proxied, name)
