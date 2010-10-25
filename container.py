from collections import Sequence
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
    def __getitem__(self, item):
        return self._proxied.getByIndex(item)
    def __iter__(self):
        for i in range(len(self)):
            yield self[i]
    def __contains__(self, value):
        return value in iter(self)
    def __getattr__(self, name):
        return getattr(self._proxied, name)
