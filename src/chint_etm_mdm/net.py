"""Opening web addresses from the installed program."""
from __future__ import annotations

import ssl
import urllib.error
import urllib.request


def open_url(request: urllib.request.Request, timeout: float):
    """urlopen that falls back to the program's own certificate bundle.

    A portable build may not find the system certificates, or find a chain the
    system does not trust; then the bundled certifi list is tried once.
    """
    try:
        return urllib.request.urlopen(request, timeout=timeout)
    except urllib.error.URLError as exc:
        if not isinstance(exc.reason, ssl.SSLError):
            raise
        import certifi

        context = ssl.create_default_context(cafile=certifi.where())
        return urllib.request.urlopen(request, timeout=timeout, context=context)
