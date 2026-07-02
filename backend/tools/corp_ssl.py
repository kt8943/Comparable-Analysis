"""Trust the OS certificate store so corporate TLS-intercepting proxies work.

Corporate proxies (e.g. Zscaler) re-sign HTTPS with a company root CA. That CA
lives in the OS/Windows trust store (so browsers and pip work) but NOT in
Python's bundled certifi store, which causes:
    SSL: CERTIFICATE_VERIFY_FAILED: unable to get local issuer certificate

Importing this module makes Python verify against the OS trust store instead,
fixing those errors for every HTTPS library (requests, urllib, httpx, ...).

No-op if ``truststore`` is not installed (e.g. a dev machine not behind such a
proxy), so it is always safe to import. SSL-only — nothing else.
"""

try:
    import truststore
    truststore.inject_into_ssl()
except Exception:
    # truststore absent or injection failed — keep default SSL behaviour.
    pass
