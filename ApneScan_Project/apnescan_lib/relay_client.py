"""Phone-Relay client — "Scan from phone, anywhere" ka PC-side HTTP client.

Pure Python (urllib) — koi Qt/UI nahi, isliye alag se test ho sakta hai.
Server: phone_relay.php (status.apnesoft.com). Suraksha model:
  - TOKEN (QR me) sirf phone ko UPLOAD karne deta hai,
  - KEY (sirf is PC ke paas) se hi files WAPAS milti hain (status/take/stop).
Har file 'take' hote hi server se DELETE ho jaati hai; session apne aap
expire hota hai — server par kuch nahi bachta.
"""

import json
import os
import re
import secrets
import ssl
import string
import tempfile
import urllib.error
import urllib.parse
import urllib.request

TIMEOUT = 12          # seconds — mobile-data wale uploads ke liye aaraam se
UA = {"User-Agent": "ApneScan-Relay"}


def _open(req, timeout=TIMEOUT):
    """HTTPS open jo PyInstaller/frozen Windows .exe me bhi chale. Frozen exe
    ko kai baar CA-certificate store nahi milta -> SSL verify fail -> pehle
    'internet-mode' chup-chaap fail hoke same-WiFi par gir jaata tha. SSLError
    par bina-verify dobara try karte hain (yahi app ke stats/update-check me
    bhi hota hai) taaki phone-relay internet se hamesha chale."""
    try:
        return urllib.request.urlopen(req, timeout=timeout)
    except urllib.error.URLError as ex:
        if isinstance(getattr(ex, "reason", None), ssl.SSLError):
            return urllib.request.urlopen(
                req, timeout=timeout, context=ssl._create_unverified_context())
        raise


def new_credentials():
    """(token, key): token QR me jaata hai (A-Z0-9, 16), key sirf PC ke paas."""
    alpha = string.ascii_uppercase + string.digits
    token = "".join(secrets.choice(alpha) for _ in range(16))
    key = secrets.token_hex(24)
    return token, key


def _get(url):
    req = urllib.request.Request(url, headers=UA)
    with _open(req) as r:
        return r.read()


def _post(url, fields):
    data = urllib.parse.urlencode(fields).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=UA)
    with _open(req) as r:
        return r.read()


def create_session(base, ttl=1800):
    """Server par nayi session banao. Returns dict:
    {token, key, url (phone ke liye), ttl}. Fail par exception."""
    token, key = new_credentials()
    raw = _post(base + "?api=create", {"t": token, "k": key, "ttl": int(ttl)})
    j = json.loads(raw.decode("utf-8", "replace"))
    if not j.get("ok"):
        raise RuntimeError(j.get("err") or "create failed")
    return {"token": token, "key": key, "ttl": int(j.get("ttl") or ttl),
            "url": base + "?u=" + token}


def status(base, token, key):
    """{ok, files:[{id,name,size}], taken, left(seconds)} — PC hi bula sakta."""
    raw = _get(base + "?api=status&t=" + urllib.parse.quote(token)
               + "&k=" + urllib.parse.quote(key))
    return json.loads(raw.decode("utf-8", "replace"))


def take(base, token, key, file_id, name, dest_dir):
    """Ek file utar kar dest_dir me save karo (server se turant delete ho
    jaati hai). Returns local path."""
    url = (base + "?api=take&t=" + urllib.parse.quote(token)
           + "&k=" + urllib.parse.quote(key) + "&id=" + urllib.parse.quote(file_id))
    req = urllib.request.Request(url, headers=UA)
    ext = os.path.splitext(name or "")[1].lower()
    if not re.match(r"^\.(jpe?g|png|webp|heic|heif|pdf|tiff?)$", ext):
        ext = ".bin"
    fd, path = tempfile.mkstemp(prefix="phone_", suffix=ext, dir=dest_dir)
    try:
        with _open(req, timeout=60) as r, os.fdopen(fd, "wb") as fh:
            while True:
                chunk = r.read(65536)
                if not chunk:
                    break
                fh.write(chunk)
        if os.path.getsize(path) <= 0:
            raise RuntimeError("empty download")
        return path
    except Exception:
        try:
            os.remove(path)
        except Exception:
            pass
        raise


def stop(base, token, key):
    """Session band + server par bachi files delete. Kabhi exception nahi."""
    try:
        _post(base + "?api=stop", {"t": token, "k": key})
        return True
    except Exception:
        return False
