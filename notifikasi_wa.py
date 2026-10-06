import hashlib
import hmac
import json
import os
import urllib.parse
import urllib.request


def link_pdf_permohonan(permohonan_id, secret_key):
    base = os.environ.get("PUBLIC_BASE_URL", "https://maslahaportal.online").rstrip("/")
    sig = hmac.new(secret_key.encode(), f"permohonan:{permohonan_id}".encode(), hashlib.sha256).hexdigest()[:24]
    return f"{base}/publik/permohonan/{permohonan_id}/{sig}.pdf"


def verifikasi_sig_permohonan(permohonan_id, sig, secret_key):
    harapan = hmac.new(secret_key.encode(), f"permohonan:{permohonan_id}".encode(), hashlib.sha256).hexdigest()[:24]
    return hmac.compare_digest(harapan, sig)


def kirim_wa_grup(pesan):
    token = os.environ.get("FONNTE_TOKEN", "").strip()
    target = os.environ.get("FONNTE_TARGET_GROUP", "").strip()
    if not token or not target:
        return False, "Token atau grup tujuan WA belum diatur di server."

    data = urllib.parse.urlencode({"target": target, "message": pesan, "countryCode": "0"}).encode()
    req = urllib.request.Request(
        "https://api.fonnte.com/send", data=data, method="POST", headers={"Authorization": token},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            hasil = json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        return False, str(e)
    return bool(hasil.get("status")), hasil.get("reason") or hasil.get("detail")
