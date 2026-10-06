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


def kirim_wa_grup(pesan, tujuan="produksi"):
    url = os.environ.get("NOTIFICATION_SERVICE_URL", "").rstrip("/")
    api_key = os.environ.get("NOTIF_API_KEY", "")
    if not url or not api_key:
        return False, "Notification Service belum diatur di server."

    body = json.dumps({"tujuan": tujuan, "pesan": pesan}).encode()
    req = urllib.request.Request(
        f"{url}/kirim", data=body, method="POST",
        headers={"Content-Type": "application/json", "X-API-Key": api_key},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            hasil = json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        return False, str(e)
    return bool(hasil.get("status")), hasil.get("info")
