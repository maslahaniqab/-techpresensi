import os

from flask import Flask, jsonify, request

from providers.fonnte import FontteProvider

app = Flask(__name__)

API_KEY = os.environ.get("NOTIF_API_KEY", "")
PROVIDER = FontteProvider(os.environ.get("FONNTE_TOKEN", ""))

GRUP = {
    "permohonan": os.environ.get("GRUP_PERMOHONAN", ""),
    "produksi": os.environ.get("GRUP_PRODUKSI", ""),
    "aff": os.environ.get("GRUP_AFF", ""),
    "marketing": os.environ.get("GRUP_MARKETING", ""),
    "live": os.environ.get("GRUP_LIVE", ""),
}


@app.post("/kirim")
def kirim():
    if not API_KEY or request.headers.get("X-API-Key") != API_KEY:
        return jsonify({"status": False, "reason": "unauthorized"}), 401

    data = request.get_json(silent=True) or {}
    tujuan = data.get("tujuan", "")
    pesan = (data.get("pesan") or "").strip()
    if not pesan:
        return jsonify({"status": False, "reason": "pesan kosong"}), 400

    target_ids = [GRUP[t] for t in tujuan.split(",") if GRUP.get(t)]
    if not target_ids:
        return jsonify({"status": False, "reason": "tujuan tidak dikenal"}), 400

    ok, info = PROVIDER.kirim(target_ids, pesan)
    return jsonify({"status": ok, "info": info}), (200 if ok else 502)
