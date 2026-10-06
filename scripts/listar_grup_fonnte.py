import json
import os
import urllib.request

TOKEN = os.environ.get("FONNTE_TOKEN", "").strip()
if not TOKEN:
    raise SystemExit("Set dulu environment FONNTE_TOKEN dengan token dari dashboard Fonnte.")


def panggil(url):
    req = urllib.request.Request(url, method="POST", headers={"Authorization": TOKEN})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode("utf-8"))


print("Memperbarui daftar grup...")
print(panggil("https://api.fonnte.com/fetch-group"))

hasil = panggil("https://api.fonnte.com/get-whatsapp-group")
print("\nDaftar grup:")
print(json.dumps(hasil, ensure_ascii=False, indent=2))
