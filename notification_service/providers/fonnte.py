import json
import urllib.parse
import urllib.request


class FontteProvider:
    API_URL = "https://api.fonnte.com/send"

    def __init__(self, token):
        self.token = token

    def kirim(self, target_ids, pesan):
        data = urllib.parse.urlencode({
            "target": ",".join(target_ids),
            "message": pesan,
            "countryCode": "0",
        }).encode()
        req = urllib.request.Request(
            self.API_URL, data=data, method="POST", headers={"Authorization": self.token},
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            hasil = json.loads(resp.read().decode("utf-8"))
        return bool(hasil.get("status")), hasil.get("reason") or hasil.get("detail")
