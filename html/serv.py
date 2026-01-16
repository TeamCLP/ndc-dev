
#!/usr/bin/env python3
import http.server
import socketserver
import os
import io
from urllib.parse import urlparse, unquote
import http.client

# --- Config par variables d'env ---
PORT = int(os.environ.get("PORT", "6006"))  # mets 6006 si tu remplaces TensorBoard
HOST = os.environ.get("HOST", "0.0.0.0")
DOCROOT = os.environ.get("DOCROOT", "/home/quentin/scribe/html")
PATH_PREFIX = os.environ.get("PATH_PREFIX", "/scribe-ai/scribe2/url-1").rstrip("/")
API_UPSTREAM_HOST = os.environ.get("API_UPSTREAM_HOST", "127.0.0.1")
API_UPSTREAM_PORT = int(os.environ.get("API_UPSTREAM_PORT", "5000"))
API_PREFIX = os.environ.get("API_PREFIX", f"{PATH_PREFIX}/api").rstrip("/")

class Handler(http.server.SimpleHTTPRequestHandler):
    # ---- Static sous PATH_PREFIX ----
    def translate_path(self, path):
        raw_path = unquote(urlparse(path).path)
        # Proxy branch → ne pas traduire en chemin local
        if raw_path.startswith(API_PREFIX + "/") or raw_path == API_PREFIX:
            # Retourne chaîne vide pour indiquer qu'on gère ailleurs
            return ""
        # Enforce prefix pour le static
        if PATH_PREFIX and not raw_path.startswith(PATH_PREFIX):
            self.send_error(404, "Path en dehors du prefix")
            return ""
        stripped = (raw_path[len(PATH_PREFIX):] or "/")
        return os.path.join(DOCROOT, stripped.lstrip("/"))

    # ---- Proxy branch: /.../api/... → http://127.0.0.1:5000/... ----
    def _proxy(self):
        # Construit le chemin amont sans le préfixe API
        parsed = urlparse(self.path)
        in_path = unquote(parsed.path)
        # Support /prefix/api et /prefix/api/...
        tail = in_path[len(API_PREFIX):]
        if tail.startswith("/"):
            tail = tail[1:]
        upstream_path = f"/{tail}"
        if parsed.query:
            upstream_path += f"?{parsed.query}"

        # Prépare connexion HTTP à l’API locale
        conn = http.client.HTTPConnection(API_UPSTREAM_HOST, API_UPSTREAM_PORT, timeout=300)

        # Filtrage minimal d’en-têtes (on évite hop-by-hop)
        headers = {}
        for k, v in self.headers.items():
            lk = k.lower()
            if lk in ("host", "connection", "keep-alive", "proxy-authenticate", "proxy-authorization",
                      "te", "trailers", "transfer-encoding", "upgrade"):
                continue
            headers[k] = v

        # Corps éventuel (POST/PUT/PATCH…)
        body = None
        if "content-length" in self.headers:
            cl = int(self.headers["content-length"])
            body = self.rfile.read(cl) if cl > 0 else None

        # Envoie la requête amont
        method = self.command
        conn.request(method, upstream_path, body=body, headers=headers)
        upstream_resp = conn.getresponse()

        # Réexpédie le statut + headers + body au client
        self.send_response(upstream_resp.status, upstream_resp.reason)

        # Copie des headers (sauf hop-by-hop)
        ignore = {"connection", "keep-alive", "proxy-authenticate", "proxy-authorization",
                  "te", "trailers", "transfer-encoding", "upgrade"}
        for (hk, hv) in upstream_resp.headers.items():
            if hk.lower() in ignore:
                continue
            # Optionnel: réécrire location si l’API redirige vers absolu http://127.0.0.1
            if hk.lower() == "location" and hv.startswith("http://"):
                # Réécrit vers même origine + API_PREFIX
                hv = f"{API_PREFIX}{urlparse(hv).path}"
            self.send_header(hk, hv)
        self.end_headers()

        # Flux de contenu
        data = upstream_resp.read()
        if data:
            self.wfile.write(data)
        conn.close()

    def do_GET(self):
        if self._maybe_proxy():
            return
        return super().do_GET()

    def do_POST(self):
        if self._maybe_proxy():
            return
        return super().do_POST()

    def do_PUT(self):
        if self._maybe_proxy():
            return
        return super().do_PUT()

    def do_DELETE(self):
        if self._maybe_proxy():
            return
        return super().do_DELETE()

    def do_OPTIONS(self):
        # Si le front envoie des preflight, on peut les proxifier aussi
        if self._maybe_proxy():
            return
        return super().do_OPTIONS()

    def _maybe_proxy(self):
        # Détermine si le path doit être proxifié
        raw_path = unquote(urlparse(self.path).path)
        if raw_path == API_PREFIX or raw_path.startswith(API_PREFIX + "/"):
            try:
                self._proxy()
            except Exception as e:
                self.send_error(502, f"Erreur proxy: {e}")
            return True
        return False

if __name__ == "__main__":
    os.chdir(DOCROOT)
    with socketserver.TCPServer((HOST, PORT), Handler) as httpd:
        print(f"[serve] Static: {DOCROOT} sous {PATH_PREFIX} | Proxy API: {API_UPSTREAM_HOST}:{API_UPSTREAM_PORT} sous {API_PREFIX} | Listen {HOST}:{PORT}")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            pass
