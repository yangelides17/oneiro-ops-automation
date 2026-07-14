"""
Lightweight on-demand PDF fill endpoint, served in a background thread by
the worker process (watch_and_fill.py). The Node webapp POSTs a fill spec
and streams the returned PDF straight to the browser as a download — no
Drive queue, no Approvals.

Endpoint:
  GET  /health                      -> {"ok": true}
  POST /fill   (header X-Fill-Key)   body {doc_kind, fields, filename}
       -> application/pdf bytes (or {error} JSON on failure)

Auth: a shared secret in the X-Fill-Key header (env FILL_SERVER_KEY),
matching the value the Node service sends. Reachable only over Railway's
private network in production.
"""
import json
import socket
import tempfile
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from fill_form import DOC_KIND_TEMPLATE, fill_acroform


class _DualStackServer(ThreadingHTTPServer):
    # Railway's private network is IPv6-only; bind :: dual-stack so the web
    # service can reach us over railway.internal (and IPv4 localhost too).
    address_family = socket.AF_INET6

    def server_bind(self):
        try:
            self.socket.setsockopt(socket.IPPROTO_IPV6, socket.IPV6_V6ONLY, 0)
        except (AttributeError, OSError):
            pass
        super().server_bind()


def _make_handler(service, get_template, key):
    class Handler(BaseHTTPRequestHandler):
        def _send(self, code, body=b'', ctype='application/json', extra=None):
            self.send_response(code)
            self.send_header('Content-Type', ctype)
            self.send_header('Content-Length', str(len(body)))
            for k, v in (extra or {}).items():
                self.send_header(k, v)
            self.end_headers()
            if body:
                self.wfile.write(body)

        def do_GET(self):
            if self.path.rstrip('/') == '/health':
                self._send(200, b'{"ok":true}')
            else:
                self._send(404, b'{"error":"not found"}')

        def do_POST(self):
            if self.path.rstrip('/') != '/fill':
                return self._send(404, b'{"error":"not found"}')
            if key and self.headers.get('X-Fill-Key') != key:
                return self._send(401, b'{"error":"unauthorized"}')
            try:
                n = int(self.headers.get('Content-Length') or 0)
                data = json.loads(self.rfile.read(n) or b'{}')
                doc_kind = str(data.get('doc_kind') or '').upper()
                fields   = data.get('fields') or {}
                filename = str(data.get('filename') or 'month-end.pdf')
                doc_type = DOC_KIND_TEMPLATE.get(doc_kind)
                if not doc_type:
                    return self._send(400, json.dumps({'error': f'unknown doc_kind {doc_kind!r}'}).encode())

                template = get_template(service, doc_type)
                with tempfile.TemporaryDirectory() as td:
                    out = Path(td) / 'filled.pdf'
                    fill_acroform(str(template), fields, str(out))
                    pdf = out.read_bytes()
                safe = filename.replace('"', '').replace('\r', '').replace('\n', '')
                self._send(200, pdf, 'application/pdf',
                           {'Content-Disposition': f'attachment; filename="{safe}"'})
            except Exception as e:
                self._send(500, json.dumps({'error': str(e)}).encode())

        def log_message(self, *a):   # keep the worker log quiet
            pass

    return Handler


def start_fill_server(service, get_template, port: int, key: str):
    """Start the fill HTTP server in a daemon thread. Returns the server."""
    httpd = _DualStackServer(('::', port), _make_handler(service, get_template, key))
    threading.Thread(target=httpd.serve_forever, daemon=True, name='fill-server').start()
    return httpd
