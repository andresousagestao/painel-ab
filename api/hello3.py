from http.server import BaseHTTPRequestHandler

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        try:
            from parser_core import parse_all
            msg = f"OK - parser_core importado, parse_all = {parse_all}"
            code = 200
        except Exception as e:
            msg = f"ERRO ao importar parser_core: {type(e).__name__}: {e}"
            code = 500
        self.send_response(code)
        self.send_header('Content-Type', 'text/plain; charset=utf-8')
        self.end_headers()
        self.wfile.write(msg.encode('utf-8'))
