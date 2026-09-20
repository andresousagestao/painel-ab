from http.server import BaseHTTPRequestHandler
import openpyxl
import requests

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        msg = f"OK - openpyxl {openpyxl.__version__}, requests {requests.__version__}"
        self.send_response(200)
        self.send_header('Content-Type', 'text/plain; charset=utf-8')
        self.end_headers()
        self.wfile.write(msg.encode('utf-8'))
