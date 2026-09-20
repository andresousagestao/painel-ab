"""
Painel Comercial A&B — função serverless (Vercel Python Runtime).

Corre sob pedido (não agendada) sempre que o painel a chama, tipicamente
ao abrir. Vai buscar os 3 ficheiros aos links do OneDrive, trata-os, e
devolve o resultado ao painel. Também guarda uma cópia na base de dados
Supabase, para servir de reserva caso uma chamada futura falhe (rede
lenta, ficheiro temporariamente indisponível, etc.).

Disponível em: /api/refresh

Variáveis de ambiente necessárias (Vercel → Settings → Environment Variables):
    ONEDRIVE_DASHBOARD_URL
    ONEDRIVE_WORKORDERS_URL
    ONEDRIVE_SALES_REPORT_URL
    SUPABASE_URL
    SUPABASE_SERVICE_ROLE_KEY
"""
import os, sys, json
from io import BytesIO
from http.server import BaseHTTPRequestHandler

from parser_core import parse_all  # co-located in api/ so Vercel bundles it reliably

import requests
from supabase import create_client


def to_direct_download(url):
    sep = '&' if '?' in url else '?'
    return url if 'download=1' in url else f'{url}{sep}download=1'


def fetch_bytes(url):
    resp = requests.get(to_direct_download(url), allow_redirects=True, timeout=25)
    ctype = resp.headers.get('Content-Type', '')
    if resp.status_code != 200 or 'text/html' in ctype:
        raise RuntimeError(f"Não consegui descarregar de {url} sem autenticação "
                            f"(status {resp.status_code}, tipo {ctype}).")
    return resp.content


def cache_to_supabase(result):
    sb = create_client(os.environ['SUPABASE_URL'], os.environ['SUPABASE_SERVICE_ROLE_KEY'])
    sb.table('meta').upsert({'id': 1, 'snapshot_date': result['meta']['snapshot_date']}).execute()
    store_rows = [{'gid': gid, 'name': s['name'], 'aging': s['aging'], 'periods': s['periods']}
                  for gid, s in result['stores'].items()]
    sb.table('stores').upsert(store_rows).execute()
    emp_rows = [{'employee_id': eid, 'store_gid': e['store_gid'], 'periods': e['periods'], 'today': e['today']}
                for eid, e in result['employees'].items()]
    for i in range(0, len(emp_rows), 20):
        sb.table('employees').upsert(emp_rows[i:i + 20]).execute()


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        try:
            dash = BytesIO(fetch_bytes(os.environ['ONEDRIVE_DASHBOARD_URL']))
            wo = BytesIO(fetch_bytes(os.environ['ONEDRIVE_WORKORDERS_URL']))
            sr = BytesIO(fetch_bytes(os.environ['ONEDRIVE_SALES_REPORT_URL']))
            result = parse_all(dash, wo, sr)

            try:
                cache_to_supabase(result)
            except Exception as cache_err:
                # Não falha o pedido só porque a cópia de reserva falhou —
                # o painel já tem os dados frescos para mostrar.
                result['cache_warning'] = str(cache_err)

            self._send(200, result)
        except Exception as e:
            self._send(502, {'error': str(e)})

    def _send(self, code, obj):
        body = json.dumps(obj, default=str).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)
