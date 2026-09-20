"""
Painel Comercial A&B — função serverless (Vercel Python Runtime).

Corre sob pedido (não agendada) sempre que o painel a chama, tipicamente
ao abrir. Vai buscar os 3 ficheiros aos links partilhados (OneDrive ou
Google Drive — deteta automaticamente qual dos dois), trata-os, e
devolve o resultado ao painel.

(Nota: a cópia de reserva na base de dados Supabase foi retirada desta
função por agora, para reduzir dependências e risco de falha. Pode ser
reintroduzida mais tarde.)

Disponível em: /api/refresh

Variáveis de ambiente necessárias (Vercel → Settings → Environment Variables):
    ONEDRIVE_DASHBOARD_URL       — link de partilha (OneDrive ou Google Drive)
    ONEDRIVE_WORKORDERS_URL      — idem
    ONEDRIVE_SALES_REPORT_URL    — idem
"""
import os, re, json
from io import BytesIO
from http.server import BaseHTTPRequestHandler

from parser_core import parse_all  # co-located in api/ so Vercel bundles it reliably

import requests


def _gdrive_file_id(url):
    m = re.search(r'/d/([a-zA-Z0-9_-]+)', url) or re.search(r'[?&]id=([a-zA-Z0-9_-]+)', url)
    return m.group(1) if m else None


def _to_onedrive_direct(url):
    sep = '&' if '?' in url else '?'
    return url if 'download=1' in url else f'{url}{sep}download=1'


def fetch_bytes(url):
    # Google Sheets nativo (o Drive converteu o .xlsx ao fazer upload) —
    # pede a exportação em .xlsx, não o "ficheiro" em si.
    if 'docs.google.com/spreadsheets' in url:
        gid = _gdrive_file_id(url)
        if not gid:
            raise RuntimeError(f"Não consegui identificar o ID da folha de cálculo em {url}")
        resp = requests.get(f'https://docs.google.com/spreadsheets/d/{gid}/export?format=xlsx', timeout=25)
        ctype = resp.headers.get('Content-Type', '')
        if resp.status_code != 200 or 'text/html' in ctype:
            raise RuntimeError(f"Não consegui exportar a folha de cálculo {url} "
                                f"(status {resp.status_code}, tipo {ctype}). "
                                f"Confirma que a partilha está como 'Qualquer pessoa com o link'.")
        return resp.content

    gid = _gdrive_file_id(url) if 'drive.google.com' in url else None

    if gid:
        # Ficheiro .xlsx "cru" no Drive: needs uc?export=download&id=..., e
        # ficheiros maiores/sinalizados mostram uma página HTML de aviso
        # ("não foi possível analisar quanto a vírus") com um token de
        # confirmação que é preciso reenviar num segundo pedido.
        session = requests.Session()
        resp = session.get(f'https://drive.google.com/uc?export=download&id={gid}', timeout=25)
        ctype = resp.headers.get('Content-Type', '')
        if 'text/html' in ctype:
            token = None
            m = re.search(r'confirm=([0-9A-Za-z_-]+)', resp.text)
            if m:
                token = m.group(1)
            else:
                for k, v in resp.cookies.items():
                    if k.startswith('download_warning'):
                        token = v
            if token:
                resp = session.get(
                    f'https://drive.google.com/uc?export=download&id={gid}&confirm={token}',
                    timeout=25)
                ctype = resp.headers.get('Content-Type', '')
        if resp.status_code != 200 or 'text/html' in ctype:
            raise RuntimeError(f"Não consegui descarregar de {url} sem autenticação "
                                f"(status {resp.status_code}, tipo {ctype}). "
                                f"Confirma que a partilha está como 'Qualquer pessoa com o link'.")
        return resp.content

    # OneDrive / SharePoint
    resp = requests.get(_to_onedrive_direct(url), allow_redirects=True, timeout=25)
    ctype = resp.headers.get('Content-Type', '')
    if resp.status_code != 200 or 'text/html' in ctype:
        raise RuntimeError(f"Não consegui descarregar de {url} sem autenticação "
                            f"(status {resp.status_code}, tipo {ctype}).")
    return resp.content


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        try:
            dash_url = os.environ.get('ONEDRIVE_DASHBOARD_URL')
            wo_url = os.environ.get('ONEDRIVE_WORKORDERS_URL')
            sr_url = os.environ.get('ONEDRIVE_SALES_REPORT_URL')
            missing = [n for n, v in [('ONEDRIVE_DASHBOARD_URL', dash_url),
                                       ('ONEDRIVE_WORKORDERS_URL', wo_url),
                                       ('ONEDRIVE_SALES_REPORT_URL', sr_url)] if not v]
            if missing:
                raise RuntimeError(f"Variáveis de ambiente em falta na Vercel: {', '.join(missing)}")

            dash = BytesIO(fetch_bytes(dash_url))
            wo = BytesIO(fetch_bytes(wo_url))
            sr = BytesIO(fetch_bytes(sr_url))
            result = parse_all(dash, wo, sr)

            self._send(200, result)
        except Exception as e:
            self._send(502, {'error': f'{type(e).__name__}: {e}'})

    def _send(self, code, obj):
        body = json.dumps(obj, default=str).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)
