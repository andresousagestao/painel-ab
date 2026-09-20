"""
Painel Comercial A&B — lógica de leitura/tratamento dos 3 ficheiros Excel.
Partilhada entre a função serverless (api/refresh.py) e qualquer uso manual.
Não faz pedidos de rede nem toca na base de dados — só recebe os ficheiros
(caminho ou objeto tipo-ficheiro) e devolve os dados já tratados.
"""
import re
from datetime import datetime
import warnings
import pandas as pd

warnings.filterwarnings('ignore')

EMP_RE = re.compile(r'^st[mo]\d{3}a\d+$', re.I)
ID_STORE_RE = re.compile(r'^st([mo])(\d{3})a\d+$', re.I)
PERIODS = ['LD', 'WTD', 'MTD', 'QTD', 'YTD']

ADDITIVE = {'net_sales', 'sales_opt', 'sales_sun', 'sales_cl', 'sales_oth', 'cp_volume',
            'opt_volume', 'vaas', 'vaas_volume', 'insurance_sales', 'vs_tgt_eur'}
VOL_KEY = {'com': 'cp_volume', 'acc': 'opt_volume'}
ASP_KEY = {'com': 'cp_asp', 'acc': 'opt_asp'}

COM_KEYS = {
    'net_sales': 'NET SALES | Commercial Sales', 'vs_py_pct': 'NET SALES | vs PY %',
    'sales_opt': 'CATEGORY SALES | Net Sales OPT', 'sales_sun': 'CATEGORY SALES | Net Sales SUN',
    'sales_cl': 'CATEGORY SALES | Net Sales CL', 'sales_oth': 'CATEGORY SALES | Net Sales OTH',
    'cp_volume': 'COMPLETE PAIR | Volume', 'cp_asp': 'COMPLETE PAIR | ASP',
    'mf_pct': 'OPTICAL KPI | MF %', 'photochromic_pct': 'OPTICAL KPI | Photochromic %',
    'blue_pct': 'OPTICAL KPI | Blue %', 'vaas': 'OPTICAL KPI | VaaS €',
    'vaas_volume': 'OPTICAL KPI | VaaS Volume', 'vaas_share_pct': 'OPTICAL KPI | VaaS Share %',
}
ACC_KEYS = {
    'net_sales': 'NET SALES | Accounting Sales', 'vs_py_pct': 'NET SALES | vs PY %',
    'vs_tgt_eur': 'NET SALES | vs TGT €', 'vs_tgt_pct': 'NET SALES | vs TGT %',
    'sales_opt': 'CATEGORY NET SALES | Sales OPT', 'sales_sun': 'CATEGORY NET SALES | Sales SUN',
    'sales_cl': 'CATEGORY NET SALES | Sales CL', 'sales_oth': 'CATEGORY NET SALES | Sales OTH',
    'opt_volume': 'TOTAL OPTICAL | Volume', 'opt_asp': 'TOTAL OPTICAL | ASP',
    'mf_pct': 'OPTICAL KPI | MF %', 'photochromic_pct': 'OPTICAL KPI | Photochromic %',
    'blue_pct': 'OPTICAL KPI | Blue %', 'vaas': 'OPTICAL KPI | VaaS Sales',
    'vaas_volume': 'OPTICAL KPI | VaaS Volume', 'vaas_share_pct': 'OPTICAL KPI | VaaS Share',
    'insurance_sales': 'OTHER | Insurance Sales',
}


def g(rec, key):
    v = rec.get(key)
    return None if v in ('X', '#DIV/0', None) else v


def parse_kpi_sheet(src, sheet_name):
    df = pd.read_excel(src, sheet_name=sheet_name, header=None)
    r1 = df.iloc[1, 1]
    r2, r3 = df.iloc[2], df.iloc[3]
    colnames, last_group = {}, None
    for i in range(6, df.shape[1]):
        sub = r2[i] if isinstance(r2[i], str) else None
        metric = r3[i] if isinstance(r3[i], str) else None
        if sub:
            last_group = sub
        if metric:
            colnames[i] = f"{last_group} | {metric}" if last_group else metric
    employees, stores, current_store = [], [], None
    for idx in range(4, df.shape[0]):
        row = df.iloc[idx]
        country, lux_id = row[1], row[5]
        if isinstance(country, str) and country == 'PT':
            current_store = {'global_id': row[3], 'store_name': row[4]}
            continue
        if isinstance(row[1], str) and ' - ' in row[1] and pd.isna(lux_id):
            gid, name = row[1].split(' - ', 1)
            rec = {'global_id': gid.strip(), 'store_name': name.strip()}
            for i, cn in colnames.items():
                v = row[i]
                rec[cn] = None if pd.isna(v) else (float(v) if isinstance(v, (int, float)) else v)
            stores.append(rec)
            continue
        if pd.notna(lux_id) and current_store:
            lux_str = str(lux_id).strip()
            if EMP_RE.match(lux_str):
                rec = {'employee_id': lux_str.lower(), **current_store}
                for i, cn in colnames.items():
                    v = row[i]
                    rec[cn] = None if pd.isna(v) else (float(v) if isinstance(v, (int, float)) else v)
                employees.append(rec)
    return {'period_label': r1, 'employees': employees, 'stores': stores}


def home_store_from_id(emp_id):
    m = ID_STORE_RE.match(emp_id)
    return f'5-{m.group(1).upper()}{m.group(2)}' if m else None


def combine_period_rows(rows, keys, kind):
    mapped = [{k: g(r, v) for k, v in keys.items()} | {'global_id': r['global_id'], 'store_name': r['store_name']}
              for r in rows]
    by_store = {m['global_id']: {k: m[k] for k in keys} | {'store_name': m['store_name']} for m in mapped}
    if len(mapped) == 1:
        total = dict(mapped[0])
        del total['global_id']; del total['store_name']
        return total, by_store
    total = {}
    for k in keys:
        vals = [m[k] for m in mapped if m.get(k) is not None]
        total[k] = round(sum(vals), 2) if (k in ADDITIVE and vals) else None
    dominant = max(mapped, key=lambda m: abs(m.get('net_sales') or 0))
    for k in keys:
        if total.get(k) is None and k not in ADDITIVE:
            total[k] = dominant.get(k)
    vk, ak = VOL_KEY[kind], ASP_KEY[kind]
    if vk in keys and ak in keys and total.get(vk):
        tot_sales = total.get('sales_opt')
        total[ak] = round(tot_sales / total[vk], 2) if tot_sales is not None else dominant.get(ak)
    return total, by_store


def parse_all(dash_src, wo_src, sr_src):
    """dash_src/wo_src/sr_src: file path OR file-like object (e.g. io.BytesIO)
    for the dashboard, work-orders and sales-report workbooks respectively.
    Returns {'meta': {...}, 'stores': {...}, 'employees': {...}}."""
    kpi_data = {p: {k: parse_kpi_sheet(dash_src, f'{p}-{k}') for k in ['COM', 'ACC']} for p in PERIODS}

    snapshot_date = datetime.today().strftime('%d/%m/%Y')
    m = re.search(r'Date:\s*(\d{2}/\d{2}/\d{2})', str(kpi_data['LD']['COM']['period_label']))
    if m:
        d, mo, y = m.group(1).split('/')
        snapshot_date = f'{d}/{mo}/20{y}'

    sr = pd.read_excel(sr_src, sheet_name='LD Sales', header=6).drop(columns=['Unnamed: 0'])
    sr['Employee ID'] = sr['Employee ID'].astype(str).str.lower()
    sr = sr[sr['Employee ID'].str.match(EMP_RE)]
    txn_summary = {}
    for emp, grp in sr.groupby('Employee ID'):
        sales_rows = grp[grp['Sales Detail'] == 'SALES']
        top_products = (sales_rows[sales_rows['Net Sales'] > 0]
                         .groupby(['SAP Article Desc'])['Net Sales'].sum()
                         .sort_values(ascending=False).head(8).round(2).to_dict())
        txn_summary[emp] = {
            'net_sales_today': round(float(grp['Net Sales'].sum()), 2),
            'transactions': int(grp['Receipt Number'].nunique()),
            'top_products': top_products,
        }

    wo = pd.read_excel(wo_src, sheet_name='RECAP BY STORE AND CLUSTER', header=None)
    aging, i = {}, 0
    while i < len(wo):
        row = wo.iloc[i]
        if row[1] == 'Global Store ID':
            cluster_label, j = None, i + 1
            while j < len(wo) and pd.notna(wo.iloc[j][1]) and wo.iloc[j][1] != 'Global Store ID':
                r = wo.iloc[j]
                gid, cluster, amt = r[1], r[3], r[4]
                if pd.notna(cluster):
                    cluster_label = cluster
                if pd.notna(gid) and pd.notna(amt):
                    aging.setdefault(gid, {})[cluster_label] = float(amt)
                j += 1
            i = j
        else:
            i += 1

    store_names = {}
    for p in PERIODS:
        for s in kpi_data[p]['ACC']['stores']:
            store_names[s['global_id']] = s['store_name']

    stores_out = {}
    for gid, name in store_names.items():
        periods_out = {}
        for p in PERIODS:
            com_list = [s for s in kpi_data[p]['COM']['stores'] if s['global_id'] == gid]
            acc_list = [s for s in kpi_data[p]['ACC']['stores'] if s['global_id'] == gid]
            pdata = {}
            if com_list:
                pdata['com'] = {k: g(com_list[0], v) for k, v in COM_KEYS.items()}
            if acc_list:
                pdata['acc'] = {k: g(acc_list[0], v) for k, v in ACC_KEYS.items()}
            periods_out[p] = pdata
        stores_out[gid] = {'name': name, 'aging': aging.get(gid, {}), 'periods': periods_out}

    emp_ids = {e['employee_id'] for p in PERIODS for e in kpi_data[p]['COM']['employees']}
    employees_out = {}
    for emp in sorted(emp_ids):
        home_gid = home_store_from_id(emp)
        periods_out = {}
        for p in PERIODS:
            com_rows = [e for e in kpi_data[p]['COM']['employees'] if e['employee_id'] == emp]
            acc_rows = [e for e in kpi_data[p]['ACC']['employees'] if e['employee_id'] == emp]
            pdata, com_by_store, acc_by_store = {}, {}, {}
            if com_rows:
                pdata['com'], com_by_store = combine_period_rows(com_rows, COM_KEYS, 'com')
            if acc_rows:
                pdata['acc'], acc_by_store = combine_period_rows(acc_rows, ACC_KEYS, 'acc')
            by_store = {}
            for gid in set(com_by_store) | set(acc_by_store):
                c, a = com_by_store.get(gid, {}), acc_by_store.get(gid, {})
                if not (c.get('net_sales') or a.get('net_sales')):
                    continue
                by_store[gid] = {
                    'store_name': c.get('store_name') or a.get('store_name') or store_names.get(gid, gid),
                    'com': {k: c[k] for k in COM_KEYS} if c else None,
                    'acc': {k: a[k] for k in ACC_KEYS} if a else None,
                }
            pdata['by_store'] = by_store
            periods_out[p] = pdata
        employees_out[emp] = {
            'store_gid': home_gid,
            'periods': periods_out,
            'today': txn_summary.get(emp),
        }

    return {
        'meta': {'snapshot_date': snapshot_date},
        'stores': stores_out,
        'employees': employees_out,
    }
