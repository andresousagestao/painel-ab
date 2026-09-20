"""
Painel Comercial A&B — utilitário manual (não faz parte da automação).

Útil para testar a leitura dos ficheiros localmente, ou para popular a
base de dados uma primeira vez antes de a função serverless (api/refresh.py)
começar a correr sozinha.

Uso:
    python process_files.py <dashboard.xlsx> <workorders.xlsx> <sales_report.xlsx>

Variáveis de ambiente necessárias:
    SUPABASE_URL
    SUPABASE_SERVICE_ROLE_KEY
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'api'))
from parser_core import parse_all
from supabase import create_client


def main():
    if len(sys.argv) != 4:
        sys.exit("Uso: python process_files.py <dashboard.xlsx> <workorders.xlsx> <sales_report.xlsx>")

    result = parse_all(sys.argv[1], sys.argv[2], sys.argv[3])

    sb = create_client(os.environ['SUPABASE_URL'], os.environ['SUPABASE_SERVICE_ROLE_KEY'])
    sb.table('meta').upsert({'id': 1, 'snapshot_date': result['meta']['snapshot_date']}).execute()
    store_rows = [{'gid': gid, 'name': s['name'], 'aging': s['aging'], 'periods': s['periods']}
                  for gid, s in result['stores'].items()]
    sb.table('stores').upsert(store_rows).execute()
    emp_rows = [{'employee_id': eid, 'store_gid': e['store_gid'], 'periods': e['periods'], 'today': e['today']}
                for eid, e in result['employees'].items()]
    for i in range(0, len(emp_rows), 20):
        sb.table('employees').upsert(emp_rows[i:i+20]).execute()

    print(f"OK — snapshot {result['meta']['snapshot_date']}: "
          f"{len(store_rows)} lojas, {len(emp_rows)} colaboradores.")


if __name__ == '__main__':
    main()
