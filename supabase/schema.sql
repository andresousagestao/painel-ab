-- Painel Comercial A&B — esquema da base de dados Supabase
-- Corre isto em: Supabase Dashboard → SQL Editor → New query → cola tudo → Run

-- ---------- Tabelas de dados (cópia de reserva escrita por api/refresh.py) ----------

create table if not exists meta (
  id int primary key default 1,
  snapshot_date text,
  updated_at timestamptz default now()
);

create table if not exists stores (
  gid text primary key,
  name text not null,
  aging jsonb default '{}'::jsonb,
  periods jsonb default '{}'::jsonb
);

create table if not exists employees (
  employee_id text primary key,
  store_gid text references stores(gid) on delete set null,
  periods jsonb default '{}'::jsonb,
  today jsonb
);

-- ---------- Tabelas de configuração (escritas pela equipa no painel) ----------

create table if not exists collaborator_info (
  employee_id text primary key,
  display_name text,
  role text
);

create table if not exists store_goals (
  store_gid text not null,
  period text not null check (period in ('MTD','YTD')),
  target numeric not null,
  primary key (store_gid, period)
);

-- ---------- Segurança (RLS) ----------
-- Qualquer pessoa com sessão iniciada (login por email) pode ler tudo.
-- A função api/refresh.py escreve em meta/stores/employees usando a
-- service_role key (que ignora RLS), por isso não precisa de política de
-- escrita aqui — essas tabelas servem apenas de cópia de reserva.
-- collaborator_info e store_goals podem ser escritos por qualquer utilizador
-- com sessão iniciada — a proteção da página "Administrador" (palavra-passe)
-- é a barreira do lado da aplicação; isto é só para bloquear acesso anónimo.

alter table meta enable row level security;
alter table stores enable row level security;
alter table employees enable row level security;
alter table collaborator_info enable row level security;
alter table store_goals enable row level security;

create policy "leitura autenticada" on meta for select using (auth.role() = 'authenticated');
create policy "leitura autenticada" on stores for select using (auth.role() = 'authenticated');
create policy "leitura autenticada" on employees for select using (auth.role() = 'authenticated');
create policy "leitura autenticada" on collaborator_info for select using (auth.role() = 'authenticated');
create policy "leitura autenticada" on store_goals for select using (auth.role() = 'authenticated');

create policy "escrita autenticada" on collaborator_info for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "escrita autenticada" on store_goals for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
