# Painel Comercial A&B

Painel de vendas, objetivos e comissões para as lojas Multiopticas do grupo Aguiar&Biscaia.

## Como funciona (sem automação agendada)

1. Alguém substitui manualmente os 3 ficheiros Excel numa pasta OneDrive partilhada, todos os dias — sempre com o mesmo nome (para o link de partilha se manter estável).
2. Sempre que alguém abre o painel, ele chama automaticamente `/api/refresh` — uma função que corre **na hora** (não agendada), vai buscar os 3 ficheiros ao OneDrive, trata-os, e devolve o resultado.
3. Essa função também guarda uma cópia do resultado na base de dados (Supabase), para servir de reserva se uma chamada futura falhar (rede lenta, ficheiro em falta, etc.) — nesse caso o painel mostra um aviso e os últimos dados válidos.

Não há GitHub Actions, Power Automate, nem nada a correr sozinho em segundo plano.

## Estrutura do projeto

```
index.html                       — o painel (site completo)
api/refresh.py                   — função serverless: busca + trata os dados, sob pedido
api/parser_core.py                — lógica de leitura dos 3 Excel (partilhada)
supabase/schema.sql               — tabelas e permissões da base de dados
scripts/process_files_manual.py   — utilitário manual (testes, primeira carga)
requirements.txt                  — dependências Python (usadas pela função serverless)
vercel.json                       — configuração da função (duração máxima)
```

## Configurar (primeira vez)

### 1. Base de dados (Supabase)
1. Cria um projeto em [supabase.com](https://supabase.com).
2. **SQL Editor** → cola o conteúdo de `supabase/schema.sql` → **Run**.
3. Em **Project Settings → API**, copia o **Project URL**, a **anon public key** e a **service_role key**.

### 2. Links dos ficheiros no OneDrive
1. Cria uma pasta dedicada no OneDrive (ex: `/PainelAB/diario/`) com os 3 ficheiros, sempre com o mesmo nome (ex: `dashboard.xlsx`, `workorders.xlsx`, `sales_report.xlsx`).
2. Para cada um: botão direito → **Partilhar** → "Pessoas da organização com o link" → **Copiar link**.
3. **Testa cada link numa janela anónima do navegador** — tem de abrir sem pedir login. Se pedir login, a função `/api/refresh` não vai conseguir aceder sem alterações adicionais.

### 3. O site (`index.html`)
Abre `index.html` e substitui no topo do `<script>`:
```js
const SUPABASE_URL = 'COLOCA_AQUI_O_TEU_PROJECT_URL';
const SUPABASE_ANON_KEY = 'COLOCA_AQUI_A_TUA_ANON_KEY';
```

### 4. Publicar no Vercel
1. Envia este projeto para um repositório GitHub (privado, recomendado).
2. Em [vercel.com](https://vercel.com) → **New Project** → importa o repositório → **Deploy**.
3. Em **Settings → Environment Variables**, adiciona:

| Variável | Valor |
|---|---|
| `ONEDRIVE_DASHBOARD_URL` | link de partilha do ficheiro do dashboard |
| `ONEDRIVE_WORKORDERS_URL` | link de partilha do ficheiro de work orders |
| `ONEDRIVE_SALES_REPORT_URL` | link de partilha do ficheiro de sales report |
| `SUPABASE_URL` | o mesmo Project URL de cima |
| `SUPABASE_SERVICE_ROLE_KEY` | a **service_role key** (nunca a anon key aqui) |

4. Re-implanta (Redeploy) depois de adicionar as variáveis, para que a função as passe a usar.

> ⚠️ **Sobre o tempo de execução:** ler e tratar 3 ficheiros Excel pode demorar alguns segundos, e o arranque "a frio" do Python com a biblioteca pandas também demora o seu tempo. Configurei `vercel.json` para dar até 60 segundos à função — se o vosso plano gratuito não permitir esse valor, o deploy avisa-vos e é preciso ajustar (ou aceitar o limite do plano, o que normalmente ainda chega).

### 5. Login da equipa
Login por email (link mágico, sem password) via Supabase — já funciona assim que o site estiver publicado.

### 6. Testar
Abre o site publicado, faz login, e o painel deve mostrar "A atualizar dados a partir do OneDrive…" por alguns segundos e depois os dados de hoje. Há também um botão **"🔄 Atualizar dados"** na barra lateral para forçar uma nova busca sem recarregar a página.

## Nomes, funções e objetivos
Cada pessoa com login pode ver os dados. A secção **🔒 Administrador** (protegida com a palavra-passe já definida no código) permite editar nomes, funções e objetivos — ficam guardados na base de dados, visíveis para toda a gente, e não são afetados pela atualização diária dos dados de vendas.
