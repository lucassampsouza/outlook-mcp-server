# Outlook Calendar MCP Server

Servidor MCP que expõe operações de calendário do Microsoft Outlook via **Microsoft Graph API**. Compatível com Claude Desktop, Claude Code e qualquer cliente MCP.

## Funcionalidades

- Listar, criar, atualizar e deletar eventos de calendário
- Consultar disponibilidade (free/busy) de um ou mais usuários
- Buscar eventos por palavra-chave
- Suporte a múltiplas contas (diferentes tenants/usuários do Azure AD simultaneamente)
- Dois modos de autenticação: **Application** (client secret) e **Delegated** (refresh token via Device Code flow)

## Tools disponíveis

| Tool | Descrição |
|---|---|
| `list_accounts` | Lista todas as contas configuradas |
| `list_calendars` | Lista todos os calendários de um usuário |
| `get_calendar_events` | Busca eventos em uma janela de tempo |
| `get_event` | Retorna detalhes completos de um evento específico |
| `create_event` | Cria um novo evento (com opção de link Teams) |
| `update_event` | Atualiza campos de um evento existente |
| `delete_event` | Deleta um evento |
| `get_free_busy` | Retorna agenda de disponibilidade de um ou mais usuários |
| `search_events` | Busca eventos por palavra-chave |
| `start_device_code_auth` | Inicia o fluxo de autenticação Device Code para adicionar uma conta delegada |
| `complete_device_code_auth` | Conclui a autenticação pendente e salva o refresh token |
| `get_admin_consent_url` | Gera uma URL de consentimento de admin para um tenant |

---

## 1. Criando o App Registration no Azure

Antes de rodar o servidor, você precisa de um App Registration no Azure.

1. Acesse o [Azure Portal](https://portal.azure.com) → **Azure Active Directory** → **Registros de aplicativo** → **Novo registro**
2. Dê um nome (ex.: `outlook-mcp-server`), deixe o URI de redirecionamento em branco e clique em **Registrar**
3. Anote:
   - **ID do aplicativo (cliente)** → `AZURE_CLIENT_ID`
   - **ID do diretório (locatário)** → `AZURE_TENANT_ID`

### Permissões de API

Vá em **Permissões de API** → **Adicionar uma permissão** → **Microsoft Graph**:

| Permissão | Tipo | Finalidade |
|---|---|---|
| `Calendars.Read` | Application ou Delegated | Leitura de eventos |
| `Calendars.ReadWrite` | Application ou Delegated | Criar / atualizar / deletar eventos |
| `Calendars.Read.Shared` | Delegated | Acessar calendários compartilhados |
| `User.Read` | Delegated | Ler perfil do usuário autenticado |

- **Permissões de aplicativo** exigem clicar em **Conceder consentimento do administrador** (requer Global Admin)
- **Permissões delegadas** exigem que o usuário autentique via Device Code flow (sem necessidade de admin)

---

## 2. Modos de autenticação

### Modo 1 — Application (Client Credentials)

Ideal para acesso automatizado a qualquer caixa de correio no tenant. Requer consentimento do administrador.

1. No App Registration, vá em **Certificados e segredos** → **Novo segredo do cliente**
2. Copie o **valor** do segredo (não o ID) — esse é o `AZURE_CLIENT_SECRET`

Variáveis de ambiente necessárias:
```
AZURE_TENANT_ID=seu-tenant-id
AZURE_CLIENT_ID=seu-client-id
AZURE_CLIENT_SECRET=seu-client-secret
```

### Modo 2 — Delegated (Device Code Flow)

Ideal para uso pessoal ou quando não há consentimento de admin disponível. Não precisa de client secret.

Nesse modo, o `REFRESH_TOKEN` **não é configurado manualmente** — ele é obtido de forma dinâmica via autenticação interativa, usando as próprias tools do servidor.

**Configuração inicial** (apenas Tenant ID e Client ID):
```
AZURE_TENANT_ID=seu-tenant-id
AZURE_CLIENT_ID=seu-client-id
```

> Importante: no App Registration, habilite **"Allow public client flows"** em **Autenticação** → **Configurações avançadas**.

**Autenticando via tools MCP** (dentro do Claude):
```
1. start_device_code_auth(account_name="default", tenant_id="...", client_id="...")
   → Retorna um código e uma URL. Acesse a URL, faça login com sua conta Microsoft.

2. complete_device_code_auth(account_name="default")
   → Conclui a autenticação e salva o refresh token automaticamente no .env.
```

A partir daí o servidor gerencia o token sozinho — o refresh token é rotacionado automaticamente a cada uso e salvo no `.env`, sem necessidade de reautenticar.

**Alternativa: setup interativo via terminal**
```bash
python setup.py
```

---

## 3. Instalação e execução

### Opção A — uvx (recomendado, sem clonar o repositório)

```bash
uvx --from git+https://github.com/lucassampsouza/outlook-mcp-server outlook-mcp-server
```

### Opção B — Clone local

```bash
git clone https://github.com/lucassampsouza/outlook-mcp-server
cd outlook-mcp-server
cp .env.example .env
# Preencha o .env com suas credenciais

python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt

python server.py
```

### Modo HTTP/SSE

```bash
fastmcp run server.py --transport sse --port 8000
```

---

## 4. Conectando ao Claude

### stdio (Claude Desktop / Claude Code)

Adicione ao `claude_desktop_config.json` (geralmente em `~/Library/Application Support/Claude/` no macOS ou `%APPDATA%\Claude\` no Windows):

**Via uvx — Modo Application:**
```json
{
  "mcpServers": {
    "outlook": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/lucassampsouza/outlook-mcp-server", "outlook-mcp-server"],
      "env": {
        "AZURE_TENANT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
        "AZURE_CLIENT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
        "AZURE_CLIENT_SECRET": "seu-client-secret"
      }
    }
  }
}
```

**Via uvx — Modo Delegated (sem refresh token inicial):**
```json
{
  "mcpServers": {
    "outlook": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/lucassampsouza/outlook-mcp-server", "outlook-mcp-server"],
      "env": {
        "AZURE_TENANT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
        "AZURE_CLIENT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
      }
    }
  }
}
```

> Após adicionar o servidor, use as tools `start_device_code_auth` e `complete_device_code_auth` diretamente pelo Claude para autenticar. O refresh token será salvo automaticamente.

**Via clone local:**
```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["/caminho/absoluto/para/outlook-mcp-server/server.py"],
      "env": {
        "AZURE_TENANT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
        "AZURE_CLIENT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
        "AZURE_CLIENT_SECRET": "seu-client-secret"
      }
    }
  }
}
```

### HTTP/SSE

Inicie o servidor em modo SSE:
```bash
fastmcp run server.py --transport sse --port 8000
```

Conecte seu cliente MCP a:
```
http://localhost:8000/sse
```

Para o Claude Code, adicione ao seu arquivo de configuração MCP:
```json
{
  "mcpServers": {
    "outlook": {
      "type": "sse",
      "url": "http://localhost:8000/sse"
    }
  }
}
```

---

## 5. Múltiplas contas

O servidor suporta múltiplas contas do Azure AD simultaneamente. Cada tool aceita um parâmetro opcional `account` para selecionar qual credencial usar.

### Variáveis de ambiente

```dotenv
# Conta padrão (Application auth)
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_SECRET=seu-secret-principal

# Conta "work" — tenant diferente
ACCOUNT_WORK_TENANT_ID=yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy
ACCOUNT_WORK_CLIENT_ID=yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy
ACCOUNT_WORK_CLIENT_SECRET=seu-secret-work

# Conta "personal" — delegated auth, mesmo app registration
# (TENANT_ID e CLIENT_ID omitidos: herdam os valores AZURE_* acima)
ACCOUNT_PERSONAL_REFRESH_TOKEN=refresh-token-pessoal
```

> Quando `TENANT_ID` ou `CLIENT_ID` são omitidos em uma conta nomeada, os valores `AZURE_*` padrão são usados como fallback — útil quando múltiplos usuários compartilham o mesmo app registration.

### Uso

```
list_accounts()
  → {"accounts": ["default", "work", "personal"], ...}

get_calendar_events(user_email="alice@empresa.com")
  → usa a conta default

get_calendar_events(user_email="bob@work.com", account="work")
  → usa as credenciais da conta work

list_calendars(user_email="eu@gmail.com", account="personal")
  → usa o token delegado da conta personal
```

### Configuração multi-conta com uvx

```json
{
  "mcpServers": {
    "outlook": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/lucassampsouza/outlook-mcp-server", "outlook-mcp-server"],
      "env": {
        "AZURE_TENANT_ID": "...",
        "AZURE_CLIENT_ID": "...",
        "AZURE_CLIENT_SECRET": "...",
        "ACCOUNT_WORK_TENANT_ID": "...",
        "ACCOUNT_WORK_CLIENT_ID": "...",
        "ACCOUNT_WORK_CLIENT_SECRET": "..."
      }
    }
  }
}
```

---

## Referência de variáveis de ambiente

| Variável | Obrigatória | Descrição |
|---|---|---|
| `AZURE_TENANT_ID` | Sim | ID do diretório (tenant) do Azure AD |
| `AZURE_CLIENT_ID` | Sim | ID do aplicativo (cliente) |
| `AZURE_CLIENT_SECRET` | Modo 1 | Valor do segredo do cliente (Application auth) |
| `AZURE_REFRESH_TOKEN` | Automático | Refresh token (Delegated auth — gerado via Device Code flow) |
| `ACCOUNT_{NOME}_TENANT_ID` | Não | Tenant ID de uma conta nomeada |
| `ACCOUNT_{NOME}_CLIENT_ID` | Não | Client ID de uma conta nomeada |
| `ACCOUNT_{NOME}_CLIENT_SECRET` | Não | Client secret de uma conta nomeada (Application auth) |
| `ACCOUNT_{NOME}_REFRESH_TOKEN` | Automático | Refresh token de uma conta nomeada (Delegated auth) |

---

## Requisitos

- Python 3.10+
- `fastmcp >= 2.0.0`
- `httpx >= 0.27.0`
- `python-dotenv >= 1.0.0`
- `msal >= 1.28.0`
