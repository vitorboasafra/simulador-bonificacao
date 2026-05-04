# 🔐 Autenticação com Supabase - Documentação

## Visão Geral

O sistema de autenticação do Simulador de Bonificação está totalmente integrado com **Supabase Auth** e mantém perfis de usuários na tabela `usuarios`. Suporta:

- ✅ Criação de contas com email corporativo
- ✅ Login com Supabase Auth
- ✅ Migração automática de usuários legados
- ✅ Gerenciamento de perfis por admins
- ✅ Controle de permissões (admin/usuário)
- ✅ Verificação automática de sessão

---

## Endpoints de Autenticação

### Login

#### `POST /api/login` - Login via API
Autentica usuário e retorna dados do perfil.

**Requisição:**
```json
{
  "email": "usuario@boasafrasementes.com.br",
  "senha": "senha123"
}
```

**Resposta (200):**
```json
{
  "usuario": "usuario",
  "email": "usuario@boasafrasementes.com.br",
  "admin": false
}
```

**Erros:**
- `400`: Email ou senha ausentes
- `401`: Credenciais inválidas
- `500`: Erro ao processar

#### `POST /login` - Login via Formulário HTML
Autentica e redireciona para `/painel.html` ou `/` em caso de erro.

---

### Registro

#### `POST /api/registrar` - Registro via API
Cria novo usuário com email corporativo.

**Requisição:**
```json
{
  "usuario": "novo.usuario",
  "email": "novo.usuario@boasafrasementes.com.br",
  "senha": "senha123"
}
```

**Validações:**
- Email deve ser `@boasafrasementes.com.br`
- Senha mínimo 4 caracteres
- Email e usuário únicos

**Erros:**
- `400`: Dados inválidos
- `409`: Email ou usuário já existe
- `500`: Erro ao criar no Supabase

#### `POST /cadastro` - Registro via Formulário HTML
Cria novo usuário e redireciona com status.

---

### Logout

#### `GET /logout` - Logout com Redirecionamento
Destroi sessão e redireciona para `/`.

#### `POST /api/logout` - Logout via API
Destroi sessão e retorna confirmação.

**Resposta (200):**
```json
{
  "mensagem": "Logout realizado com sucesso."
}
```

---

### Verificação de Sessão

#### `GET /api/session` - Verificar Sessão (Requer Autenticação)
Valida sessão atual e retorna dados do usuário.

**Resposta (200):**
```json
{
  "usuario": "usuario",
  "email": "usuario@boasafrasementes.com.br",
  "admin": false
}
```

**Erros:**
- `401`: Sessão inválida ou expirada

#### `GET /api/me` - Dados do Usuário Atual (Requer Autenticação)
Similiar a `/api/session`, retorna informações do usuário logado.

---

## Endpoints de Gerenciamento de Perfis (Admin)

#### `GET /api/usuarios` - Listar Todos os Usuários (Requer Admin)
```json
{
  "usuarios": [
    {
      "id": "uuid-001",
      "usuario": "vitor",
      "email": "vitor@boasafrasementes.com.br",
      "admin": true,
      "confirmado": true,
      "created_at": "2026-04-30T10:00:00Z"
    }
  ]
}
```

#### `POST /api/usuarios` - Criar Novo Usuário (Requer Admin)
**Requisição:**
```json
{
  "usuario": "novo.usuario",
  "email": "novo@boasafrasementes.com.br",
  "senha": "senha123",
  "admin": false
}
```

#### `PUT /api/usuarios/:usuario/senha` - Alterar Senha (Requer Autenticação)
Usuário pode alterar sua própria senha, admin pode alterar de qualquer um.

**Requisição:**
```json
{
  "senha": "novaSenha123"
}
```

#### `PUT /api/usuarios/:usuario/admin` - Alterar Permissão (Requer Admin)
Promove ou remove permissão de admin.

**Requisição:**
```json
{
  "email": "usuario@boasafrasementes.com.br",
  "admin": true
}
```

#### `DELETE /api/usuarios/:usuario` - Deletar Usuário (Requer Admin)
Não permite deletar o usuário logado. Remove tanto de `usuarios` quanto de Supabase Auth.

---

## Fluxo de Autenticação

### 1️⃣ Login
```
Usuario abre index.html
   ↓
Script verifica GET /api/session
   ├─ Autenticado → Redireciona para painel.html
   └─ Não autenticado → Mostra formulário de login
   
Usuario faz login (POST /login)
   ↓
Servidor valida credenciais com Supabase Auth
   ├─ Sucesso → Cria/busca perfil em `usuarios`
   │            Cria sessão Express
   │            Redireciona para painel.html
   └─ Erro → Redireciona para index.html?erro=1
```

### 2️⃣ Acesso ao Painel
```
Usuario acessa painel.html
   ↓
Script chama GET /api/session (requireAuth)
   ├─ Autenticado → Carrega dados do painel
   └─ Não autenticado → Redireciona para index.html
```

### 3️⃣ Logout
```
Usuario clica em "Sair"
   ↓
POST /api/logout (requireAuth)
   ↓
Sessão destruída
   ↓
Redireciona para index.html
```

---

## Tabela de Usuários (Supabase)

```sql
CREATE TABLE usuarios (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  usuario TEXT UNIQUE NOT NULL,
  email TEXT UNIQUE NOT NULL,
  senha TEXT NOT NULL,
  admin BOOLEAN DEFAULT FALSE,
  confirmado BOOLEAN DEFAULT FALSE,
  created_at TIMESTAMP DEFAULT NOW()
);
```

**Campos:**
- `id`: UUID único (linkado a Supabase Auth)
- `usuario`: Nome de usuário único
- `email`: Email corporativo único
- `senha`: Armazenado (legacy compatibility)
- `admin`: True se administrador
- `confirmado`: True se email confirmado
- `created_at`: Data de criação

---

## Migração de Usuários Legados

Se um usuário tenta fazer login com credenciais armazenadas em arquivo JSON:

1. Supabase Auth tenta autenticar com `/login`
2. Se falhar, busca credenciais na tabela `usuarios` (legacy)
3. Se encontrar, cria automaticamente conta no Supabase Auth
4. Usa essa conta para próximos logins

**Função:** `signInOrMigrateLegacyUser(email, senha)`

---

## Segurança

### RLS (Row Level Security)
Tabela `usuarios` tem RLS habilitado:
- Usuários comum podem ler apenas seus próprios dados
- Admins podem ler todos os dados

### Middleware de Autenticação
- `requireAuth`: Verifica se sessão está ativa
- `requireAdmin`: Valida se usuário é admin

### Validações
- ✅ Email deve ser `@boasafrasementes.com.br`
- ✅ Senha mínimo 4 caracteres
- ✅ Usernames sanitizados
- ✅ Proteção contra SQL injection (Supabase)

---

## Teste de Fluxo

### 1. Criar Usuário de Teste
```bash
curl -X POST http://localhost:3000/api/usuarios \
  -H "Cookie: connect.sid=..." \
  -H "Content-Type: application/json" \
  -d '{
    "usuario": "teste",
    "email": "teste@boasafrasementes.com.br",
    "senha": "teste123",
    "admin": false
  }'
```

### 2. Fazer Login
```bash
curl -X POST http://localhost:3000/api/login \
  -H "Content-Type: application/json" \
  -d '{
    "email": "teste@boasafrasementes.com.br",
    "senha": "teste123"
  }'
```

### 3. Verificar Sessão
```bash
curl http://localhost:3000/api/session \
  -H "Cookie: connect.sid=..."
```

### 4. Logout
```bash
curl -X POST http://localhost:3000/api/logout \
  -H "Cookie: connect.sid=..."
```

---

## Variáveis de Ambiente

`.env` deve conter:
```
SUPABASE_URL=https://seu-project.supabase.co
SUPABASE_ANON_KEY=sua-chave-anonima
SUPABASE_SERVICE_ROLE_KEY=sua-chave-service-role
SESSION_SECRET=seu-secret-aleatorio
```

---

## Troubleshooting

### "Não autenticado" no acesso ao painel
- Sessão expirou (máximo 24h)
- Cookie de sessão foi perdido
- Solução: Fazer login novamente

### Email já existe no Supabase Auth
- Usuário foi criado mas não em `usuarios`
- Solução: Admin deleta e recria o usuário

### Usuário legado não consegue fazer login
- Credenciais armazenadas em JSON podem ter caracteres especiais
- Solução: Admin reseta senha via `/api/usuarios/:usuario/senha`

---

## Próximas Melhorias

- [ ] Email de confirmação antes de ativar conta
- [ ] Recuperação de senha por email
- [ ] Autenticação 2FA
- [ ] Auditoria de logins
- [ ] Expiração automática de sessão

