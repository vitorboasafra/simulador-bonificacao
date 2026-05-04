# 🚀 Guia de Migração para Supabase

## Passo 1: Criar Projeto Supabase

1. Acesse [supabase.com](https://supabase.com) e faça login
2. Clique em "New Project"
3. Configure:
   - **Name**: `simulador-bonificacao`
   - **Database Password**: Gere uma senha forte (salve em local seguro!)
   - **Region**: `South America (São Paulo)`
4. Copie as credenciais na aba "Project Settings > API"
   - `Project URL` → `SUPABASE_URL`
   - `anon public` → `SUPABASE_ANON_KEY`
   - `service_role` → `SUPABASE_SERVICE_ROLE_KEY`

---

## Passo 2: Criar Tabelas no Supabase

Acesse **SQL Editor** no painel Supabase e execute:

### Tabela de Usuários
```sql
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

CREATE TABLE IF NOT EXISTS usuarios (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  usuario TEXT UNIQUE NOT NULL,
  email TEXT UNIQUE NOT NULL,
  senha TEXT NOT NULL,
  admin BOOLEAN DEFAULT FALSE,
  confirmado BOOLEAN DEFAULT FALSE,
  created_at TIMESTAMP DEFAULT NOW()
);

ALTER TABLE usuarios ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Usuários podem ler seus dados" 
ON usuarios FOR SELECT 
USING (auth.uid() = id);

CREATE POLICY "Admin pode ler todos" 
ON usuarios FOR SELECT 
USING (
  EXISTS (
    SELECT 1 FROM usuarios WHERE id = auth.uid() AND admin = true
  )
);
```

### Tabela de Histórico
```sql
CREATE TABLE IF NOT EXISTS historico (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  usuario_id UUID NOT NULL REFERENCES usuarios(id) ON DELETE CASCADE,
  produtorNome TEXT NOT NULL,
  usuario TEXT NOT NULL,
  totalGeralFmt TEXT NOT NULL,
  detalhes JSONB NOT NULL,
  itens INTEGER DEFAULT 0,
  createdAt TIMESTAMP DEFAULT NOW()
);

CREATE INDEX idx_historico_usuario ON historico(usuario_id);
CREATE INDEX idx_historico_data ON historico(createdAt DESC);

ALTER TABLE historico ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Usuários veem seu histórico" 
ON historico FOR SELECT 
USING (usuario_id = auth.uid());

CREATE POLICY "Admin vê todos" 
ON historico FOR SELECT 
USING (
  EXISTS (
    SELECT 1 FROM usuarios WHERE id = auth.uid() AND admin = true
  )
);
```

### Tabela de Configurações
```sql
CREATE TABLE IF NOT EXISTS configuracoes (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  chave TEXT UNIQUE NOT NULL,
  valor TEXT,
  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW()
);

ALTER TABLE configuracoes ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Qualquer um pode ler" 
ON configuracoes FOR SELECT USING (true);

CREATE POLICY "Apenas admin edita" 
ON configuracoes FOR UPDATE USING (
  EXISTS (
    SELECT 1 FROM usuarios WHERE id = auth.uid() AND admin = true
  )
);
```

---

## Passo 3: Instalar Dependências

```bash
npm install @supabase/supabase-js dotenv
```

---

## Passo 4: Configurar .env

1. Copie `.env.example` para `.env`
2. Preencha com suas credenciais Supabase
3. Gere uma chave secreta forte:
   ```bash
   node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"
   ```

---

## Passo 5: Migrar Dados Locais (Opcional)

Execute o script de migração:

```bash
node scripts/migrate-to-supabase.js
```

Isso transfere:
- Usuários de `usuarios.json`
- Histórico de `data/historico.json`

---

## Passo 6: Usar Servidor com Supabase

Renomeie `server.js` para `server-local.js` e use `server-supabase.js`:

```bash
# Copiar server-supabase.js para server.js
cp server-supabase.js server.js
```

---

## Passo 7: Testar Localmente

```bash
npm run dev
```

Teste login e histórico normalmente.

---

## Passo 8: Deploy em Vercel/Railway

### Opção A: Vercel
```bash
npm install -g vercel
vercel
```

### Opção B: Railway
```bash
npm install -g railway
railway login
railway up
```

Adicione variáveis de ambiente no painel de deployment.

---

## 🔒 Segurança

- ✅ Nunca comite `.env` (já está em `.gitignore`)
- ✅ `SUPABASE_SERVICE_ROLE_KEY` apenas no servidor
- ✅ `SUPABASE_ANON_KEY` pode ser pública (frontend)
- ✅ Row Level Security ativado em todas as tabelas

---

## Rollback (Voltar ao Sistema Anterior)

```bash
# Restaurar servidor local
cp server-local.js server.js
npm run dev
```

---

## Suporte

Documentação Supabase: https://supabase.com/docs
