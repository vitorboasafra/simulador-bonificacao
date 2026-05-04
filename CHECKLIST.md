# ✅ Checklist Migração para Supabase + Deploy

## 📋 Passo 1: Preparar Supabase

- [ ] Criar conta em [supabase.com](https://supabase.com)
- [ ] Criar novo projeto
- [ ] Executar SQL do `SETUP_SUPABASE.md` (seção "Criar Tabelas")
- [ ] Copiar credenciais:
  - [ ] `SUPABASE_URL`
  - [ ] `SUPABASE_ANON_KEY`
  - [ ] `SUPABASE_SERVICE_ROLE_KEY`

## 🔧 Passo 2: Configurar Projeto Local

```bash
# Instalar dependências
npm install

# Copiar arquivo de ambiente
cp .env.example .env

# Editar .env com suas credenciais Supabase
# Gerar SESSION_SECRET:
node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"
```

- [ ] `.env` preenchido corretamente
- [ ] `SMTP_*` configurado (opcional para email)

## 🧪 Passo 3: Testar Localmente

```bash
# Testar com servidor Supabase
npm run dev
```

- [ ] Acesse `http://localhost:3000`
- [ ] Login funciona
- [ ] Histórico salva no Supabase
- [ ] Gerar PDF funciona

## 🚀 Passo 4: Migrar Dados Existentes (Opcional)

```bash
# Se tem dados antigos em usuarios.json e historico.json
npm run migrate
```

- [ ] Usuários migrados para Supabase
- [ ] Histórico migrado para Supabase
- [ ] Dados ainda funcionam após migração

## 📦 Passo 5: Preparar Para Deploy

```bash
# Inicializar git (se não tiver)
git init
git add .
git commit -m "Initial commit - Supabase ready"

# Criar repositório no GitHub
# Suba para GitHub
git remote add origin https://github.com/seu-usuario/simulador-bonificacao.git
git branch -M main
git push -u origin main
```

- [ ] Repositório criado no GitHub
- [ ] `.env` NÃO está commitado (verifique)
- [ ] `server.js` apontando para `server-supabase.js`

## 🌐 Passo 6: Deploy (Escolha uma opção)

### Opção A: Railway (Recomendado)

```bash
npm install -g railway
railway login
railway up
```

- [ ] Repositório conectado
- [ ] Variáveis de ambiente adicionadas no painel
- [ ] Deploy concluído
- [ ] URL obtida: `https://seu-projeto.railway.app`

### Opção B: Vercel

```bash
npm install -g vercel
vercel
```

- [ ] Repositório conectado
- [ ] Variáveis de ambiente adicionadas
- [ ] Deploy concluído
- [ ] URL obtida: `https://seu-projeto.vercel.app`

### Opção C: Render

- [ ] Criar conta em [render.com](https://render.com)
- [ ] "New Web Service" → conectar GitHub
- [ ] Configurar variáveis de ambiente
- [ ] Deploy automático
- [ ] URL obtida: `https://seu-projeto.onrender.com`

## 🔍 Passo 7: Validar Deploy

- [ ] Acessar URL pública
- [ ] Login funciona
- [ ] Cálculos funcionam
- [ ] PDF baixa corretamente
- [ ] Histórico salva

## 🎉 Pronto para Usar!

Seu projeto está online e qualquer pessoa pode acessar usando a URL de deploy.

---

## 📝 Próximas Melhorias (Opcional)

- [ ] Adicionar hash de senha (bcrypt)
- [ ] Autenticação social (Google/GitHub)
- [ ] Domínio customizado
- [ ] SSL/HTTPS (já incluído)
- [ ] Analytics

## 🆘 Problemas Comuns

| Erro | Solução |
|------|---------|
| "SUPABASE_URL não definido" | Verificar variáveis de ambiente no painel |
| "Conexão recusada" | Aguardar 2-3 min para servidor iniciar |
| "Usuário não encontrado" | Rodar `npm run migrate` para transferir usuários |
| "Erro ao fazer login" | Verificar senha no Supabase (tabela usuarios) |

---

**Documentação Completa**: Veja `SETUP_SUPABASE.md` e `DEPLOY.md`
