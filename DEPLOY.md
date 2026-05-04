# 🚀 Deploy do Simulador de Bonificação

Guia rápido para publicar seu projeto online.

---

## Opção 1: **Railway** (Recomendado - Mais Fácil)

### 1. Criar conta
- Acesse [railway.app](https://railway.app)
- Faça login com GitHub

### 2. Conectar seu repositório
```bash
# No seu projeto local, inicialize git
git init
git add .
git commit -m "Initial commit"
```

- Suba para GitHub (crie um repositório público)

### 3. Deploy na Railway
1. No painel Railway: "New Project" → "Deploy from GitHub"
2. Selecione seu repositório
3. Adicione variáveis de ambiente:
   - `SUPABASE_URL`
   - `SUPABASE_SERVICE_ROLE_KEY`
   - `SESSION_SECRET`
   - `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS`

4. Clique em "Deploy"

**URL será**: `https://seu-projeto.railway.app`

---

## Opção 2: **Vercel** (Para Frontend + API)

### 1. Instale Vercel CLI
```bash
npm install -g vercel
```

### 2. Deploy
```bash
vercel
```

### 3. Configure no painel
- Environment Variables: adicione as mesmas do Railway

**URL será**: `https://seu-projeto.vercel.app`

---

## Opção 3: **Heroku** (Requer Cartão de Crédito)

Heroku descontinuou plano gratuito em 2022. Não recomendado.

---

## Opção 4: **Render.com** (Alternativa Simples)

### 1. Criar conta
- Acesse [render.com](https://render.com)

### 2. New → Web Service
- Conecte seu repositório GitHub
- Defina:
  - **Build Command**: `npm install`
  - **Start Command**: `npm start`
  - **Environment Variables**: mesmas do Railway

### 3. Deploy automático

**URL será**: `https://seu-projeto.onrender.com`

---

## ✅ Checklist Pré-Deploy

- [ ] Variáveis de ambiente configuradas
- [ ] Supabase com tabelas criadas
- [ ] `.env` NÃO está commitado (verifique `.gitignore`)
- [ ] `server.js` → aponta para `server-supabase.js`
- [ ] Testes locais funcionando: `npm run dev`

---

## 🔒 Segurança

```bash
# Gere uma chave secreta forte
node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"
```

Use esse valor em `SESSION_SECRET`.

---

## 📊 Monitorar Logs

### Railway
```bash
vercel logs
```

### Vercel
```bash
vercel logs
```

### Render
Painel → Logs

---

## 💰 Custos Estimados

| Serviço | Custo |
|---------|-------|
| Supabase | Grátis até 500k requisições |
| Railway | ~$5/mês (ou grátis com $5 crédito inicial) |
| Vercel | Grátis |
| Render | $7/mês (Web Service mínimo) |
| **Total** | **~$5-7/mês OU Grátis** |

---

## 🆘 Troubleshooting

### Erro: "SUPABASE_URL não definido"
```bash
# Verifique se as variáveis estão no painel do provedor
# Não coloque aspas nos valores!
SUPABASE_URL=https://xxx.supabase.co
```

### Erro: "Banco de dados recusando conexão"
- Verifique se o IP do servidor está autorizado no Supabase
- Configurar: Supabase → Settings → Database → Connection Pooling

### Erro: "Session secret inválido"
- Regenere com o comando acima
- Redeploy

---

## 📞 Suporte

- Supabase Docs: https://supabase.com/docs
- Railway Docs: https://docs.railway.app
- Vercel Docs: https://vercel.com/docs

---

## Próximas Melhorias

- [ ] Implementar hash de senha (bcrypt)
- [ ] Adicionar autenticação OAuth (Google/GitHub)
- [ ] Backup automático de PDFs
- [ ] Rate limiting para upload
- [ ] Analytics e dashboards

