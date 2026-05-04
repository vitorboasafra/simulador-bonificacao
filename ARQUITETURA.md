# 🏗️ Arquitetura Supabase + Deploy

## Antes (Local)

```
┌─────────────────────────┐
│   Seu Computador        │
│  ┌───────────────────┐  │
│  │ Node.js Express   │  │
│  │ + Dados em JSON   │  │
│  │ + PDFs local      │  │
│  └───────────────────┘  │
└─────────────────────────┘

❌ Não funciona online
❌ Dados perdidos se reiniciar
❌ Difícil compartilhar
```

---

## Depois (Com Supabase + Deploy)

```
                ┌─────────────────────────────┐
                │   Internet (Pública)        │
                │  https://seu-projeto.app    │
                └──────────────┬──────────────┘
                               │
                ┌──────────────▼──────────────┐
                │   Railway / Vercel / Render │
                │  ┌──────────────────────┐  │
                │  │ Node.js Express      │  │
                │  │ + server-supabase.js │  │
                │  │ + Gera PDFs          │  │
                │  └──────────────────────┘  │
                └──────────────┬──────────────┘
                               │
        ┌──────────────────────┘
        │
        ▼
    ┌──────────────────────────┐
    │   Supabase (Cloud)       │
    │ ┌────────────────────┐   │
    │ │ PostgreSQL:        │   │
    │ │ • usuarios         │   │
    │ │ • historico        │   │
    │ │ • configuracoes    │   │
    │ └────────────────────┘   │
    │ + Backups automáticos    │
    │ + Row Level Security     │
    └──────────────────────────┘

✅ Acessa de qualquer lugar
✅ Dados persistentes
✅ Escalável
✅ Seguro
```

---

## Fluxo de Autenticação

```
1. Usuário acessa https://seu-projeto.app
                            ▼
2. Frontend carrega (index.html + painel.html)
                            ▼
3. Faz POST /api/login com usuario + senha
                            ▼
4. server-supabase.js valida contra Supabase
                            ▼
5. Supabase retorna dados do usuário
                            ▼
6. Session criada localmente
                            ▼
7. Acesso às rotas autenticadas ✅
```

---

## Fluxo de Cálculo de Bonificação

```
Frontend
   │
   ├─ 1. Seleciona fornecedor/centros/cultivares
   │
   ├─ 2. POST /api/filtros
   │      └─ server-supabase.js lê planilha XLSX
   │         └─ Retorna opções
   │
   ├─ 3. POST /api/calcular
   │      └─ Processa dados
   │      └─ Salva em histórico (Supabase)
   │
   └─ 4. POST /api/relatorio
          └─ Gera PDF
          └─ Download no navegador
```

---

## Dados no Supabase

### Tabela: usuarios
```
id         | usuario      | email                | admin | confirmado
────────────────────────────────────────────────────────────────────
uuid...    | admin        | admin@...            | true  | true
uuid...    | Torre        | torre@...            | true  | true
uuid...    | novo_user    | novo@...             | false | true
```

### Tabela: historico
```
id         | usuario_id | usuario  | produtorNome | totalGeralFmt | detalhes (JSON)
──────────────────────────────────────────────────────────────────────────────────
uuid...    | uuid...    | admin    | Boa Safra    | "R$ 5.000,00" | {...}
uuid...    | uuid...    | Torre    | Boa Safra    | "R$ 3.500,00" | {...}
```

---

## Variáveis de Ambiente

### No Seu PC (`.env`)
```
SUPABASE_URL=https://xxx.supabase.co
SUPABASE_SERVICE_ROLE_KEY=eyxxx... (SECRETO!)
SESSION_SECRET=aleatorio32chars...
```

### No Servidor Deploy (Railway/Vercel)
```
SUPABASE_URL=https://xxx.supabase.co
SUPABASE_SERVICE_ROLE_KEY=eyxxx... (SECRETO!)
SESSION_SECRET=aleatorio32chars...
```

⚠️ **Nunca** coloque em `.env` commitado no GitHub!

---

## Custos Mensais

| Serviço | Preço |
|---------|-------|
| Supabase | **Grátis** (até 500k req/mês) |
| Railway | **~$5** (ou $5 crédito inicial) |
| Vercel | **Grátis** |
| Render | **$7** (mínimo) |
| Domínio customizado | **~$10/ano** (opcional) |
| **TOTAL** | **~$5-7/mês OU GRÁTIS** |

---

## Comandos Úteis

```bash
# Testar localmente
npm run dev

# Instalar dependências
npm install

# Migrar dados antigos
npm run migrate

# Conectar com Supabase em produção
npm run supabase

# Ver logs (Railway)
railway logs

# Ver logs (Vercel)
vercel logs
```

---

## Próximos Passos

1. **Hoje**: Seguir `CHECKLIST.md`
2. **Amanhã**: Validar em produção
3. **Semana que vem**:
   - [ ] Adicionar hash de senha (bcrypt)
   - [ ] Autenticação OAuth
   - [ ] Domínio customizado
   - [ ] Monitoramento (Sentry/LogRocket)

---

## 📞 Suporte Rápido

- Supabase Down? → [status.supabase.com](https://status.supabase.com)
- Railway Down? → [status.railway.app](https://status.railway.app)
- Vercel Down? → [status.vercel.com](https://status.vercel.com)

---

**Seu projeto está pronto para escala! 🚀**
