const express = require("express");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const session = require("express-session");
const multer = require("multer");
const nodemailer = require("nodemailer");
const XLSX = require("xlsx");
const PDFKitDocument = require("pdfkit");
const { PDFDocument: PDFLibDocument } = require("pdf-lib");
const { createClient } = require("@supabase/supabase-js");
require("dotenv").config();

const app = express();
const PORT = process.env.PORT || 3000;

// ============= SUPABASE SETUP =============
const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;

if (!SUPABASE_URL || !SUPABASE_KEY) {
  console.error("❌ SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY são obrigatórios!");
  process.exit(1);
}

const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

// ============= PATHS =============
const ROOT_DIR = __dirname;
const PUBLIC_DIR = path.join(ROOT_DIR, "public");
const DATA_DIR = path.join(ROOT_DIR, "data");
const UNITS_WORKBOOK = path.join(ROOT_DIR, "Unidades Boa Safra.xlsx");
const MODEL_PDF = path.join(ROOT_DIR, "NF MODELO BONIFICACAO.pdf");
const LOGIN_BACKGROUND = path.join(ROOT_DIR, "plano de fundo.png");

fs.mkdirSync(DATA_DIR, { recursive: true });

// ============= CONFIG =============
const SESSION_SECRET = process.env.SESSION_SECRET || "boa-safra-secret";
const HISTORY_TTL_MS = 24 * 60 * 60 * 1000;
const TOKEN_TTL_MS = 24 * 60 * 60 * 1000;
const SMTP_HOST = process.env.SMTP_HOST || "";
const SMTP_PORT = Number(process.env.SMTP_PORT || "587");
const SMTP_USER = process.env.SMTP_USER || "";
const SMTP_PASS = process.env.SMTP_PASS || "";
const EMAIL_FROM = process.env.EMAIL_FROM || "no-reply@boasafrasementes.com.br";
const EMAIL_DOMAIN = "@boasafrasementes.com.br";

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (![".xlsx", ".xls"].includes(ext)) {
      return cb(new Error("Envie uma planilha nos formatos .xlsx ou .xls."));
    }
    return cb(null, true);
  }
});

// ============= MIDDLEWARE =============
app.use(express.static(PUBLIC_DIR));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(session({
  secret: SESSION_SECRET,
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 1000 * 60 * 60 * 24 }
}));

app.get("/health", (_req, res) => {
  res.status(200).json({ status: "ok" });
});

// ============= HELPERS =============
function formatNumberPtBr(n) {
  const num = Number(n) || 0;
  return num.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function isEmailValid(email) {
  return /^[^\s@]+@boasafrasementes\.com\.br$/i.test(String(email || "").trim());
}

function normalizeBoolean(value) {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value === 1;
  return ["true", "1", "sim", "yes"].includes(String(value || "").trim().toLowerCase());
}

function isAdminProfile(profile) {
  return normalizeBoolean(profile?.admin);
}

function requireAuth(req, res, next) {
  if (!req.session.user) {
    return res.status(401).json({ erro: "Não autenticado." });
  }
  return next();
}

async function isAdminUser(usuario) {
  const { data, error } = await supabase
    .from("usuarios")
    .select("admin")
    .eq("usuario", usuario)
    .single();

  if (error || !data) return false;
  return normalizeBoolean(data.admin);
}

function requireAdmin(req, res, next) {
  return isAdminUser(req.session.user).then(isAdmin => {
    if (!isAdmin) {
      return res.status(403).json({ erro: "Acesso permitido apenas para administrador." });
    }
    return next();
  });
}

async function createSupabaseAuthUser({ email, senha, usuario }) {
  const { data, error } = await supabase.auth.admin.createUser({
    email,
    password: senha,
    email_confirm: true,
    user_metadata: { usuario }
  });

  return { user: data?.user, error };
}

async function signInSupabaseAuth({ email, senha }) {
  const { data, error } = await supabase.auth.signInWithPassword({
    email,
    password: senha
  });

  return { user: data?.user, session: data?.session, error };
}

async function findOrCreateUserProfile(email, usuario, senha = "") {
  const normalizedEmail = String(email || "").trim().toLowerCase();
  if (!normalizedEmail) return null;

  const { data: existingByEmail } = await supabase
    .from("usuarios")
    .select("*")
    .eq("email", normalizedEmail)
    .maybeSingle();

  if (existingByEmail) {
    return existingByEmail;
  }

  let candidate = String(usuario || "").trim() || normalizedEmail.split("@")[0];
  if (!candidate) candidate = normalizedEmail;

  let uniqueCandidate = candidate;
  let counter = 1;
  while (true) {
    const { data: existingByUsuario } = await supabase
      .from("usuarios")
      .select("id")
      .eq("usuario", uniqueCandidate)
      .maybeSingle();

    if (!existingByUsuario) break;
    uniqueCandidate = `${candidate}${counter++}`;
  }

  const { data, error } = await supabase.from("usuarios").insert({
    usuario: uniqueCandidate,
    email: normalizedEmail,
    senha,
    confirmado: true,
    admin: false
  }).select("*").maybeSingle();

  if (error) return null;
  return data;
}

async function getCurrentProfile(req) {
  let query = supabase.from("usuarios").select("*");
  if (req.session.userId) {
    query = query.eq("id", req.session.userId);
  } else {
    query = query.eq("usuario", req.session.user);
  }

  const { data, error } = await query.maybeSingle();
  if (error || !data) return null;

  req.session.user = data.usuario;
  req.session.userId = data.id;
  return data;
}

function isSupabaseDuplicateAuthError(error) {
  const message = String(error?.message || "").toLowerCase();
  return message.includes("already registered") || message.includes("duplicate") || message.includes("already exists") || message.includes("already been taken");
}

async function signInOrMigrateLegacyUser(email, senha) {
  const { user, error } = await signInSupabaseAuth({ email, senha });
  if (user) {
    return { user, error: null };
  }

  const { data: legacyUser, error: legacyError } = await supabase
    .from("usuarios")
    .select("*")
    .eq("email", email)
    .eq("senha", senha)
    .maybeSingle();

  if (legacyError) {
    return { user: null, error: legacyError };
  }

  if (!legacyUser) {
    return { user: null, error };
  }

  const { user: authUser, error: createError } = await createSupabaseAuthUser({
    email,
    senha,
    usuario: legacyUser.usuario
  });

  if (createError && !isSupabaseDuplicateAuthError(createError)) {
    return { user: null, error: createError };
  }

  return {
    user: authUser || { email, user_metadata: { usuario: legacyUser.usuario } },
    error: null
  };
}

// ============= DATABASE FUNCTIONS =============

// Planilha atual
function readWorkbookMeta() {
  const WORKBOOK_META = path.join(DATA_DIR, "planilha.json");
  if (!fs.existsSync(WORKBOOK_META)) return null;
  try {
    return JSON.parse(fs.readFileSync(WORKBOOK_META, "utf8"));
  } catch (_err) {
    return null;
  }
}

function findWorkbookPath() {
  const candidates = ["zmm113.xlsx", "zmm113.XLSX", "zmm113.xls", "zmm113.XLS"];
  for (const file of candidates) {
    const absolute = path.join(ROOT_DIR, file);
    if (fs.existsSync(absolute)) return absolute;
  }
  return null;
}

function findUploadedWorkbookPath() {
  const meta = readWorkbookMeta();
  if (meta && meta.storedName) {
    const metaPath = path.join(DATA_DIR, meta.storedName);
    if (fs.existsSync(metaPath)) return metaPath;
  }
  for (const file of ["planilha-atual.xlsx", "planilha-atual.xls"]) {
    const absolute = path.join(DATA_DIR, file);
    if (fs.existsSync(absolute)) return absolute;
  }
  return null;
}

function getWorkbookInfo() {
  const uploaded = findUploadedWorkbookPath();
  const meta = readWorkbookMeta();

  if (uploaded) {
    return {
      origem: "upload",
      arquivo: meta?.originalName || path.basename(uploaded),
      enviadoEm: meta?.uploadedAt || null
    };
  }

  const fallback = findWorkbookPath();
  if (fallback) {
    return {
      origem: "padrão",
      arquivo: path.basename(fallback),
      enviadoEm: null
    };
  }

  return { origem: "nenhuma", arquivo: null, enviadoEm: null };
}

function getWorkbook() {
  const workbookPath = findUploadedWorkbookPath() || findWorkbookPath();
  if (!workbookPath) return null;

  try {
    return XLSX.readFile(workbookPath);
  } catch (_err) {
    return null;
  }
}

// ============= AUTH ROUTES =============

app.post("/api/login", async (req, res) => {
  const email = String(req.body.email || "").trim().toLowerCase();
  const senha = String(req.body.senha || "").trim();

  if (!email || !senha) {
    return res.status(400).json({ erro: "E-mail e senha são obrigatórios." });
  }

  try {
    const { user, error } = await signInOrMigrateLegacyUser(email, senha);
    if (error || !user) {
      return res.status(401).json({ erro: "Usuário ou senha incorretos." });
    }

    const profile = await findOrCreateUserProfile(email, user.user_metadata?.usuario || "", senha);
    if (!profile) {
      return res.status(500).json({ erro: "Erro ao buscar o perfil do usuário." });
    }

    req.session.user = profile.usuario;
    req.session.userId = profile.id;
    res.json({ usuario: profile.usuario, email: profile.email, admin: isAdminProfile(profile) });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao processar login." });
  }
});

app.post("/api/registrar", async (req, res) => {
  const usuario = String(req.body.usuario || "").trim();
  const email = String(req.body.email || "").trim().toLowerCase();
  const senha = String(req.body.senha || "").trim();

  if (!usuario || !email || !senha) {
    return res.status(400).json({ erro: "Todos os campos são obrigatórios." });
  }

  if (!isEmailValid(email)) {
    return res.status(400).json({ erro: "Email deve ser @boasafrasementes.com.br" });
  }

  if (senha.length < 4) {
    return res.status(400).json({ erro: "Senha deve ter mínimo 4 caracteres." });
  }

  try {
    const { user, error } = await createSupabaseAuthUser({ email, senha, usuario });
    if (error) {
      if (isSupabaseDuplicateAuthError(error)) {
        return res.status(409).json({ erro: "Email já existe." });
      }
      return res.status(500).json({ erro: "Erro ao criar usuário no Supabase." });
    }

    const profile = await findOrCreateUserProfile(email, usuario, senha);
    if (!profile) {
      return res.status(500).json({ erro: "Erro ao criar perfil do usuário." });
    }

    res.json({ mensagem: "Usuário criado com sucesso!", usuario: profile.usuario, email: profile.email });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao processar registro." });
  }
});

app.post("/login", async (req, res) => {
  const email = String(req.body.email || "").trim().toLowerCase();
  const senha = String(req.body.senha || "").trim();

  if (!email || !senha) {
    return res.redirect("/?erro=1");
  }

  try {
    const { user, error } = await signInSupabaseAuth({ email, senha });
    if (error || !user) {
      return res.redirect("/?erro=1");
    }

    const profile = await findOrCreateUserProfile(email, user.user_metadata?.usuario || "", senha);
    if (!profile) {
      return res.redirect("/?erro=1");
    }

    req.session.user = profile.usuario;
    req.session.userId = profile.id;
    return res.redirect("/painel.html");
  } catch (error) {
    return res.redirect("/?erro=1");
  }
});

app.post("/cadastro", async (req, res) => {
  const usuario = String(req.body.usuario || "").trim();
  const email = String(req.body.email || "").trim().toLowerCase();
  const senha = String(req.body.senha || "").trim();

  if (!usuario || !email || !senha) {
    return res.redirect("/?cadastro=erro");
  }

  if (!isEmailValid(email)) {
    return res.redirect("/?cadastro=erro");
  }

  if (senha.length < 4) {
    return res.redirect("/?cadastro=erro");
  }

  try {
    const { user, error } = await createSupabaseAuthUser({ email, senha, usuario });
    if (error) {
      if (isSupabaseDuplicateAuthError(error)) {
        return res.redirect("/?cadastro=emailexistente");
      }
      return res.redirect("/?cadastro=erro");
    }

    const profile = await findOrCreateUserProfile(email, usuario, senha);
    if (!profile) {
      return res.redirect("/?cadastro=erro");
    }

    return res.redirect("/?cadastro=ok");
  } catch (error) {
    return res.redirect("/?cadastro=erro");
  }
});

app.get("/api/me", requireAuth, async (req, res) => {
  try {
    const { data } = await supabase
      .from("usuarios")
      .select("*")
      .eq("usuario", req.session.user)
      .single();

    if (!data) {
      return res.status(404).json({ erro: "Usuário não encontrado." });
    }

    res.json({ usuario: data.usuario, email: data.email, admin: isAdminProfile(data) });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao buscar dados." });
  }
});

// ============= DATA ROUTES =============

app.get("/api/fornecedores", requireAuth, (_req, res) => {
  try {
    const workbook = getWorkbook();
    if (!workbook) {
      return res.json({ fornecedores: [] });
    }

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet);

    const fornecedores = [...new Set(rows.map(r => String(r["Fornecedor"] || "").trim()).filter(Boolean))]
      .sort((a, b) => a.localeCompare(b, "pt-BR"));

    res.json({ fornecedores });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao buscar fornecedores." });
  }
});

app.get("/api/filtros", requireAuth, (req, res) => {
  const fornecedor = String(req.query.fornecedor || "").trim();
  const centrosParam = String(req.query.centros || "").trim();

  try {
    const workbook = getWorkbook();
    if (!workbook) {
      return res.json({ centros: [], cultivares: [] });
    }

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet);

    const filtered = rows.filter(r => {
      const fornecedorMatch = String(r["Fornecedor"] || "").trim().toLowerCase() === fornecedor.toLowerCase();
      if (centrosParam) {
        const centros = centrosParam.split(",").map(c => c.trim().toLowerCase());
        const centroMatch = centros.includes(String(r["Centro"] || "").trim().toLowerCase());
        return fornecedorMatch && centroMatch;
      }
      return fornecedorMatch;
    });

    const centros = [...new Set(filtered.map(r => String(r["Centro"] || "").trim()).filter(Boolean))];
    const cultivares = [...new Set(filtered.map(r => String(r["Descrição (Material)"] || "").trim()).filter(Boolean))];

    res.json({
      centros: centros.sort((a, b) => a.localeCompare(b, "pt-BR")),
      cultivares: cultivares.sort((a, b) => a.localeCompare(b, "pt-BR"))
    });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao buscar filtros." });
  }
});

app.post("/api/calcular", requireAuth, (req, res) => {
  const { fornecedorNome, centros, cultivares, modoGlobal, globalPerc, globalValor } = req.body;

  if (!fornecedorNome || !centros?.length || !cultivares?.length) {
    return res.status(400).json({ erro: "Dados incompletos." });
  }

  try {
    const workbook = getWorkbook();
    if (!workbook) {
      return res.status(500).json({ erro: "Planilha não encontrada." });
    }

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet);

    const combinacoes = rows.filter(row => {
      const rowFornecedor = String(row["Fornecedor"] || "").trim().toLowerCase();
      const rowCentro = String(row["Centro"] || "").trim();
      const rowCultivar = String(row["Descrição (Material)"] || "").trim();

      return rowFornecedor === fornecedorNome.toLowerCase() &&
             centros.includes(rowCentro) &&
             cultivares.includes(rowCultivar);
    }).map(row => ({
      inscricaoEstadual: String(row["Inscrição Estadual"] || ""),
      centro: String(row["Centro"] || ""),
      cultivar: String(row["Descrição (Material)"] || ""),
      pesoBase: Number(row["Peso Base (kg)"] || 0),
      pesoBaseFmt: formatNumberPtBr(row["Peso Base (kg)"] || 0),
      nfs: String(row["NF"] || "")
    }));

    res.json({ combinacoes, payloadPdf: { usuario: req.session.user, produtorNome: fornecedorNome, detalhes: combinacoes } });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao calcular." });
  }
});

// ============= HISTÓRICO =============

app.post("/api/historico", requireAuth, async (req, res) => {
  const { totalGeralFmt, detalhes, itens } = req.body;

  try {
    const profile = await getCurrentProfile(req);
    if (!profile) {
      return res.status(401).json({ erro: "Sessao invalida." });
    }

    const { error } = await supabase.from("historico").insert({
      usuario_id: profile.id,
      usuario: profile.usuario,
      produtorNome: "Cálculo de Bonificação",
      totalGeralFmt,
      detalhes,
      itens: detalhes?.length || 0
    });

    if (error) throw error;

    res.json({ mensagem: "Cálculo salvo no histórico." });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao salvar histórico." });
  }
});

app.get("/api/historico", requireAuth, async (req, res) => {
  try {
    const profile = await getCurrentProfile(req);
    if (!profile) {
      return res.status(401).json({ erro: "Sessao invalida." });
    }

    let query = supabase
      .from("historico")
      .select("*")
      .order("createdAt", { ascending: false })
      .limit(20);

    if (!isAdminProfile(profile)) {
      query = query.eq("usuario_id", profile.id);
    }

    const { data, error } = await query;
    if (error) throw error;

    res.json({ historico: data || [] });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao buscar histórico." });
  }
});

// ============= PDF =============

app.post("/api/relatorio", (req, res) => {
  const { usuario, produtorNome, detalhes } = req.body;

  const doc = new PDFKitDocument({ size: "A4", margin: 30 });
  const chunks = [];

  doc.on("data", (chunk) => chunks.push(chunk));
  doc.on("end", () => {
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", "attachment; filename=Bonificacao.pdf");
    res.send(Buffer.concat(chunks));
  });
  doc.on("error", () => {
    res.status(500).json({ erro: "Erro ao gerar PDF." });
  });

  doc.fontSize(18).fillColor("#4A7C2A").text("RELATÓRIO DE BONIFICAÇÃO", { align: "center" });
  doc.moveDown(0.7);
  doc.strokeColor("#E0E3E8").lineWidth(1).moveTo(30, doc.y).lineTo(565, doc.y).stroke();
  doc.moveDown(0.7);

  doc.fillColor("#222222").fontSize(10)
    .text(`Produtor: ${produtorNome}`)
    .text(`Usuário: ${usuario}`)
    .text(`Data Emissão: ${new Date().toLocaleString("pt-BR")}`);
  doc.moveDown(0.9);

  for (const item of detalhes) {
    doc.fontSize(10).text(`Inscrição: ${item.inscricaoEstadual} | Centro: ${item.centro} | Material: ${item.cultivar}`);
    doc.text(`Peso Base: ${item.pesoBaseFmt} kg`);
    doc.moveDown(0.5);
  }

  doc.end();
});

// ============= UPLOAD =============

app.post("/api/planilha", requireAuth, requireAdmin, upload.single("planilha"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ erro: "Nenhum arquivo enviado." });
  }

  try {
    const filename = `planilha-${Date.now()}.xlsx`;
    const filepath = path.join(DATA_DIR, filename);
    fs.writeFileSync(filepath, req.file.buffer);

    fs.writeFileSync(path.join(DATA_DIR, "planilha.json"), JSON.stringify({
      storedName: filename,
      originalName: req.file.originalname,
      uploadedAt: new Date().toISOString()
    }));

    res.json({ mensagem: "Planilha atualizada com sucesso." });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao fazer upload." });
  }
});

app.get("/api/planilha", requireAuth, (_req, res) => {
  const info = getWorkbookInfo();
  res.json(info);
});

// ============= LOGOUT =============

app.get("/logout", (req, res) => {
  req.session.destroy(() => {
    res.redirect("/");
  });
});

// ============= START =============

app.listen(PORT, "0.0.0.0", () => {
  console.log(`🚀 Servidor rodando em http://localhost:${PORT}`);
  console.log(`✅ Conectado ao Supabase: ${SUPABASE_URL}`);
});
