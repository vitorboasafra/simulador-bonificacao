const express = require("express");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const session = require("express-session");
const multer = require("multer");
const nodemailer = require("nodemailer");
const CFB = require("cfb");
const XLSX = require("xlsx");
const PDFKitDocument = require("pdfkit");
const { PDFDocument: PDFLibDocument } = require("pdf-lib");
const { createClient } = require("@supabase/supabase-js");
require("dotenv").config();

const app = express();
const PORT = process.env.PORT || 3000;

// ============= SUPABASE SETUP =============
const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_ANON_KEY = process.env.SUPABASE_ANON_KEY;
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;

if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
  console.error("❌ SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY são obrigatórios!");
  process.exit(1);
}

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY || SUPABASE_SERVICE_ROLE_KEY, {
  auth: { persistSession: false }
});
const supabaseAdmin = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, {
  auth: { persistSession: false }
});

// ============= PATHS =============
const ROOT_DIR = __dirname;
const PUBLIC_DIR = path.join(ROOT_DIR, "public");
const DATA_DIR = path.join(ROOT_DIR, "data");
const WORKBOOK_SNAPSHOT_CONFIG_KEY = "planilha_atual_snapshot";
const PRACA_PREP_META = path.join(DATA_DIR, "pracas-preparadas.json");
const UNITS_WORKBOOK = path.join(ROOT_DIR, "Unidades Boa Safra.xlsx");
const MODEL_PDF = path.join(ROOT_DIR, "NF MODELO BONIFICACAO.pdf");
const LOGO_IMAGE = path.join(ROOT_DIR, "logo.png");
const PUBLIC_LOGO_IMAGE = path.join(PUBLIC_DIR, "assets", "boa-safra-logo.png");
const LOGIN_BACKGROUND = path.join(ROOT_DIR, "plano de fundo.png");
const LETTERHEAD_IMAGE = path.join(ROOT_DIR, "Boa Safra - Papel Timbrado.png");
const REPORT_CONTENT_TOP = 145;
const REPORT_BOTTOM_MARGIN = 125;

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
const APP_URL = process.env.APP_URL || "";
const DEFAULT_PUBLIC_APP_URL = "https://simulador-bonificacao-production.up.railway.app";

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
app.get("/logo.png", (_req, res) => {
  const logoPath = fs.existsSync(LOGO_IMAGE) ? LOGO_IMAGE : PUBLIC_LOGO_IMAGE;
  if (!fs.existsSync(logoPath)) {
    return res.status(404).send("Logo não encontrada.");
  }
  return res.sendFile(logoPath);
});
app.get("/plano-de-fundo.png", (_req, res) => {
  if (!fs.existsSync(LOGIN_BACKGROUND)) {
    return res.status(404).send("Plano de fundo não encontrado.");
  }
  return res.sendFile(LOGIN_BACKGROUND);
});
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

app.get(["/cadastro-confirmado", "/auth/callback", "/auth/confirmado"], (req, res) => {
  const hasError = req.query.error || req.query.error_code || req.query.error_description;
  res.redirect(hasError ? "/?confirmado=erro" : "/?cadastro=confirmado");
});

// ============= HELPERS =============
function formatNumberPtBr(n) {
  const num = Number(n) || 0;
  return num.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function isEmailValid(email) {
  return /^[^\s@]+@boasafrasementes\.com\.br$/i.test(String(email || "").trim());
}

function normalizeEmail(email) {
  return String(email || "").trim().toLowerCase();
}

function sanitizeUsuario(usuario) {
  return String(usuario || "").trim();
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
  const { data, error } = await supabaseAdmin
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

function normalizeBaseUrl(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const withProtocol = /^https?:\/\//i.test(raw) ? raw : `https://${raw}`;
  return withProtocol.replace(/\/+$/, "");
}

function getConfiguredAppUrl() {
  return normalizeBaseUrl(
    APP_URL ||
    process.env.RAILWAY_PUBLIC_DOMAIN ||
    process.env.RAILWAY_STATIC_URL ||
    process.env.RAILWAY_SERVICE_SIMULADOR_BONIFICACAO_URL ||
    DEFAULT_PUBLIC_APP_URL
  );
}

function getAppBaseUrl(req) {
  const configuredUrl = getConfiguredAppUrl();
  if (configuredUrl) return configuredUrl;
  const protocol = req.get("x-forwarded-proto") || req.protocol || "http";
  return `${protocol}://${req.get("host")}`;
}

function getEmailRedirectTo(req) {
  return `${getAppBaseUrl(req)}/cadastro-confirmado`;
}

async function createSupabaseAuthUser({ email, senha, usuario }) {
  const { data, error } = await supabaseAdmin.auth.admin.createUser({
    email,
    password: senha,
    email_confirm: true,
    user_metadata: { usuario }
  });

  return { user: data?.user, error };
}

async function signUpSupabaseAuthUser({ email, senha, usuario }, req) {
  const { data, error } = await supabase.auth.signUp({
    email,
    password: senha,
    options: {
      data: { usuario },
      emailRedirectTo: getEmailRedirectTo(req)
    }
  });

  return { user: data?.user, session: data?.session, error };
}

async function resendSignupConfirmationEmail(email, req) {
  const { data, error } = await supabase.auth.resend({
    type: "signup",
    email,
    options: {
      emailRedirectTo: getEmailRedirectTo(req)
    }
  });

  return { data, error };
}

async function signInSupabaseAuth({ email, senha }) {
  const { data, error } = await supabase.auth.signInWithPassword({
    email,
    password: senha
  });

  return { user: data?.user, session: data?.session, error };
}

async function getAuthUserByEmail(email) {
  const normalizedEmail = normalizeEmail(email);
  if (!normalizedEmail) return null;

  const { data, error } = await supabaseAdmin.auth.admin.listUsers({ page: 1, perPage: 1000 });
  if (error) return null;

  return (data?.users || []).find(user => normalizeEmail(user.email) === normalizedEmail) || null;
}

async function getAuthUserForProfile(profile) {
  if (!profile) return null;

  if (profile.id) {
    const { data } = await supabaseAdmin.auth.admin.getUserById(profile.id);
    if (data?.user) return data.user;
  }

  return getAuthUserByEmail(profile.email);
}

async function findOrCreateUserProfile(email, usuario, senha = "", authUserId = null, options = {}) {
  const normalizedEmail = normalizeEmail(email);
  console.log(`🔍 Procurando perfil para: ${normalizedEmail}`);

  if (!normalizedEmail) {
    console.log("❌ Email normalizado vazio");
    return null;
  }

  const { data: existingByEmail, error: emailError } = await supabaseAdmin
    .from("usuarios")
    .select("*")
    .eq("email", normalizedEmail)
    .maybeSingle();

  if (emailError) {
    console.log(`❌ Erro ao buscar por email: ${emailError.message}`);
    return null;
  }

  if (existingByEmail) {
    console.log(`✅ Perfil encontrado: ${existingByEmail.usuario}`);
    return existingByEmail;
  }

  console.log(`📝 Perfil não encontrado, criando novo...`);

  let candidate = sanitizeUsuario(usuario) || normalizedEmail.split("@")[0];
  if (!candidate) candidate = normalizedEmail;

  let uniqueCandidate = candidate;
  let counter = 1;
  while (true) {
    const { data: existingByUsuario, error: usuarioError } = await supabaseAdmin
      .from("usuarios")
      .select("id")
      .eq("usuario", uniqueCandidate)
      .maybeSingle();

    if (usuarioError) {
      console.log(`❌ Erro ao verificar usuario único: ${usuarioError.message}`);
      return null;
    }

    if (!existingByUsuario) break;
    uniqueCandidate = `${candidate}${counter++}`;
  }

  console.log(`📝 Inserindo novo perfil: usuario=${uniqueCandidate}, email=${normalizedEmail}`);

  const profile = {
    usuario: uniqueCandidate,
    email: normalizedEmail,
    senha,
    confirmado: options.confirmado ?? true,
    admin: normalizeBoolean(options.admin)
  };

  if (authUserId) {
    profile.id = authUserId;
  }

  const { data, error } = await supabaseAdmin.from("usuarios").insert(profile).select("*").maybeSingle();

  if (error) {
    console.log(`❌ Erro ao inserir perfil: ${error.message}`);
    return null;
  }

  console.log(`✅ Perfil criado: ${data.usuario}`);
  return data;
}

function isSupabaseDuplicateAuthError(error) {
  const message = String(error?.message || "").toLowerCase();
  return message.includes("already registered") || message.includes("duplicate") || message.includes("already exists") || message.includes("already been taken");
}

function isEmailNotConfirmedError(error) {
  const message = String(error?.message || "").toLowerCase();
  const code = String(error?.code || "").toLowerCase();
  return code.includes("email_not_confirmed") || message.includes("email not confirmed") || message.includes("not confirmed");
}

function isConfirmationEmailSendError(error) {
  const message = String(error?.message || "").toLowerCase();
  const code = String(error?.code || "").toLowerCase();
  return code.includes("unexpected_failure") && message.includes("email") ||
    message.includes("error sending confirmation email") ||
    message.includes("failed to send") && message.includes("email");
}

function isAuthUserConfirmed(user) {
  return Boolean(user?.email_confirmed_at || user?.confirmed_at);
}

async function syncProfileConfirmation(profile, authUser) {
  if (!profile || !isAuthUserConfirmed(authUser) || profile.confirmado) return profile;

  const { data, error } = await supabaseAdmin
    .from("usuarios")
    .update({ confirmado: true })
    .eq("email", profile.email)
    .select("*")
    .maybeSingle();

  if (error) throw error;
  return data || { ...profile, confirmado: true };
}

async function signInOrMigrateLegacyUser(email, senha) {
  const { user, error } = await signInSupabaseAuth({ email, senha });
  if (user) {
    return { user, error: null };
  }

  if (isEmailNotConfirmedError(error)) {
    return { user: null, error, emailNotConfirmed: true };
  }

  const { data: legacyUser, error: legacyError } = await supabaseAdmin
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

  if (legacyUser.confirmado === false) {
    return { user: null, error: error || new Error("Email not confirmed"), emailNotConfirmed: true };
  }

  const { user: authUser, error: createError } = await createSupabaseAuthUser({
    email,
    senha,
    usuario: legacyUser.usuario
  });

  if (createError && !isSupabaseDuplicateAuthError(createError)) {
    return { user: null, error: createError };
  }

  const existingAuthUser = authUser || await getAuthUserByEmail(email);

  return {
    user: existingAuthUser || { email, user_metadata: { usuario: legacyUser.usuario } },
    error: null
  };
}

// ============= DATABASE FUNCTIONS =============

// Planilha atual
let workbookRowsCache = null;
let pracaRowsCache = null;
let classificationRowsCache = null;
let pracaConfigCache = null;
let valorSacaCache = null;
let pracaWorkbookPrepareTimer = null;
let pracaWorkbookWatcher = null;
let zmm113WatcherTimer = null;
let zmm113Watcher = null;
let extraWorkbookWatcher = null;
let extraWorkbookSyncTimers = {};

function readWorkbookMeta() {
  const WORKBOOK_META = path.join(DATA_DIR, "planilha.json");
  if (!fs.existsSync(WORKBOOK_META)) return null;
  try {
    return JSON.parse(fs.readFileSync(WORKBOOK_META, "utf8"));
  } catch (_err) {
    return null;
  }
}

function writeWorkbookMeta(meta) {
  fs.writeFileSync(path.join(DATA_DIR, "planilha.json"), JSON.stringify(meta, null, 2));
}

function getCurrentWorkbookNames() {
  return ["planilha-atual.xlsx", "planilha-atual.xls"];
}

function buildUploadedWorkbookCandidate(file, meta) {
  const absolute = path.join(DATA_DIR, file);
  if (!fs.existsSync(absolute)) return null;

  const stat = fs.statSync(absolute);
  const timestampMatch = file.match(/^planilha-(\d+)\.(xlsx|xls)$/i);
  const isMetaFile = meta?.storedName === file || meta?.currentName === file;
  const metaUploadedAtMs = isMetaFile && meta?.uploadedAt ? Date.parse(meta.uploadedAt) : NaN;
  const filenameUploadedAtMs = timestampMatch ? Number(timestampMatch[1]) : NaN;
  const uploadedAtMs = Number.isFinite(metaUploadedAtMs)
    ? metaUploadedAtMs
    : Number.isFinite(filenameUploadedAtMs)
      ? filenameUploadedAtMs
      : stat.mtimeMs;

  return {
    absolute,
    storedName: file,
    originalName: isMetaFile ? meta?.originalName : "",
    uploadedAt: isMetaFile ? meta?.uploadedAt : new Date(uploadedAtMs || stat.mtimeMs).toISOString(),
    uploadedAtMs: uploadedAtMs || stat.mtimeMs,
    uploadedBy: isMetaFile ? meta?.uploadedBy : null,
    uploadedByEmail: isMetaFile ? meta?.uploadedByEmail : null,
    mtimeMs: stat.mtimeMs
  };
}

function removeOldUploadedWorkbookFiles(keepName) {
  if (!fs.existsSync(DATA_DIR)) return;

  for (const file of fs.readdirSync(DATA_DIR)) {
    if (file === keepName) continue;
    if (!/^planilha-(?:atual|\d+)\.(xlsx|xls)$/i.test(file)) continue;

    try {
      fs.unlinkSync(path.join(DATA_DIR, file));
    } catch (error) {
      console.warn(`Não foi possível remover planilha antiga ${file}: ${error.message}`);
    }
  }
}

function saveCurrentUploadedWorkbook(buffer, originalName) {
  const originalExt = path.extname(originalName).toLowerCase();
  const ext = originalExt === ".xls" ? ".xls" : ".xlsx";
  const currentName = `planilha-atual${ext}`;
  const currentPath = path.join(DATA_DIR, currentName);

  fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(currentPath, buffer);
  removeOldUploadedWorkbookFiles(currentName);

  return {
    currentName,
    currentPath,
    ext,
    size: buffer.length,
    sha256: crypto.createHash("sha256").update(buffer).digest("hex")
  };
}

async function saveWorkbookSnapshot(meta, buffer) {
  await setConfigValue(WORKBOOK_SNAPSHOT_CONFIG_KEY, JSON.stringify({
    ...meta,
    contentBase64: buffer.toString("base64")
  }));
}

async function restoreWorkbookSnapshot() {
  let raw = "";
  try {
    raw = await getConfigValue(WORKBOOK_SNAPSHOT_CONFIG_KEY);
  } catch (error) {
    console.warn(`Não foi possível buscar a planilha persistida: ${error.message}`);
    return false;
  }

  if (!raw) return false;

  let snapshot = null;
  try {
    snapshot = JSON.parse(raw);
  } catch (error) {
    console.warn(`Metadados da planilha persistida inválidos: ${error.message}`);
    return false;
  }

  if (!snapshot?.contentBase64) return false;

  const localMeta = readWorkbookMeta();
  const localName = localMeta?.currentName || localMeta?.storedName;
  const localPath = localName ? path.join(DATA_DIR, localName) : "";
  const localExists = localPath && fs.existsSync(localPath);
  const localUploadedAtMs = localMeta?.uploadedAt ? Date.parse(localMeta.uploadedAt) : NaN;
  const snapshotUploadedAtMs = snapshot.uploadedAt ? Date.parse(snapshot.uploadedAt) : NaN;

  if (
    localExists &&
    Number.isFinite(localUploadedAtMs) &&
    Number.isFinite(snapshotUploadedAtMs) &&
    localUploadedAtMs >= snapshotUploadedAtMs
  ) {
    return false;
  }

  const buffer = Buffer.from(snapshot.contentBase64, "base64");
  const ext = snapshot.ext === ".xls" ? ".xls" : ".xlsx";
  const currentName = `planilha-atual${ext}`;
  const currentPath = path.join(DATA_DIR, currentName);
  const { contentBase64: _contentBase64, ...meta } = snapshot;

  fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(currentPath, buffer);
  removeOldUploadedWorkbookFiles(currentName);
  writeWorkbookMeta({
    ...meta,
    currentName,
    storedName: currentName,
    ext,
    size: buffer.length,
    sha256: crypto.createHash("sha256").update(buffer).digest("hex")
  });
  clearWorkbookCache();
  return true;
}

function findWorkbookPath() {
  const candidates = ["zmm113.xlsx", "zmm113.XLSX", "zmm113.xls", "zmm113.XLS"];
  for (const file of candidates) {
    const absolute = path.join(ROOT_DIR, file);
    if (fs.existsSync(absolute)) return absolute;
  }
  return null;
}

function findPracaWorkbookPath() {
  const candidates = [
    "pracas.xlsx",
    "pracas.xls",
    "praças.xlsx",
    "praças.xls",
    "praca.xlsx",
    "praca.xls",
    "praça.xlsx",
    "praça.xls"
  ];
  for (const file of candidates) {
    const absolute = path.join(ROOT_DIR, file);
    if (fs.existsSync(absolute)) return absolute;
  }
  return null;
}

function readPracaPrepMeta() {
  if (!fs.existsSync(PRACA_PREP_META)) return null;
  try {
    return JSON.parse(fs.readFileSync(PRACA_PREP_META, "utf8"));
  } catch (_err) {
    return null;
  }
}

function writePracaPrepMeta(meta) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(PRACA_PREP_META, JSON.stringify(meta, null, 2));
}

function decodeXmlText(value) {
  return String(value || "")
    .replace(/&#x([0-9a-f]+);/gi, (_match, hex) => String.fromCharCode(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_match, decimal) => String.fromCharCode(parseInt(decimal, 10)))
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&gt;/g, ">")
    .replace(/&lt;/g, "<")
    .replace(/&amp;/g, "&");
}

function escapeXmlText(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function getCfbEntryText(cfb, entryName) {
  const entry = CFB.find(cfb, `/${entryName}`);
  if (!entry?.content) return "";
  return Buffer.from(entry.content).toString("utf8");
}

function updateCfbEntryText(cfb, entryName, transform) {
  const entry = CFB.find(cfb, `/${entryName}`);
  if (!entry?.content) return { changed: false };

  const text = Buffer.from(entry.content).toString("utf8");
  const newText = transform(text);
  if (newText === text) return { changed: false };

  const content = Buffer.from(newText, "utf8");
  entry.content = content;
  entry.size = content.length;
  return { changed: true };
}

function convertToIntegerText(value) {
  const clean = decodeXmlText(value).trim().replace(",", ".");
  if (!clean) return "";

  const number = Number(clean);
  if (!Number.isFinite(number)) return "";
  return number.toFixed(0);
}

function convertInscricaoColumnXml(text, column) {
  let convertedCells = 0;
  const pattern = new RegExp(`<c\\b(?=[^>]*\\br="${column}(\\d+)")[^>]*>.*?<\\/c>`, "gs");
  const xml = text.replace(pattern, (cell, rowText) => {
    const row = Number(rowText);
    if (row <= 1 || /\bt="s"/.test(cell)) return cell;

    const rawMatch = cell.match(/<v>(.*?)<\/v>/s) || cell.match(/<t[^>]*>(.*?)<\/t>/s);
    if (!rawMatch) return cell;

    const decoded = decodeXmlText(rawMatch[1]).trim();
    if (!/^-?\d+(?:[\.,]\d+)?(?:E[+-]?\d+)?$/i.test(decoded)) return cell;

    const value = convertToIntegerText(decoded);
    if (!value) return cell;

    const styleMatch = cell.match(/\bs="([^"]+)"/);
    const styleAttr = styleMatch ? ` s="${styleMatch[1]}"` : "";
    const newCell = `<c r="${column}${row}"${styleAttr} t="inlineStr"><is><t>${escapeXmlText(value)}</t></is></c>`;
    if (newCell !== cell) convertedCells += 1;
    return newCell;
  });

  return { xml, convertedCells };
}

function getSheetTargetsByName(cfb) {
  const workbookText = getCfbEntryText(cfb, "xl/workbook.xml");
  const relsText = getCfbEntryText(cfb, "xl/_rels/workbook.xml.rels");
  const relTargets = new Map();

  for (const match of relsText.matchAll(/<Relationship\b[^>]*\/?>/g)) {
    const tag = match[0];
    const id = tag.match(/\bId="([^"]+)"/)?.[1];
    let target = tag.match(/\bTarget="([^"]+)"/)?.[1];
    if (!id || !target) continue;

    target = target.trim().replace(/^\/+/, "");
    if (!target.startsWith("xl/")) target = `xl/${target}`;
    relTargets.set(id, target);
  }

  const sheets = new Map();
  for (const match of workbookText.matchAll(/<sheet\b[^>]*\/?>/g)) {
    const tag = match[0];
    const name = tag.match(/\bname="([^"]+)"/)?.[1];
    const id = tag.match(/\br:id="([^"]+)"/)?.[1];
    const target = id ? relTargets.get(id) : "";
    if (name && target) sheets.set(decodeXmlText(name), target);
  }

  return sheets;
}

function preparePracaWorkbookZip(workbookPath) {
  const cfb = CFB.read(workbookPath, { type: "file" });
  const changedEntries = new Set();
  let convertedCells = 0;

  if (updateCfbEntryText(cfb, "xl/workbook.xml", (text) =>
    text
      .replace(/<workbookProtection\b[^>]*\/>/g, "")
      .replace(/<workbookProtection\b[^>]*>.*?<\/workbookProtection>/gs, "")
      .replace(/<fileSharing\b[^>]*\/>/g, "")
      .replace(/<fileSharing\b[^>]*>.*?<\/fileSharing>/gs, "")
  ).changed) {
    changedEntries.add("xl/workbook.xml");
  }

  const worksheetEntries = cfb.FullPaths
    .map(fullPath => fullPath.replace(/^Root Entry\//, ""))
    .filter(entryName => /^xl\/worksheets\/sheet\d+\.xml$/i.test(entryName));

  for (const entryName of worksheetEntries) {
    if (updateCfbEntryText(cfb, entryName, (text) =>
      text
        .replace(/<sheetProtection\b[^>]*\/>/g, "")
        .replace(/<sheetProtection\b[^>]*>.*?<\/sheetProtection>/gs, "")
        .replace(/<protectedRanges\b[^>]*>.*?<\/protectedRanges>/gs, "")
        .replace(/<protectedRange\b[^>]*\/>/g, "")
    ).changed) {
      changedEntries.add(entryName);
    }
  }

  if (updateCfbEntryText(cfb, "xl/styles.xml", (text) =>
    text
      .replace(/<protection\b[^>]*\/>/g, '<protection locked="0" hidden="0"/>')
      .replace(/<protection\b[^>]*>.*?<\/protection>/gs, '<protection locked="0" hidden="0"/>')
  ).changed) {
    changedEntries.add("xl/styles.xml");
  }

  const sheetTargets = getSheetTargetsByName(cfb);
  const sheetColumns = {
    ZMM113: "C",
    DIN: "D",
    Compilado: "E"
  };

  for (const [sheetName, column] of Object.entries(sheetColumns)) {
    const entryName = sheetTargets.get(sheetName);
    if (!entryName) continue;

    const result = updateCfbEntryText(cfb, entryName, (text) => {
      const converted = convertInscricaoColumnXml(text, column);
      convertedCells += converted.convertedCells;
      return converted.xml;
    });
    if (result.changed) changedEntries.add(entryName);
  }

  if (changedEntries.size > 0) {
    CFB.writeFile(cfb, workbookPath, { fileType: "zip", compression: true });
  }

  return {
    changed: changedEntries.size > 0,
    convertedCells,
    changedEntries: [...changedEntries]
  };
}

function getTimestampForFileName(date = new Date()) {
  const pad = value => String(value).padStart(2, "0");
  return [
    date.getFullYear(),
    pad(date.getMonth() + 1),
    pad(date.getDate()),
    "-",
    pad(date.getHours()),
    pad(date.getMinutes()),
    pad(date.getSeconds())
  ].join("");
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function getPracaWorkbookBackupFiles(workbookPath) {
  const directory = path.dirname(workbookPath);
  const baseName = path.basename(workbookPath, path.extname(workbookPath));
  const extension = path.extname(workbookPath);
  const backupPattern = new RegExp(`^${escapeRegExp(baseName)}\\.backup-auto-\\d{8}-\\d{6}${escapeRegExp(extension)}$`, "i");

  return fs.readdirSync(directory)
    .filter(file => backupPattern.test(file))
    .map(file => {
      const absolute = path.join(directory, file);
      const stat = fs.statSync(absolute);
      return { absolute, file, mtimeMs: stat.mtimeMs };
    })
    .sort((a, b) => b.file.localeCompare(a.file) || b.mtimeMs - a.mtimeMs);
}

function pruneOldPracaWorkbookBackups(workbookPath, keepPath) {
  const keepAbsolute = keepPath ? path.resolve(keepPath) : null;
  const backups = getPracaWorkbookBackupFiles(workbookPath);
  const latest = keepAbsolute || backups[0]?.absolute || null;
  if (!latest) return;

  for (const backup of backups) {
    if (path.resolve(backup.absolute) === path.resolve(latest)) continue;

    try {
      fs.unlinkSync(backup.absolute);
    } catch (error) {
      console.warn(`NÃ£o foi possÃ­vel remover backup antigo de praÃ§as ${backup.file}: ${error.message}`);
    }
  }
}

function createPracaWorkbookBackup(workbookPath) {
  const directory = path.dirname(workbookPath);
  const baseName = path.basename(workbookPath, path.extname(workbookPath));
  const backupPath = path.join(directory, `${baseName}.backup-auto-${getTimestampForFileName()}.xlsx`);
  fs.copyFileSync(workbookPath, backupPath);
  pruneOldPracaWorkbookBackups(workbookPath, backupPath);
  return backupPath;
}

function getFileSha256(filePath) {
  return crypto.createHash("sha256").update(fs.readFileSync(filePath)).digest("hex");
}

function getPracaWorkbookFingerprint(workbookPath) {
  const absolute = path.resolve(workbookPath);
  const stat = fs.statSync(absolute);
  return {
    absolute,
    size: stat.size,
    mtimeMs: stat.mtimeMs,
    sha256: getFileSha256(absolute)
  };
}

function wasPracaWorkbookPrepared(fingerprint) {
  const meta = readPracaPrepMeta();
  return (
    meta?.absolute === fingerprint.absolute &&
    meta?.sha256 === fingerprint.sha256 &&
    meta?.size === fingerprint.size
  );
}

function preparePracaWorkbookIfNeeded(workbookPath) {
  if (!workbookPath) return;
  const absolute = path.resolve(workbookPath);

  if (path.extname(absolute).toLowerCase() !== ".xlsx") return;

  let before;
  try {
    before = getPracaWorkbookFingerprint(absolute);
  } catch (error) {
    console.warn(`Não foi possível verificar a planilha de praças: ${error.message}`);
    return;
  }

  if (wasPracaWorkbookPrepared(before)) return;

  try {
    const backupPath = createPracaWorkbookBackup(absolute);
    const prepared = preparePracaWorkbookZip(absolute);
    const after = getPracaWorkbookFingerprint(absolute);
    writePracaPrepMeta({
      ...after,
      originalSha256: before.sha256,
      preparedAt: new Date().toISOString(),
      backup: path.relative(ROOT_DIR, backupPath),
      convertedCells: prepared.convertedCells,
      changedEntries: prepared.changedEntries
    });
    pracaRowsCache = null;
    pracaConfigCache = null;
    console.log(`Planilha de praças preparada automaticamente: ${path.basename(absolute)}`);
  } catch (error) {
    console.warn(`Não foi possível preparar a planilha de praças automaticamente: ${error.message}`);
  }
}

function schedulePracaWorkbookPreparation(delayMs = 1500) {
  if (pracaWorkbookPrepareTimer) clearTimeout(pracaWorkbookPrepareTimer);
  pracaWorkbookPrepareTimer = setTimeout(() => {
    pracaWorkbookPrepareTimer = null;
    try {
      preparePracaWorkbookIfNeeded(findPracaWorkbookPath());
    } catch (error) {
      console.warn(`Erro na preparação automática da planilha de praças: ${error.message}`);
    }
  }, delayMs);
}

function watchPracaWorkbookChanges() {
  if (pracaWorkbookWatcher) return;

  try {
    pracaWorkbookWatcher = fs.watch(ROOT_DIR, (_eventType, filename) => {
      const name = String(filename || "").toLowerCase();
      if (name.includes(".backup-auto-")) return;
      if (name && !name.includes("praca") && !name.includes("pra")) return;
      schedulePracaWorkbookPreparation();
      scheduleExtraWorkbookSync("pracas_snapshot", 3000);
    });
  } catch (error) {
    console.warn(`Não foi possível monitorar a planilha de praças: ${error.message}`);
  }
}

async function syncZmm113ToActive() {
  const zmm113Path = findWorkbookPath();
  if (!zmm113Path) return false;

  try {
    const buffer = fs.readFileSync(zmm113Path);
    const originalName = path.basename(zmm113Path);
    const savedWorkbook = saveCurrentUploadedWorkbook(buffer, originalName);
    const meta = {
      storedName: savedWorkbook.currentName,
      currentName: savedWorkbook.currentName,
      ext: savedWorkbook.ext,
      size: savedWorkbook.size,
      sha256: savedWorkbook.sha256,
      originalName,
      uploadedAt: new Date().toISOString(),
      uploadedBy: "auto-sync",
      uploadedByEmail: null
    };
    writeWorkbookMeta(meta);
    await saveWorkbookSnapshot(meta, buffer);
    clearWorkbookCache();
    console.log(`✅ Planilha sincronizada automaticamente: ${originalName}`);
    return true;
  } catch (error) {
    console.warn(`Não foi possível sincronizar a planilha principal: ${error.message}`);
    return false;
  }
}

function scheduleZmm113Sync(delayMs = 2000) {
  if (zmm113WatcherTimer) clearTimeout(zmm113WatcherTimer);
  zmm113WatcherTimer = setTimeout(async () => {
    zmm113WatcherTimer = null;
    try {
      await syncZmm113ToActive();
    } catch (error) {
      console.warn(`Erro na sincronização automática da planilha: ${error.message}`);
    }
  }, delayMs);
}

function watchZmm113Changes() {
  if (zmm113Watcher) return;

  try {
    zmm113Watcher = fs.watch(ROOT_DIR, (_eventType, filename) => {
      const name = String(filename || "").toLowerCase();
      if (!name || name.includes(".backup-auto-") || name.startsWith("planilha")) return;
      if (!name.includes("zmm")) return;
      scheduleZmm113Sync();
    });
  } catch (error) {
    console.warn(`Não foi possível monitorar a planilha principal: ${error.message}`);
  }
}

function startSnapshotSyncPolling() {
  const INTERVAL_MS = 5 * 60 * 1000;

  setInterval(async () => {
    try {
      const restored = await restoreWorkbookSnapshot();
      if (restored) {
        console.log("✅ Planilha atualizada automaticamente do armazenamento persistente.");
      }
    } catch (error) {
      console.warn(`Não foi possível sincronizar planilha do Supabase: ${error.message}`);
    }

    for (const config of getExtraWorkbookConfigs()) {
      try {
        await restoreExtraWorkbookSnapshot(config);
      } catch (error) {
        console.warn(`Não foi possível sincronizar ${config.label}: ${error.message}`);
      }
    }
  }, INTERVAL_MS);
}

function getExtraWorkbookConfigs() {
  return [
    {
      key: "pracas_snapshot",
      label: "Praças",
      fixedRestoreName: "pracas.xlsx",
      getPath: findPracaWorkbookPath,
      clearCache() { pracaRowsCache = null; pracaConfigCache = null; },
      afterRestore(p) { try { preparePracaWorkbookIfNeeded(p); } catch (_e) {} }
    },
    {
      key: "classificacoes_snapshot",
      label: "Classificações",
      fixedRestoreName: "classificacoes.xlsx",
      getPath: findClassificationWorkbookPath,
      clearCache() { classificationRowsCache = null; },
      afterRestore: null
    },
    {
      key: "unidades_snapshot",
      label: "Unidades Boa Safra",
      fixedRestoreName: path.basename(UNITS_WORKBOOK),
      getPath() { return fs.existsSync(UNITS_WORKBOOK) ? UNITS_WORKBOOK : null; },
      clearCache() { unitsRowsCache = null; },
      afterRestore: null
    },
    {
      key: "valor_saca_snapshot",
      label: "Valor Saca",
      fixedRestoreName: "valor saca.xlsx",
      getPath: findValorSacaWorkbookPath,
      clearCache() { valorSacaCache = null; },
      afterRestore: null
    }
  ];
}

async function saveExtraWorkbookSnapshot(config) {
  const filePath = config.getPath();
  if (!filePath || !fs.existsSync(filePath)) return false;

  try {
    const stat = fs.statSync(filePath);
    const buffer = fs.readFileSync(filePath);
    const snapshot = {
      originalName: path.basename(filePath),
      fileModifiedAt: new Date(stat.mtimeMs).toISOString(),
      savedAt: new Date().toISOString(),
      contentBase64: buffer.toString("base64")
    };
    await setConfigValue(config.key, JSON.stringify(snapshot));
    console.log(`✅ ${config.label} sincronizada no Supabase: ${snapshot.originalName}`);
    return true;
  } catch (error) {
    console.warn(`Não foi possível salvar ${config.label} no Supabase: ${error.message}`);
    return false;
  }
}

async function restoreExtraWorkbookSnapshot(config) {
  let raw = "";
  try {
    raw = await getConfigValue(config.key);
  } catch (_error) {
    return false;
  }

  if (!raw) return false;

  let snapshot;
  try {
    snapshot = JSON.parse(raw);
  } catch (_err) {
    return false;
  }

  if (!snapshot?.contentBase64) return false;

  const restorePath = path.join(ROOT_DIR, config.fixedRestoreName);
  if (fs.existsSync(restorePath)) {
    try {
      const localStat = fs.statSync(restorePath);
      const snapshotFileModifiedAtMs = snapshot.fileModifiedAt ? Date.parse(snapshot.fileModifiedAt) : NaN;
      if (Number.isFinite(snapshotFileModifiedAtMs) && localStat.mtimeMs >= snapshotFileModifiedAtMs) {
        return false;
      }
    } catch (_err) {
      // proceed to restore
    }
  }

  try {
    const buffer = Buffer.from(snapshot.contentBase64, "base64");
    fs.writeFileSync(restorePath, buffer);
    config.clearCache();
    if (config.afterRestore) config.afterRestore(restorePath);
    console.log(`✅ ${config.label} restaurada do Supabase: ${snapshot.originalName || config.fixedRestoreName}`);
    return true;
  } catch (error) {
    console.warn(`Não foi possível restaurar ${config.label}: ${error.message}`);
    return false;
  }
}

async function restoreAllExtraWorkbooks() {
  for (const config of getExtraWorkbookConfigs()) {
    try {
      await restoreExtraWorkbookSnapshot(config);
    } catch (error) {
      console.warn(`Erro ao restaurar ${config.label}: ${error.message}`);
    }
  }
}

function scheduleExtraWorkbookSync(configKey, delayMs = 2000) {
  if (extraWorkbookSyncTimers[configKey]) clearTimeout(extraWorkbookSyncTimers[configKey]);
  extraWorkbookSyncTimers[configKey] = setTimeout(async () => {
    delete extraWorkbookSyncTimers[configKey];
    const config = getExtraWorkbookConfigs().find(c => c.key === configKey);
    if (!config) return;
    try {
      await saveExtraWorkbookSnapshot(config);
    } catch (error) {
      console.warn(`Erro ao sincronizar ${config.label}: ${error.message}`);
    }
  }, delayMs);
}

function watchExtraWorkbookChanges() {
  if (extraWorkbookWatcher) return;

  const PATTERNS = [
    { test: (n) => /classifica/i.test(n), key: "classificacoes_snapshot" },
    { test: (n) => /unidade/i.test(n), key: "unidades_snapshot" },
    { test: (n) => /valor.{0,5}saca/i.test(n), key: "valor_saca_snapshot" }
  ];

  try {
    extraWorkbookWatcher = fs.watch(ROOT_DIR, (_eventType, filename) => {
      if (!filename) return;
      const name = String(filename).toLowerCase();
      if (name.includes(".backup-auto-") || name.startsWith("planilha") || name.includes("zmm")) return;
      if (name.includes("praca") || name.includes("praç")) return;

      for (const p of PATTERNS) {
        if (p.test(filename)) {
          scheduleExtraWorkbookSync(p.key);
          break;
        }
      }
    });
  } catch (error) {
    console.warn(`Não foi possível monitorar planilhas auxiliares: ${error.message}`);
  }
}

function findValorSacaWorkbookPath() {
  const candidates = [
    "valor saca.xlsx",
    "valor saca.xls",
    "valor-saca.xlsx",
    "valor-saca.xls",
    "valor_saca.xlsx",
    "valor_saca.xls",
    "valorsaca.xlsx",
    "valorsaca.xls"
  ];
  for (const file of candidates) {
    const absolute = path.join(ROOT_DIR, file);
    if (fs.existsSync(absolute)) return absolute;
  }
  return null;
}

function findClassificationWorkbookPath() {
  const extensions = new Set([".csv", ".xlsx", ".xls"]);
  const candidates = fs.readdirSync(ROOT_DIR)
    .filter(file => /classifica/i.test(file) && extensions.has(path.extname(file).toLowerCase()))
    .map(file => {
      const absolute = path.join(ROOT_DIR, file);
      const stat = fs.statSync(absolute);
      return { absolute, mtimeMs: stat.mtimeMs };
    })
    .sort((a, b) => b.mtimeMs - a.mtimeMs);

  return candidates[0]?.absolute || null;
}

function findUploadedWorkbookPath() {
  return getLatestUploadedWorkbook()?.absolute || null;
}

function getUploadedWorkbookCandidates() {
  const meta = readWorkbookMeta();
  if (!fs.existsSync(DATA_DIR)) return [];

  return fs.readdirSync(DATA_DIR)
    .filter(file => /^planilha-(\d+)\.(xlsx|xls)$/i.test(file) || /^planilha-atual\.(xlsx|xls)$/i.test(file))
    .map(file => buildUploadedWorkbookCandidate(file, meta))
    .filter(Boolean)
    .sort((a, b) => (b.uploadedAtMs - a.uploadedAtMs) || (b.mtimeMs - a.mtimeMs));
}

function getLatestUploadedWorkbook() {
  const meta = readWorkbookMeta();
  const candidates = getUploadedWorkbookCandidates();
  const latest = candidates[0] || null;
  const currentNames = [
    meta?.currentName,
    ...getCurrentWorkbookNames()
  ].filter(Boolean);

  for (const file of [...new Set(currentNames)]) {
    const candidate = buildUploadedWorkbookCandidate(file, meta);
    if (candidate && (!latest || candidate.uploadedAtMs >= latest.uploadedAtMs)) {
      return candidate;
    }
  }

  if (!latest) return null;
  if (meta?.storedName === latest.storedName || meta?.currentName === latest.storedName) {
    return latest;
  }

  return {
    ...latest,
    originalName: latest.originalName || latest.storedName,
    uploadedBy: latest.uploadedBy || null,
    uploadedByEmail: latest.uploadedByEmail || null
  };
}

function getWorkbookInfo() {
  const uploaded = getLatestUploadedWorkbook();

  if (uploaded) {
    return {
      origem: "upload",
      arquivo: uploaded.originalName || uploaded.storedName || path.basename(uploaded.absolute),
      enviadoEm: uploaded.uploadedAt || null,
      atualizadoPor: uploaded.uploadedBy || null,
      atualizadoPorEmail: uploaded.uploadedByEmail || null
    };
  }

  const fallback = findWorkbookPath();
  if (fallback) {
    return {
      origem: "padrão",
      arquivo: path.basename(fallback),
      enviadoEm: null,
      atualizadoPor: null,
      atualizadoPorEmail: null
    };
  }

  return { origem: "nenhuma", arquivo: null, enviadoEm: null, atualizadoPor: null, atualizadoPorEmail: null };
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

function getWorkbookRows() {
  const workbookPath = findUploadedWorkbookPath() || findWorkbookPath();
  if (!workbookPath) return null;

  try {
    const stat = fs.statSync(workbookPath);
    const cacheKey = `${workbookPath}|${stat.mtimeMs}|${stat.size}`;
    if (workbookRowsCache?.cacheKey === cacheKey) {
      return workbookRowsCache.rows;
    }

    const workbook = XLSX.readFile(workbookPath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    workbookRowsCache = { cacheKey, rows };
    return rows;
  } catch (error) {
    console.error(`Erro ao ler planilha: ${error.message}`);
    workbookRowsCache = null;
    return null;
  }
}

function clearWorkbookCache() {
  workbookRowsCache = null;
}

function getPracaRows() {
  const workbookPath = findPracaWorkbookPath();
  if (!workbookPath) return [];

  try {
    preparePracaWorkbookIfNeeded(workbookPath);
    const stat = fs.statSync(workbookPath);
    const cacheKey = `${workbookPath}|${stat.mtimeMs}|${stat.size}`;
    if (pracaRowsCache?.cacheKey === cacheKey) {
      return pracaRowsCache.rows;
    }

    const workbook = XLSX.readFile(workbookPath);
    const sheetName = workbook.SheetNames.includes("Compilado") ? "Compilado" : workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "", raw: false });
    pracaRowsCache = { cacheKey, rows };
    pracaConfigCache = null;
    return rows;
  } catch (error) {
    console.error(`Erro ao ler planilha de praças: ${error.message}`);
    pracaRowsCache = null;
    pracaConfigCache = null;
    return [];
  }
}

function getClassificationRows() {
  const workbookPath = findClassificationWorkbookPath();
  if (!workbookPath) return [];

  try {
    const stat = fs.statSync(workbookPath);
    const cacheKey = `${workbookPath}|${stat.mtimeMs}|${stat.size}`;
    if (classificationRowsCache?.cacheKey === cacheKey) {
      return classificationRowsCache.rows;
    }

    const workbook = XLSX.readFile(workbookPath, { raw: false });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "", raw: false });
    classificationRowsCache = { cacheKey, rows };
    return rows;
  } catch (error) {
    console.error(`Erro ao ler classificacoes: ${error.message}`);
    classificationRowsCache = null;
    return [];
  }
}

function normalizePracaKey(value) {
  return normalizeText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\s*\/\s*/g, "/")
    .replace(/\s+/g, " ")
    .trim();
}

function getDefaultValorSacaPorPraca() {
  return new Map([
    ["posse/go", { nome: "POSSE/GO", valor: 104.05 }],
    ["buritis/mg", { nome: "BURITIS/MG", valor: 105.45 }],
    ["formosa/go", { nome: "FORMOSA/GO", valor: 106.55 }],
    ["primavera do leste/mt", { nome: "PRIMAVERA DO LESTE/MT", valor: 106.60 }],
    ["brasilia/df", { nome: "BRASILIA/DF", valor: 106.80 }],
    ["unai/mg", { nome: "UNAI/MG", valor: 107.45 }],
    ["rondonopolis/mt", { nome: "RONDONOPOLIS/MT", valor: 108.78 }],
    ["rio verde/go", { nome: "RIO VERDE/GO", valor: 109.33 }],
    ["cristalina/go", { nome: "CRISTALINA/GO", valor: 110.33 }],
    ["patrocinio/mg", { nome: "PATROCINIO/MG", valor: 111.75 }],
    ["barreiras/ba", { nome: "BARREIRAS/BA", valor: 112.00 }],
    ["uberlandia/mg", { nome: "UBERLANDIA/MG", valor: 115.75 }],
    ["sao luis/ma", { nome: "SAO LUIS/MA", valor: 123.80 }]
  ]);
}

function getValorSacaPorPraca() {
  const workbookPath = findValorSacaWorkbookPath();
  const defaults = getDefaultValorSacaPorPraca();
  if (!workbookPath) return new Map([...defaults].map(([key, item]) => [key, item.valor]));

  try {
    const stat = fs.statSync(workbookPath);
    const cacheKey = `${workbookPath}|${stat.mtimeMs}|${stat.size}`;
    if (valorSacaCache?.cacheKey === cacheKey) {
      return valorSacaCache.valores;
    }

    const workbook = XLSX.readFile(workbookPath);
    const valores = new Map([...defaults].map(([key, item]) => [key, item.valor]));
    const nomes = new Map([...defaults].map(([key, item]) => [key, item.nome]));

    for (const sheetName of workbook.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: false });
      for (const row of rows) {
        const praca = normalizeText(row[0]);
        const valor = parseFlexibleNumber(row[1]);
        const pracaKey = normalizePracaKey(praca);

        if (!pracaKey || !valor) continue;
        if (pracaKey.includes("locais") || pracaKey.includes("calculo")) continue;
        if (!valores.has(pracaKey)) valores.set(pracaKey, valor);
        if (!nomes.has(pracaKey)) nomes.set(pracaKey, praca);
      }
    }

    valorSacaCache = { cacheKey, valores, nomes };
    return valores;
  } catch (error) {
    console.error(`Erro ao ler planilha de valor da saca: ${error.message}`);
    valorSacaCache = null;
    return new Map();
  }
}

function shouldIgnoreCultivar(cultivar) {
  return normalizeComparable(cultivar).includes("soja intacta");
}

function isCalculationRow(row) {
  return !shouldIgnoreCultivar(extractCultivar(row));
}

function getCalculationRows() {
  const rows = getWorkbookRows();
  return rows ? rows.filter(isCalculationRow) : null;
}

function getRowValue(row, candidates) {
  for (const key of candidates) {
    if (Object.prototype.hasOwnProperty.call(row, key) && row[key] !== null && row[key] !== undefined && row[key] !== "") {
      return row[key];
    }
  }

  const normalizedCandidates = new Set(candidates.map(normalizeColumnKey));
  for (const [key, value] of Object.entries(row)) {
    if (
      normalizedCandidates.has(normalizeColumnKey(key)) &&
      value !== null &&
      value !== undefined &&
      value !== ""
    ) {
      return value;
    }
  }
  return "";
}

function normalizeText(value) {
  return String(value || "").trim();
}

function normalizeColumnKey(value) {
  return normalizeText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function normalizeComparable(value) {
  return normalizeText(value).toLowerCase();
}

function parseFlexibleNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;

  let text = normalizeText(value);
  if (!text || text === "-") return 0;

  text = text
    .replace(/%/g, "")
    .replace(/r\$/gi, "")
    .replace(/\s/g, "")
    .replace(/[^\d,.-]/g, "");

  if (!text || text === "-") return 0;

  const lastComma = text.lastIndexOf(",");
  const lastDot = text.lastIndexOf(".");

  if (lastComma >= 0 && lastDot >= 0) {
    if (lastComma > lastDot) {
      text = text.replace(/\./g, "").replace(",", ".");
    } else {
      text = text.replace(/,/g, "");
    }
  } else if (lastComma >= 0) {
    text = text.replace(",", ".");
  }

  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : 0;
}

function cleanNumericCode(value) {
  const text = normalizeText(value);
  if (!text) return "";
  if (/^\d+(\.0+)?$/.test(text)) return String(Number(text));
  return text;
}

function extractFornecedorCodigo(row) {
  return cleanNumericCode(getRowValue(row, ["Fornecedor", "Codigo Fornecedor", "Código Fornecedor"]));
}

function extractFornecedorNome(row) {
  return normalizeText(getRowValue(row, ["Nome (Fornecedor)", "Fornecedor"]));
}

function extractCentro(row) {
  return normalizeText(getRowValue(row, ["Centro", "Unidade"]));
}

function extractCultivar(row) {
  return normalizeText(getRowValue(row, ["Descrição (Material)", "Descricao (Material)", "Material"]));
}

function prefixSojaSemente(cultivar) {
  const value = normalizeText(cultivar);
  if (!value) return "";
  return /^soja\s+semente\b/i.test(value) ? value : `SOJA SEMENTE ${value}`;
}

function extractQualityProdutor(row) {
  return normalizeText(getRowValue(row, ["Cliente", "Nome (Fornecedor)", "Fornecedor"]));
}

function extractQualityCentro(row) {
  return normalizeText(getRowValue(row, ["Unidade", "Centro"]));
}

function extractQualityCultivar(row) {
  return prefixSojaSemente(getRowValue(row, ["Cultivar", "Descricao (Material)", "Material"]));
}

function extractQualityPesoLiquido(row) {
  return parseFlexibleNumber(getRowValue(row, [
    "Peso Liquido",
    "Peso Liquido Kg",
    "Peso Base (kg)"
  ]));
}

function extractQualityNota(row) {
  const nota = normalizeText(getRowValue(row, ["Nota", "Classificacao"]))
    .toUpperCase()
    .trim();

  if (nota === "G") return null;
  if (nota === "A") return 1;
  if (nota === "B") return 2;
  if (nota === "C") return 3;

  const parsed = parseFlexibleNumber(nota);
  return parsed > 0 ? parsed : null;
}

function extractMaterialCodigo(row) {
  return cleanNumericCode(getRowValue(row, ["Material", "Codigo Material", "Código Material"]));
}

function extractInscricao(row) {
  return cleanNumericCode(getRowValue(row, ["Inscrição Estadual", "Inscricao Estadual", "Fornecedor", "ID Fiscal"]));
}

function extractPesoLiquido(row) {
  return parseFlexibleNumber(getRowValue(row, ["Peso Liquido Kg", "Peso Líquido Kg", "Peso Base (kg)", "Recepção Kg", "Recepcao Kg"]));
}

function extractPesoBruto(row) {
  return parseFlexibleNumber(getRowValue(row, ["Peso Carga Bruto Kg", "Peso Bruto Kg", "Peso Inicial Kg"]));
}

function extractQualidade(row) {
  const rawValue = getRowValue(row, [
    "Qualidade",
    "Qualidade (%)",
    "% Qualidade",
    "Nota Qualidade",
    "Nota de Qualidade",
    "Pontuacao Qualidade",
    "Pontuação Qualidade",
    "Score Qualidade",
    "Score de Qualidade",
    "Indice Qualidade",
    "Índice Qualidade",
    "Germinacao (%)",
    "Germinação (%)",
    "% Germinacao",
    "% Germinação",
    "Vigor (%)",
    "% Vigor"
  ]);

  if (rawValue === "" || rawValue === null || rawValue === undefined) return null;

  const parsed = parseFlexibleNumber(rawValue);
  const text = normalizeText(rawValue);
  const isExplicitZero = /^0([,.]0+)?%?$/.test(text);
  if (!parsed && !isExplicitZero) return null;

  return parsed > 0 && parsed <= 1 ? parsed * 100 : parsed;
}

function extractPesoBase(row) {
  return extractPesoLiquido(row);
}

function extractNf(row) {
  return cleanNumericCode(getRowValue(row, ["Nº.da NF do Produtor", "N°.da NF do Produtor", "NF"]));
}

function extractRomaneio(row) {
  return cleanNumericCode(getRowValue(row, ["Romaneio"]));
}

function extractDirecaoMovimento(row) {
  return normalizeText(getRowValue(row, ["Direção do Movimento", "Direcao do Movimento"]));
}

function isEntradaRow(row) {
  return normalizeComparable(extractDirecaoMovimento(row)) === "entrada";
}

function createRankingItem(base) {
  return {
    ...base,
    pesoBruto: 0,
    pesoLiquido: 0,
    entradas: 0,
    romaneios: new Set()
  };
}

function addRankingRow(item, row) {
  item.pesoBruto += extractPesoBruto(row);
  item.pesoLiquido += extractPesoLiquido(row);
  item.entradas += 1;

  const romaneio = extractRomaneio(row);
  if (romaneio) item.romaneios.add(romaneio);
}

function addProducerCentro(item, row, cache) {
  const centro = extractCentro(row);
  if (!centro) return;

  const unit = getRankingUnitInfo(centro, cache);
  const key = unit.centroCode || normalizeComparable(centro);
  if (!item.centros) item.centros = new Map();
  item.centros.set(key, unit.centroCode || extractCentroCode(centro) || centro);
}

function finalizeRankingItems(groups) {
  return [...groups.values()]
    .map(item => {
      const { romaneios, centros, ...rest } = item;
      return {
        ...rest,
        romaneios: romaneios instanceof Set ? romaneios.size : 0,
        ...(centros instanceof Map
          ? { centros: [...centros.values()].sort((a, b) => a.localeCompare(b, "pt-BR")) }
          : {})
      };
    })
    .sort((a, b) =>
      (b.pesoLiquido - a.pesoLiquido) ||
      (b.pesoBruto - a.pesoBruto) ||
      String(a.label || "").localeCompare(String(b.label || ""), "pt-BR")
    );
}

function getRankingUnitInfo(centro, cache) {
  const centroCode = extractCentroCode(centro);
  if (cache.has(centroCode)) return cache.get(centroCode);

  const unit = findUnitByCentro(centro);
  const label = unit.centro || centro;
  const cidadeUf = [unit.cidade, unit.uf].filter(Boolean).join("/");
  const info = {
    centro,
    centroCode,
    unidade: label,
    cidade: unit.cidade || "",
    uf: unit.uf || "",
    local: cidadeUf
  };

  cache.set(centroCode, info);
  return info;
}

function buildProdutorRanking(rows) {
  const groups = new Map();
  const unitCache = new Map();

  for (const row of rows) {
    const produtor = extractFornecedorNome(row);
    if (!produtor) continue;

    const key = normalizeComparable(produtor);
    if (!groups.has(key)) {
      groups.set(key, createRankingItem({
        id: key,
        tipo: "produtor",
        label: produtor,
        produtor,
        codigo: extractFornecedorCodigo(row)
      }));
    }

    const item = groups.get(key);
    if (!item.codigo) item.codigo = extractFornecedorCodigo(row);
    addRankingRow(item, row);
    addProducerCentro(item, row, unitCache);
  }

  return finalizeRankingItems(groups);
}

function buildUnidadeRanking(rows) {
  const groups = new Map();
  const unitCache = new Map();

  for (const row of rows) {
    const centro = extractCentro(row);
    if (!centro) continue;

    const unit = getRankingUnitInfo(centro, unitCache);
    const key = unit.centroCode || normalizeComparable(centro);
    if (!groups.has(key)) {
      groups.set(key, createRankingItem({
        id: key,
        tipo: "unidade",
        label: unit.unidade,
        ...unit
      }));
    }

    addRankingRow(groups.get(key), row);
  }

  return finalizeRankingItems(groups);
}

function buildCultivarRanking(rows) {
  const groups = new Map();

  for (const row of rows) {
    const cultivar = extractCultivar(row);
    if (!cultivar) continue;

    const codigo = extractMaterialCodigo(row);
    const key = normalizeComparable(cultivar);
    if (!groups.has(key)) {
      groups.set(key, createRankingItem({
        id: key,
        tipo: "cultivar",
        label: cultivar,
        cultivar,
        codigo
      }));
    }

    const item = groups.get(key);
    if (!item.codigo && codigo) item.codigo = codigo;
    addRankingRow(item, row);
  }

  return finalizeRankingItems(groups);
}

function buildCultivarRankingsByCentro(rows) {
  const rowsByCentro = new Map();
  const unitCache = new Map();

  for (const row of rows) {
    const centro = extractCentro(row);
    const cultivar = extractCultivar(row);
    if (!centro || !cultivar) continue;

    const unit = getRankingUnitInfo(centro, unitCache);
    const key = unit.centroCode || normalizeComparable(centro);

    if (!rowsByCentro.has(key)) {
      rowsByCentro.set(key, {
        id: key,
        label: unit.unidade || centro,
        centro: unit.centro || centro,
        centroCode: unit.centroCode || key,
        local: unit.local || "",
        rows: []
      });
    }

    rowsByCentro.get(key).rows.push(row);
  }

  return [...rowsByCentro.values()]
    .map(({ rows: centroRows, ...centro }) => ({
      ...centro,
      ranking: buildCultivarRanking(centroRows)
    }))
    .sort((a, b) => String(a.label || "").localeCompare(String(b.label || ""), "pt-BR"));
}

function createQualityRankingItem(base) {
  return {
    ...base,
    qualidadeSomaPonderada: 0,
    qualidadePeso: 0,
    qualidadeSomaSimples: 0,
    amostras: 0,
    pesoA: 0,
    pesoB: 0,
    pesoC: 0,
    pesoBruto: 0,
    pesoLiquido: 0
  };
}

function addQualityRow(item, row) {
  const nota = extractQualityNota(row);
  if (nota === null) return false;

  const pesoLiquido = extractQualityPesoLiquido(row);
  const pesoBruto = 0;

  if (nota === 1) item.pesoA += pesoLiquido;
  if (nota === 2) item.pesoB += pesoLiquido;
  if (nota === 3) item.pesoC += pesoLiquido;
  item.amostras += 1;
  item.pesoLiquido += pesoLiquido;
  item.pesoBruto += pesoBruto;
  return true;
}

function finalizeQualityRankingItems(groups) {
  return [...groups.values()]
    .map(item => {
      const {
        qualidadeSomaPonderada,
        qualidadePeso,
        qualidadeSomaSimples,
        ...rest
      } = item;
      return {
        ...rest,
        percentualA: rest.pesoLiquido ? (rest.pesoA / rest.pesoLiquido) * 100 : 0
      };
    })
    .filter(item => item.amostras > 0 && item.pesoLiquido > 0)
    .sort((a, b) =>
      (b.percentualA - a.percentualA) ||
      (b.pesoLiquido - a.pesoLiquido) ||
      String(a.label || "").localeCompare(String(b.label || ""), "pt-BR")
    );
}

function buildQualityProdutorRanking(rows) {
  const groups = new Map();
  const unitCache = new Map();

  for (const row of rows) {
    const produtor = extractQualityProdutor(row);
    if (!produtor) continue;

    const key = normalizeComparable(produtor);
    if (!groups.has(key)) {
      groups.set(key, createQualityRankingItem({
        id: key,
        tipo: "produtor",
        label: produtor,
        produtor,
        codigo: "",
        centros: new Map()
      }));
    }

    const item = groups.get(key);
    if (addQualityRow(item, row)) addProducerCentro(item, row, unitCache);
  }

  return finalizeQualityRankingItems(groups).map(({ centros, ...item }) => ({
    ...item,
    centros: centros instanceof Map
      ? [...centros.values()].sort((a, b) => a.localeCompare(b, "pt-BR"))
      : []
  }));
}

function buildQualityUnidadeRanking(rows) {
  const groups = new Map();
  const unitCache = new Map();

  for (const row of rows) {
    const centro = extractQualityCentro(row);
    if (!centro) continue;

    const unit = getRankingUnitInfo(centro, unitCache);
    const key = unit.centroCode || normalizeComparable(centro);
    if (!groups.has(key)) {
      groups.set(key, createQualityRankingItem({
        id: key,
        tipo: "unidade",
        label: unit.unidade,
        ...unit
      }));
    }

    addQualityRow(groups.get(key), row);
  }

  return finalizeQualityRankingItems(groups);
}

function buildQualityCultivarRanking(rows) {
  const groups = new Map();

  for (const row of rows) {
    const cultivar = extractQualityCultivar(row);
    if (!cultivar) continue;

    const key = normalizeComparable(cultivar);
    if (!groups.has(key)) {
      groups.set(key, createQualityRankingItem({
        id: key,
        tipo: "cultivar",
        label: cultivar,
        cultivar,
        codigo: ""
      }));
    }

    const item = groups.get(key);
    addQualityRow(item, row);
  }

  return finalizeQualityRankingItems(groups);
}

function buildQualityProdutorRankingsByCentro(rows) {
  const rowsByCentro = new Map();
  const unitCache = new Map();

  for (const row of rows) {
    const centro = extractQualityCentro(row);
    const produtor = extractQualityProdutor(row);
    if (!centro || !produtor) continue;

    const unit = getRankingUnitInfo(centro, unitCache);
    const key = unit.centroCode || normalizeComparable(centro);

    if (!rowsByCentro.has(key)) {
      rowsByCentro.set(key, {
        id: key,
        label: unit.unidade || centro,
        centro: unit.centro || centro,
        centroCode: unit.centroCode || key,
        local: unit.local || "",
        rows: []
      });
    }

    rowsByCentro.get(key).rows.push(row);
  }

  return [...rowsByCentro.values()]
    .map(({ rows: centroRows, ...centro }) => ({
      ...centro,
      ranking: buildQualityProdutorRanking(centroRows)
    }))
    .sort((a, b) => String(a.label || "").localeCompare(String(b.label || ""), "pt-BR"));
}

function buildQualityRankings(rows) {
  return {
    produtores: buildQualityProdutorRanking(rows),
    unidades: buildQualityUnidadeRanking(rows),
    cultivares: buildQualityCultivarRanking(rows),
    produtoresPorCentro: buildQualityProdutorRankingsByCentro(rows)
  };
}

function buildWorkbookRankings() {
  const rows = getWorkbookRows();
  if (!rows) return null;

  const entradaRows = rows.filter(isEntradaRow);
  const sementesRows = entradaRows.filter(row => !shouldIgnoreCultivar(extractCultivar(row)));
  const produtoresRows = entradaRows.filter(row => Boolean(extractFornecedorNome(row)));
  const classificationRows = getClassificationRows();

  return {
    geradoEm: new Date().toISOString(),
    planilha: getWorkbookInfo(),
    totais: {
      linhas: rows.length,
      entradas: entradaRows.length,
      entradasComProdutor: produtoresRows.length,
      entradasSemSojaIntacta: sementesRows.length
    },
    filtros: {
      produtores: "Entradas com produtor informado",
      unidades: "Entradas sem Soja Intacta",
      cultivares: "Entradas sem Soja Intacta, agrupadas por cultivar no geral",
      qualidade: "Soma de Peso Líquido por classificação A, B e C, descartando G"
    },
    rankings: {
      produtores: buildProdutorRanking(sementesRows.filter(row => Boolean(extractFornecedorNome(row)))),
      unidades: buildUnidadeRanking(sementesRows),
      cultivares: buildCultivarRanking(sementesRows),
      cultivaresPorCentro: buildCultivarRankingsByCentro(sementesRows),
      qualidade: buildQualityRankings(classificationRows)
    }
  };
}

function buildUnidadesResumo() {
  const rows = getWorkbookRows();
  if (!rows) return null;

  const groups = new Map();
  const unitCache = new Map();
  const entradaRows = rows.filter(isEntradaRow).filter(row => !shouldIgnoreCultivar(extractCultivar(row)));

  for (const row of entradaRows) {
    const centro = extractCentro(row);
    if (!centro) continue;

    const unit = getRankingUnitInfo(centro, unitCache);
    const key = unit.centroCode || normalizeComparable(centro);

    if (!groups.has(key)) {
      groups.set(key, {
        id: key,
        centro: unit.centro || centro,
        centroCode: unit.centroCode || extractCentroCode(centro) || "",
        unidade: unit.unidade || centro,
        cidade: unit.cidade || "",
        uf: unit.uf || "",
        local: unit.local || "",
        entradas: 0,
        pesoBruto: 0,
        pesoLiquido: 0,
        produtores: new Set(),
        cultivares: new Set(),
        romaneios: new Set()
      });
    }

    const item = groups.get(key);
    item.entradas += 1;
    item.pesoBruto += extractPesoBruto(row);
    item.pesoLiquido += extractPesoLiquido(row);

    const produtor = extractFornecedorNome(row);
    const cultivar = extractCultivar(row);
    const romaneio = extractRomaneio(row);
    if (produtor) item.produtores.add(normalizeComparable(produtor));
    if (cultivar) item.cultivares.add(normalizeComparable(cultivar));
    if (romaneio) item.romaneios.add(romaneio);
  }

  return {
    geradoEm: new Date().toISOString(),
    planilha: getWorkbookInfo(),
    totais: {
      centros: groups.size,
      entradas: entradaRows.length,
      pesoBruto: [...groups.values()].reduce((sum, item) => sum + item.pesoBruto, 0),
      pesoLiquido: [...groups.values()].reduce((sum, item) => sum + item.pesoLiquido, 0)
    },
    unidades: [...groups.values()]
      .map(item => ({
        ...item,
        produtores: item.produtores.size,
        cultivares: item.cultivares.size,
        romaneios: item.romaneios.size,
        pesoBrutoFmt: formatNumberPtBr(item.pesoBruto),
        pesoLiquidoFmt: formatNumberPtBr(item.pesoLiquido)
      }))
      .sort((a, b) =>
        (b.pesoLiquido - a.pesoLiquido) ||
        String(a.unidade || "").localeCompare(String(b.unidade || ""), "pt-BR")
      )
  };
}

function rankingNumberFields(item) {
  return {
    ...item,
    pesoBrutoFmt: formatNumberPtBr(item.pesoBruto),
    pesoLiquidoFmt: formatNumberPtBr(item.pesoLiquido)
  };
}

function isUnidadeMatch(unit, centro, key, unidadeId) {
  const target = normalizeComparable(unidadeId);
  return [
    key,
    unit?.centroCode,
    unit?.centro,
    unit?.unidade,
    centro
  ].some(value => normalizeComparable(value) === target);
}

function buildUnidadeDetalhe(unidadeId) {
  const rows = getWorkbookRows();
  if (!rows) return null;

  const unitCache = new Map();
  const selectedRows = [];
  let unitInfo = null;

  const entradaRows = rows
    .filter(isEntradaRow)
    .filter(row => !shouldIgnoreCultivar(extractCultivar(row)));

  for (const row of entradaRows) {
    const centro = extractCentro(row);
    if (!centro) continue;

    const unit = getRankingUnitInfo(centro, unitCache);
    const key = unit.centroCode || normalizeComparable(centro);
    if (!isUnidadeMatch(unit, centro, key, unidadeId)) continue;

    selectedRows.push(row);
    unitInfo = unit;
  }

  if (!selectedRows.length) {
    return {
      geradoEm: new Date().toISOString(),
      planilha: getWorkbookInfo(),
      unidade: null,
      topCultivares: [],
      topProdutores: []
    };
  }

  const nfs = new Set();
  const produtores = new Set();
  const cultivares = new Set();
  selectedRows.forEach(row => {
    const nf = extractNf(row);
    const produtor = extractFornecedorNome(row);
    const cultivar = extractCultivar(row);
    if (nf) nfs.add(nf);
    if (produtor) produtores.add(normalizeComparable(produtor));
    if (cultivar) cultivares.add(normalizeComparable(cultivar));
  });

  const unidade = rankingNumberFields(buildUnidadeRanking(selectedRows)[0] || {
    id: unitInfo?.centroCode || normalizeComparable(unidadeId),
    label: unitInfo?.unidade || unidadeId,
    unidade: unitInfo?.unidade || unidadeId,
    centroCode: unitInfo?.centroCode || unidadeId,
    local: unitInfo?.local || "",
    pesoBruto: 0,
    pesoLiquido: 0,
    entradas: selectedRows.length,
    romaneios: 0
  });

  return {
    geradoEm: new Date().toISOString(),
    planilha: getWorkbookInfo(),
    unidade: {
      ...unitInfo,
      ...unidade,
      nfs: nfs.size,
      produtores: produtores.size,
      cultivares: cultivares.size
    },
    topCultivares: buildCultivarRanking(selectedRows).slice(0, 8).map(rankingNumberFields),
    topProdutores: buildProdutorRanking(selectedRows).slice(0, 8).map(rankingNumberFields)
  };
}

function extractCNPJ(row) {
  return cleanNumericCode(getRowValue(row, ["CNPJ", "CNPJ Fornecedor", "CNPJ do Fornecedor"]));
}

function extractPracaPercentual(row) {
  return parseFlexibleNumber(getRowValue(row, [
    "% DE BONIFICAÇÃO",
    "% DE BONIFICACAO",
    "Porcentagem (%)",
    "Porcentagem",
    "Percentual",
    "% Bonificação",
    "% Bonificacao"
  ]));
}

function extractPracaVolumeLiquido(row) {
  return parseFlexibleNumber(getRowValue(row, [
    "Volume liquido entregue",
    "Volume líquido entregue",
    "Volume Liquido Entregue",
    "Volume Líquido Entregue",
    "Peso Liquido Kg",
    "Peso Líquido Kg",
    "Peso Base (kg)"
  ]));
}

function extractPracaValorBonificado(row) {
  return parseFlexibleNumber(getRowValue(row, [
    "Valor Bonificado",
    "Valor bonificado",
    "Valor da Bonificação",
    "Valor da Bonificacao",
    "Volume Bonificado",
    "Volume bonificado",
    "Peso Bonificado",
    "Peso bonificado"
  ]));
}

function calculatePracaPercentual(row) {
  const valorBonificado = extractPracaValorBonificado(row);
  const volumeLiquido = extractPracaVolumeLiquido(row);

  if (valorBonificado > 0 && volumeLiquido > 0) {
    const pctBonificacao = valorBonificado / volumeLiquido;
    return pctBonificacao <= 1 ? pctBonificacao * 100 : pctBonificacao;
  }

  return extractPracaPercentual(row) || null;
}

function extractPracaValor(row) {
  return parseFlexibleNumber(getRowValue(row, [
    "Valor da Saca (R$)",
    "Valor da Saca",
    "Valor Saca",
    "Valor da saca R$",
    "R$ Saca",
    "Preco da Saca",
    "Preço da Saca",
    "Valor"
  ]));
}

function getValorSacaByPraca(praca) {
  const pracaKey = normalizePracaKey(praca);
  if (!pracaKey) return 0;
  return getValorSacaPorPraca().get(pracaKey) || 0;
}

function extractPracaNome(row) {
  return normalizeText(getRowValue(row, [
    "PRACA",
    "Praça",
    "Praca"
  ]));
}

function getMostFrequentNumber(values) {
  const counts = new Map();
  for (const value of values) {
    const numeric = Number(value) || 0;
    if (!numeric) continue;
    counts.set(numeric, (counts.get(numeric) || 0) + 1);
  }

  let selected = 0;
  let selectedCount = 0;
  for (const [value, count] of counts.entries()) {
    if (count > selectedCount || (count === selectedCount && value > selected)) {
      selected = value;
      selectedCount = count;
    }
  }

  return selected;
}

function getPracaOptions() {
  const valorPorPraca = getValorSacaPorPraca();
  const optionsByKey = new Map();

  for (const row of getPracaRows()) {
    const praca = extractPracaNome(row);
    const key = normalizePracaKey(praca);
    if (!key) continue;

    if (!optionsByKey.has(key)) {
      optionsByKey.set(key, {
        key,
        praca,
        valor: valorPorPraca.get(key) || 0,
        percs: []
      });
    }

    const perc = calculatePracaPercentual(row);
    if (perc) optionsByKey.get(key).percs.push(perc);
  }

  return [...optionsByKey.values()]
    .map(option => {
      const perc = getMostFrequentNumber(option.percs);
      return {
        key: option.key,
        praca: option.praca,
        perc,
        percFmt: formatNumberPtBr(perc),
        valor: option.valor,
        valorFmt: formatNumberPtBr(option.valor)
      };
    })
    .sort((a, b) => a.praca.localeCompare(b.praca, "pt-BR"));
}

function getPracaOptionDefaults(praca) {
  const pracaKey = normalizePracaKey(praca);
  if (!pracaKey) return { perc: 0, valor: 0, praca: "" };

  return getPracaOptions().find(option => option.key === pracaKey) || {
    perc: 0,
    valor: getValorSacaByPraca(praca),
    praca
  };
}

function getSupplierMatchKeys(rowOrItem) {
  const keys = [
    extractFornecedorCodigo(rowOrItem),
    extractInscricao(rowOrItem),
    normalizeText(rowOrItem.fornecedorCodigo),
    normalizeText(rowOrItem.inscricaoEstadual),
    normalizeText(rowOrItem.fornecedorNome),
    extractFornecedorNome(rowOrItem)
  ]
    .map(value => normalizeComparable(cleanNumericCode(value) || value))
    .filter(Boolean);

  return [...new Set(keys)];
}

function makePracaKey({ centro, fornecedor, cultivar }) {
  return [
    normalizeComparable(centro),
    normalizeComparable(cleanNumericCode(fornecedor) || fornecedor),
    normalizeComparable(cultivar)
  ].join("|||");
}

function buildPracaConfigIndex() {
  const rows = getPracaRows();
  getValorSacaPorPraca();
  const cacheKey = `${pracaRowsCache?.cacheKey || ""}|${valorSacaCache?.cacheKey || ""}`;
  if (pracaConfigCache?.cacheKey === cacheKey) return pracaConfigCache;

  const byFull = new Map();

  for (const row of rows) {
    const centro = extractCentro(row);
    const cultivar = extractCultivar(row);
    const supplierKeys = getSupplierMatchKeys(row);

    if (!centro || !cultivar || !supplierKeys.length) continue;

    const praca = extractPracaNome(row);
    const valorNaLinha = extractPracaValor(row);
    const config = {
      perc: calculatePracaPercentual(row),
      valor: valorNaLinha || getValorSacaByPraca(praca),
      praca
    };

    for (const fornecedor of supplierKeys) {
      const fullKey = makePracaKey({ centro, fornecedor, cultivar });
      if (!byFull.has(fullKey)) byFull.set(fullKey, config);
    }
  }

  pracaConfigCache = { cacheKey, byFull };
  return pracaConfigCache;
}

function getPracaDefaultsForItem(item) {
  const index = buildPracaConfigIndex();

  for (const fornecedor of getSupplierMatchKeys(item)) {
    const fullKey = makePracaKey({ ...item, fornecedor });
    const fullConfig = index.byFull.get(fullKey);
    if (fullConfig) return fullConfig;
  }

  return { perc: null, valor: 0, praca: "" };
}

function getItemConfig(item, body) {
  if (body.modoGlobal !== false) {
    return {
      perc: Number(body.globalPerc) || 0,
      valor: Number(body.globalValor) || 0,
      praca: item.praca || ""
    };
  }

  const key = `${item.inscricaoEstadual}|||${item.centro}|||${item.cultivar}`;
  const local = body.configs?.[key] || {};
  const praca = normalizeText(local.praca) || item.praca || "";
  const pracaDefaults = getPracaOptionDefaults(praca);

  return {
    perc: Number(local.perc) || pracaDefaults.perc || 0,
    valor: Number(local.valor) || pracaDefaults.valor || 0,
    praca
  };
}

// ============= GERAÇÃO DE NFe DE BONIFICAÇÃO =============

function generateNFeStructure(detalhes, fornecedorNome, totalGeralFmt, cnpjFornecedor = "") {
  // Gera número de série único para a NFe
  const nfeNumber = Math.floor(Math.random() * 1000000).toString().padStart(9, '0');
  const series = "1";
  const dataEmissao = new Date();
  
  // Processa rastreabilidade: agrupa as chaves/números de NFes referenciadas
  const nfsReferenciadas = new Map();
  detalhes.forEach(item => {
    const nfsArray = (item.nfs || "N/A").split(",").map(n => n.trim());
    nfsArray.forEach(nf => {
      if (nf !== "N/A" && nf) {
        if (!nfsReferenciadas.has(nf)) {
          nfsReferenciadas.set(nf, []);
        }
        nfsReferenciadas.get(nf).push({
          centro: item.centro,
          inscricao: item.inscricaoEstadual,
          cultivar: item.cultivar
        });
      }
    });
  });

  // Monta o texto de Informações Adicionais (InfAdic) com rastreabilidade
  const infAdic = Array.from(nfsReferenciadas.entries())
    .map(([nf, items]) => {
      const detalheNfs = items
        .map(i => `${i.centro}/${i.inscricao}/${i.cultivar}`)
        .join("; ");
      return `NF ${nf}: ${detalheNfs}`;
    })
    .join(" | ");

  // Agrupa itens por Inscrição Estadual (primeiro) e depois por Centro
  const itensPorIE = new Map();
  detalhes.forEach((item, index) => {
    const ie = item.inscricaoEstadual;
    const centro = item.centro;
    
    if (!itensPorIE.has(ie)) {
      itensPorIE.set(ie, new Map());
    }
    
    const centrosMap = itensPorIE.get(ie);
    if (!centrosMap.has(centro)) {
      centrosMap.set(centro, []);
    }
    
    centrosMap.get(centro).push({
      itemId: index + 1,
      codigo: `${ie.substring(0, 4)}-${centro}`,
      descricao: item.cultivar,
      ncm: "1001.90.00",
      cfop: "5101",
      quantidade: parseFloat(item.sacasFmt.replace(/\./g, '').replace(',', '.')),
      unidade: "SC",
      valorUnitario: item.valor,
      valorTotal: item.subtotal,
      centro: centro,
      inscricaoEstadual: ie,
      infAdProd: `Bonificação referente às NF: ${item.nfs}`
    });
  });

  const nfeItems = detalhes.map((item, index) => ({
    itemId: index + 1,
    codigo: `${item.inscricaoEstadual.substring(0, 4)}-${item.centro}`,
    descricao: item.cultivar,
    ncm: "1001.90.00",
    cfop: "5101",
    quantidade: parseFloat(item.sacasFmt.replace(/\./g, '').replace(',', '.')),
    unidade: "SC",
    valorUnitario: item.valor,
    valorTotal: item.subtotal,
    centro: item.centro,
    inscricaoEstadual: item.inscricaoEstadual,
    infAdProd: `Bonificação referente às NF: ${item.nfs}`
  }));

  return {
    tipo: "Nota Fiscal de Bonificação",
    numero: nfeNumber,
    serie: series,
    cfop: "5101",
    fornecedor: fornecedorNome,
    cnpj: cnpjFornecedor,
    dataEmissao: dataEmissao.toLocaleDateString('pt-BR'),
    dataEmissaoISO: dataEmissao.toISOString(),
    itens: nfeItems,
    itensPorIE: Object.fromEntries(
      Array.from(itensPorIE.entries()).map(([ie, centrosMap]) => [
        ie,
        Object.fromEntries(centrosMap)
      ])
    ),
    totalItens: nfeItems.length,
    valorTotal: totalGeralFmt,
    valorTotalNumerico: parseFloat(totalGeralFmt.replace(/\./g, '').replace(',', '.')),
    infAdic: infAdic || "Nota fiscal de bonificação gerada para fins de conferência e replicação manual",
    observacoes: "Nota fiscal de bonificação gerada para fins de conferência e replicação manual"
  };
}

function formatNFeForDisplay(nfeStructure) {
  // Cria cards flutuantes agrupados por Inscrição Estadual
  const cardsHtml = Object.entries(nfeStructure.itensPorIE || {})
    .map(([ie, centros]) => {
      const centrosHtml = Object.entries(centros)
        .map(([centro, itens]) => {
          const subtotalCentro = itens.reduce((sum, item) => sum + item.valorTotal, 0);
          const itemsHtml = itens
            .map((item, idx) => `
              <tr class="border-b border-[#cfdcbd] hover:bg-[#f8faf4]">
                <td class="px-3 py-2 text-xs font-medium text-[#203a28]">${idx + 1}</td>
                <td class="px-3 py-2 text-xs text-left text-[#38563f]">${item.descricao}</td>
                <td class="px-3 py-2 text-xs text-right text-[#38563f]">${item.quantidade.toFixed(2)}</td>
                <td class="px-3 py-2 text-xs text-center text-[#61713d]">SC</td>
                <td class="px-3 py-2 text-xs text-right font-semibold text-[#1f4d2d]">R$ ${item.valorUnitario.toFixed(2)}</td>
                <td class="px-3 py-2 text-xs text-right font-bold text-[#2f6b37]">R$ ${item.valorTotal.toFixed(2)}</td>
              </tr>
            `)
            .join('');

          return `
            <div class="rounded-lg border border-[#cfdcbd] bg-[#f8faf4] p-4 mb-3">
              <div class="flex items-center justify-between mb-3 pb-2 border-b-2 border-[#d8e5c6]">
                <h4 class="text-sm font-bold text-[#1f4d2d]">Centro: ${centro}</h4>
                <span class="text-xs font-semibold bg-[#2f6b37] text-white px-2 py-1 rounded">CFOP 5101</span>
              </div>
              <table class="w-full text-xs border-collapse">
                <thead class="bg-[#edf4e2]">
                  <tr>
                    <th class="border border-[#cfdcbd] px-2 py-2 text-left font-semibold text-[#38563f]">#</th>
                    <th class="border border-[#cfdcbd] px-2 py-2 text-left font-semibold text-[#38563f]">Cultivar</th>
                    <th class="border border-[#cfdcbd] px-2 py-2 text-right font-semibold text-[#38563f]">Qtd</th>
                    <th class="border border-[#cfdcbd] px-2 py-2 text-center font-semibold text-[#38563f]">UN</th>
                    <th class="border border-[#cfdcbd] px-2 py-2 text-right font-semibold text-[#38563f]">Vr Unit</th>
                    <th class="border border-[#cfdcbd] px-2 py-2 text-right font-semibold text-[#38563f]">Total</th>
                  </tr>
                </thead>
                <tbody>
                  ${itemsHtml}
                </tbody>
              </table>
              <div class="text-right mt-3 pt-3 border-t-2 border-[#d8e5c6]">
                <p class="text-xs text-[#61713d] font-semibold">Subtotal Centro: <span class="text-sm font-bold text-[#2f6b37]">R$ ${subtotalCentro.toFixed(2)}</span></p>
              </div>
            </div>
          `;
        })
        .join('');

      return `
        <div class="mb-6 p-5 rounded-xl border-2 border-[#2f6b37] bg-gradient-to-br from-white to-[#f8faf4] shadow-lg">
          <div class="flex items-center gap-3 mb-4 pb-4 border-b-3 border-[#7FB23D]">
            <div class="w-12 h-12 rounded-full bg-[#2f6b37] flex items-center justify-center text-white font-bold text-lg">
              ${ie.substring(0, 2)}
            </div>
            <div>
              <p class="text-xs text-[#61713d] font-semibold">INSCRIÇÃO ESTADUAL</p>
              <p class="text-lg font-bold text-[#1f4d2d]">${ie}</p>
            </div>
          </div>
          <div class="space-y-3">
            ${centrosHtml}
          </div>
        </div>
      `;
    })
    .join('');

  const infAdicHtml = nfeStructure.infAdic ? `
    <div class="bg-[#edf7e8] border-l-4 border-[#7FB23D] p-4 mt-5 rounded">
      <p class="text-xs font-semibold text-[#1f4d2d] mb-2">📋 Rastreabilidade (Dados Complementares):</p>
      <p class="text-xs text-[#38563f] leading-relaxed">${nfeStructure.infAdic}</p>
    </div>
  ` : '';

  const cnpjHtml = nfeStructure.cnpj ? `
    <div class="text-xs">
      <span class="font-semibold text-[#38563f]">CNPJ:</span>
      <span class="text-[#1f4d2d]">${nfeStructure.cnpj}</span>
    </div>
  ` : '';
  
  const numeroSerieHtml = nfeStructure.numero ? `
    <div class="text-xs">
      <span class="font-semibold text-[#38563f]">NF nº:</span>
      <span class="text-[#1f4d2d] font-mono">${nfeStructure.numero}</span>
      <span class="font-semibold text-[#38563f] ml-3">Série:</span>
      <span class="text-[#1f4d2d] font-mono">${nfeStructure.serie}</span>
    </div>
  ` : '';

  return {
    headerHtml: `
      <div class="rounded-lg border-3 border-[#2f6b37] p-6 bg-gradient-to-r from-[#edf7e8] to-white shadow-lg mb-5">
        <div class="text-center mb-6 pb-4 border-b-3 border-[#7FB23D]">
          <h3 class="text-2xl font-bold text-[#1f4d2d] mb-1">📄 NOTA FISCAL DE BONIFICAÇÃO</h3>
          <p class="text-sm font-semibold text-[#7FB23D]">Espelho para Conferência e Replicação Manual</p>
        </div>
        <div class="grid grid-cols-2 md:grid-cols-4 gap-4">
          <div>
            <p class="text-[#61713d] text-xs font-bold uppercase mb-2">Fornecedor</p>
            <p class="text-sm font-bold text-[#1f4d2d]">${nfeStructure.fornecedor}</p>
            ${cnpjHtml}
          </div>
          <div>
            <p class="text-[#61713d] text-xs font-bold uppercase mb-2">Fiscal</p>
            <p class="text-xs"><span class="font-semibold">Tipo:</span> ${nfeStructure.tipo}</p>
            <p class="text-xs"><span class="font-semibold">CFOP:</span> <span class="bg-[#2f6b37] text-white px-2 py-1 rounded font-bold">5101</span></p>
          </div>
          <div>
            <p class="text-[#61713d] text-xs font-bold uppercase mb-2">Emissão</p>
            <p class="text-xs"><span class="font-semibold">Data:</span> ${nfeStructure.dataEmissao}</p>
            ${numeroSerieHtml}
          </div>
          <div>
            <p class="text-[#61713d] text-xs font-bold uppercase mb-2">Resumo</p>
            <p class="text-xs"><span class="font-semibold">Itens:</span> ${nfeStructure.totalItens}</p>
            <p class="text-lg font-bold text-[#2f6b37]">R$ ${nfeStructure.valorTotal}</p>
          </div>
        </div>
      </div>
    `,
    itemsHtml: `
      <div class="mt-5">
        <p class="text-sm font-bold text-[#1f4d2d] mb-4">Detalhamento por Inscrição Estadual</p>
        ${cardsHtml}
        ${infAdicHtml}
      </div>
    `,
    footerHtml: `
      <div class="mt-6 space-y-4">
        <div class="rounded-lg border-2 border-[#7FB23D] p-5 bg-[#edf7e8] shadow-md">
          <div class="grid grid-cols-2 md:grid-cols-3 gap-4 mb-4">
            <div>
              <p class="text-[#61713d] text-xs font-bold">TOTAL DE ITENS</p>
              <p class="text-2xl font-bold text-[#1f4d2d]">${nfeStructure.totalItens}</p>
            </div>
            <div>
              <p class="text-[#61713d] text-xs font-bold">VALOR TOTAL</p>
              <p class="text-2xl font-bold text-[#2f6b37]">R$ ${nfeStructure.valorTotal}</p>
            </div>
            <div>
              <p class="text-[#61713d] text-xs font-bold">STATUS</p>
              <p id="nfeStatus" class="text-xs font-bold text-white bg-[#7FB23D] px-3 py-1 rounded-full inline-block">RASCUNHO</p>
            </div>
          </div>
          <p class="text-xs text-[#61713d] italic border-t border-[#cde5bd] pt-3">${nfeStructure.observacoes}</p>
        </div>
        
        <div class="flex flex-col md:flex-row gap-3">
          <button id="btnConfirmarNFe" type="button" class="flex-1 rounded-lg bg-[#2f6b37] hover:bg-[#25572d] active:scale-95 text-white font-bold py-3 px-4 transition-all duration-200 flex items-center justify-center gap-2 shadow-md">
            <i class="fa-solid fa-check-circle"></i>
            Confirmar e Emitir
          </button>
          <button id="btnReplicarNFe" type="button" class="flex-1 rounded-lg bg-[#7FB23D] hover:bg-[#6b9831] active:scale-95 text-white font-bold py-3 px-4 transition-all duration-200 flex items-center justify-center gap-2 shadow-md">
            <i class="fa-solid fa-copy"></i>
            Replicar Emissão
          </button>
          <button id="btnCancelarNFe" type="button" class="flex-1 rounded-lg border-2 border-[#d8e5c6] bg-white hover:bg-[#f8faf4] active:scale-95 text-[#38563f] font-bold py-3 px-4 transition-all duration-200 flex items-center justify-center gap-2 shadow-md">
            <i class="fa-solid fa-times-circle"></i>
            Cancelar
          </button>
        </div>
      </div>
    `
  };
}

async function getCurrentProfile(req) {
  if (!req.session.user && !req.session.userId) return null;

  let query = supabaseAdmin.from("usuarios").select("*");
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

async function listUserProfiles() {
  const { data, error } = await supabaseAdmin
    .from("usuarios")
    .select("id, usuario, email, admin, confirmado, created_at")
    .order("usuario", { ascending: true });

  if (error) throw error;
  return (data || []).map(profile => ({
    ...profile,
    admin: normalizeBoolean(profile.admin)
  }));
}

async function getUserProfileByUsuario(usuario) {
  const { data, error } = await supabaseAdmin
    .from("usuarios")
    .select("*")
    .eq("usuario", usuario)
    .maybeSingle();

  if (error) throw error;
  return data;
}

async function getUserProfileByEmail(email) {
  const { data, error } = await supabaseAdmin
    .from("usuarios")
    .select("*")
    .eq("email", normalizeEmail(email))
    .maybeSingle();

  if (error) throw error;
  return data;
}

async function getConfigValue(chave) {
  const { data, error } = await supabaseAdmin
    .from("configuracoes")
    .select("valor")
    .eq("chave", chave)
    .maybeSingle();

  if (error) throw error;
  return data?.valor || "";
}

async function setConfigValue(chave, valor) {
  const { error } = await supabaseAdmin
    .from("configuracoes")
    .upsert({ chave, valor, updated_at: new Date().toISOString() }, { onConflict: "chave" });

  if (error) throw error;
}

function getLocalDashboardLink() {
  const dashboardPath = path.join(ROOT_DIR, "Power BI - Romaneios.txt");
  if (!fs.existsSync(dashboardPath)) return "";
  return fs.readFileSync(dashboardPath, "utf8").trim();
}

async function saveHistoricoEntry({ usuarioId, usuario, produtorNome, totalGeralFmt, detalhes }) {
  const { error } = await supabaseAdmin.from("historico").insert({
    usuario_id: usuarioId,
    usuario,
    produtornome: produtorNome,
    totalgeralfmt: totalGeralFmt,
    detalhes,
    itens: detalhes?.length || 0
  });

  if (error) throw error;
}

function normalizeHistoricoRow(row) {
  return {
    ...row,
    produtorNome: row.produtorNome ?? row.produtornome ?? "",
    totalGeralFmt: row.totalGeralFmt ?? row.totalgeralfmt ?? "",
    createdAt: row.createdAt ?? row.createdat ?? row.created_at ?? ""
  };
}

// ============= AUTH ROUTES =============

app.post("/api/login", async (req, res) => {
  const email = normalizeEmail(req.body.email);
  const senha = String(req.body.senha || "").trim();

  if (!email || !senha) {
    return res.status(400).json({ erro: "E-mail e senha são obrigatórios." });
  }

  try {
    const { user, error, emailNotConfirmed } = await signInOrMigrateLegacyUser(email, senha);
    if (emailNotConfirmed) {
      return res.status(403).json({ erro: "E-mail não confirmado. Verifique seu e-mail corporativo." });
    }

    if (error || !user) {
      return res.status(401).json({ erro: "Usuário ou senha incorretos." });
    }

    let profile = await findOrCreateUserProfile(email, user.user_metadata?.usuario || "", senha, user.id);
    if (!profile) {
      return res.status(500).json({ erro: "Erro ao buscar o perfil do usuário." });
    }
    profile = await syncProfileConfirmation(profile, user);

    if (!profile.confirmado) {
      return res.status(403).json({ erro: "E-mail não confirmado. Verifique seu e-mail corporativo." });
    }

    req.session.user = profile.usuario;
    req.session.userId = profile.id;
    res.json({ usuario: profile.usuario, email: profile.email, admin: isAdminProfile(profile) });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao processar login." });
  }
});

app.post("/api/registrar", async (req, res) => {
  const usuario = sanitizeUsuario(req.body.usuario);
  const email = normalizeEmail(req.body.email);
  const senha = String(req.body.senha || "").trim();

  if (!usuario || !email || !senha) {
    return res.status(400).json({ erro: "Todos os campos são obrigatórios." });
  }

  if (!isEmailValid(email)) {
    return res.status(400).json({ erro: "E-mail deve ser @boasafrasementes.com.br" });
  }

  if (senha.length < 4) {
    return res.status(400).json({ erro: "A senha deve ter no mínimo 4 caracteres." });
  }

  try {
    if (await getUserProfileByUsuario(usuario)) {
      return res.status(409).json({ erro: "Usuário já existe." });
    }

    if (await getUserProfileByEmail(email)) {
      return res.status(409).json({ erro: "E-mail já existe." });
    }

    const { user, session, error: authError } = await signUpSupabaseAuthUser({ email, senha, usuario }, req);
    if (authError) {
      if (isSupabaseDuplicateAuthError(authError)) {
        return res.status(409).json({ erro: "E-mail já existe." });
      }
      if (isConfirmationEmailSendError(authError)) {
        console.error(`Erro ao enviar e-mail de confirmação pelo Supabase: ${authError.message}`);
        return res.status(500).json({ erro: "Não foi possível enviar o e-mail de confirmação. Verifique a configuração SMTP do Supabase." });
      }
      console.error(`Erro no cadastro Supabase: ${authError.message}`);
      return res.status(500).json({ erro: "Erro ao solicitar confirmação no Supabase." });
    }

    const profile = await findOrCreateUserProfile(email, usuario, senha, user?.id, {
      confirmado: Boolean(session) || isAuthUserConfirmed(user)
    });
    if (!profile) {
      return res.status(500).json({ erro: "Erro ao criar perfil do usuário." });
    }

    res.json({
      mensagem: profile.confirmado
        ? "Usuário criado com sucesso."
        : "Cadastro recebido. Enviamos um link de confirmação para o e-mail corporativo.",
      usuario: profile.usuario,
      email: profile.email,
      confirmado: profile.confirmado
    });
  } catch (error) {
    console.error(`Erro ao processar registro: ${error.message}`);
    res.status(500).json({ erro: "Erro ao processar registro." });
  }
});

app.post("/login", async (req, res) => {
  const email = normalizeEmail(req.body.email);
  const senha = String(req.body.senha || "").trim();

  console.log(`🔐 Tentativa de login: ${email}`);

  if (!email || !senha) {
    console.log("❌ Login falhou: campos obrigatórios ausentes");
    return res.redirect("/?erro=1");
  }

  try {
    const { user, error, emailNotConfirmed } = await signInOrMigrateLegacyUser(email, senha);
    if (emailNotConfirmed) {
      console.log(`❌ Login bloqueado: email não confirmado - ${email}`);
      return res.redirect("/?erro=confirmacao");
    }

    if (error || !user) {
      console.log(`❌ Login falhou: erro no signIn - ${error?.message || 'Usuário não encontrado'}`);
      return res.redirect("/?erro=1");
    }

    console.log(`✅ Auth bem-sucedido para: ${email}, procurando perfil...`);

    let profile = await findOrCreateUserProfile(email, user.user_metadata?.usuario || "", senha, user.id);
    if (!profile) {
      console.log(`❌ Login falhou: erro ao buscar/criar perfil para ${email}`);
      return res.redirect("/?erro=1");
    }
    profile = await syncProfileConfirmation(profile, user);

    if (!profile.confirmado) {
      console.log(`❌ Login bloqueado: perfil pendente de confirmação - ${email}`);
      return res.redirect("/?erro=confirmacao");
    }

    console.log(`✅ Perfil encontrado/criado: ${profile.usuario} (${profile.email})`);

    req.session.user = profile.usuario;
    req.session.userId = profile.id;
    console.log(`🔄 Redirecionando para /painel.html`);
    return res.redirect("/painel.html");
  } catch (error) {
    console.log(`❌ Login falhou: erro inesperado - ${error.message}`);
    return res.redirect("/?erro=1");
  }
});

app.post("/login-teste", async (req, res) => {
  const email = normalizeEmail(req.body.email);
  const senha = String(req.body.senha || "").trim();

  console.log(`Tentativa de login teste: ${email}`);

  if (!email || !senha) {
    return res.redirect("/index-teste.html?erro=1");
  }

  try {
    const { user, error, emailNotConfirmed } = await signInOrMigrateLegacyUser(email, senha);
    if (emailNotConfirmed) {
      return res.redirect("/index-teste.html?erro=confirmacao");
    }

    if (error || !user) {
      return res.redirect("/index-teste.html?erro=1");
    }

    let profile = await findOrCreateUserProfile(email, user.user_metadata?.usuario || "", senha, user.id);
    if (!profile) {
      return res.redirect("/index-teste.html?erro=1");
    }
    profile = await syncProfileConfirmation(profile, user);

    if (!profile.confirmado) {
      return res.redirect("/index-teste.html?erro=confirmacao");
    }

    req.session.user = profile.usuario;
    req.session.userId = profile.id;
    console.log("Redirecionando para /painel-teste.html");
    return res.redirect("/painel-teste.html");
  } catch (error) {
    console.log(`Login teste falhou: ${error.message}`);
    return res.redirect("/index-teste.html?erro=1");
  }
});

app.get("/teste", (_req, res) => {
  res.redirect("/index-teste.html");
});

app.post("/cadastro", async (req, res) => {
  const usuario = sanitizeUsuario(req.body.usuario);
  const email = normalizeEmail(req.body.email);
  const senha = String(req.body.senha || "").trim();

  if (!usuario || !email || !senha) {
    return res.redirect("/?cadastro=incompleto");
  }

  if (!isEmailValid(email)) {
    return res.redirect("/?cadastro=dominio");
  }

  if (senha.length < 4) {
    return res.redirect("/?cadastro=senha");
  }

  try {
    if (await getUserProfileByUsuario(usuario)) {
      return res.redirect("/?cadastro=existente");
    }

    if (await getUserProfileByEmail(email)) {
      return res.redirect("/?cadastro=emailexistente");
    }

    const { user, session, error } = await signUpSupabaseAuthUser({ email, senha, usuario }, req);
    if (error) {
      if (isSupabaseDuplicateAuthError(error)) {
        return res.redirect("/?cadastro=emailexistente");
      }
      if (isConfirmationEmailSendError(error)) {
        console.error(`Erro ao enviar e-mail de confirmação pelo Supabase: ${error.message}`);
        return res.redirect("/?cadastro=emailenvio");
      }
      console.error(`Erro no cadastro Supabase: ${error.message}`);
      return res.redirect("/?cadastro=erro");
    }

    const profile = await findOrCreateUserProfile(email, usuario, senha, user?.id, {
      confirmado: Boolean(session) || isAuthUserConfirmed(user)
    });
    if (!profile) {
      return res.redirect("/?cadastro=erro");
    }

    return res.redirect(profile.confirmado ? "/?cadastro=ok" : "/?cadastro=pendente");
  } catch (error) {
    console.error(`Erro ao processar cadastro: ${error.message}`);
    return res.redirect("/?cadastro=erro");
  }
});

app.post("/api/reenviar-confirmacao", async (req, res) => {
  const email = normalizeEmail(req.body.email);

  if (!isEmailValid(email)) {
    return res.status(400).json({ erro: "Informe um e-mail corporativo válido." });
  }

  try {
    const profile = await getUserProfileByEmail(email);
    if (!profile) {
      return res.status(404).json({ erro: "Cadastro não encontrado para este e-mail." });
    }

    if (profile.confirmado) {
      return res.json({ mensagem: "Este e-mail já está confirmado. Você já pode entrar." });
    }

    const { error } = await resendSignupConfirmationEmail(email, req);
    if (error) {
      if (isConfirmationEmailSendError(error)) {
        console.error(`Erro ao reenviar e-mail de confirmação pelo Supabase: ${error.message}`);
        return res.status(500).json({ erro: "Não foi possível enviar o e-mail de confirmação. Verifique a configuração SMTP do Supabase." });
      }
      console.error(`Erro ao reenviar confirmação: ${error.message}`);
      return res.status(500).json({ erro: "Não foi possível reenviar a confirmação agora." });
    }

    res.json({ mensagem: "Link de confirmação reenviado para seu e-mail corporativo." });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao reenviar confirmação." });
  }
});

app.get("/api/me", requireAuth, async (req, res) => {
  try {
    const data = await getCurrentProfile(req);

    if (!data) {
      return res.status(404).json({ erro: "Usuário não encontrado." });
    }

    res.json({ usuario: data.usuario, email: data.email, admin: isAdminProfile(data) });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao buscar dados." });
  }
});

app.get("/api/session", requireAuth, async (req, res) => {
  try {
    const data = await getCurrentProfile(req);
    if (!data) {
      req.session.destroy(() => {});
      return res.status(401).json({ erro: "Sessão inválida." });
    }

    res.json({ usuario: data.usuario, email: data.email, admin: isAdminProfile(data) });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao buscar sessão." });
  }
});

// ============= DATA ROUTES =============

app.get("/api/fornecedores", requireAuth, (_req, res) => {
  try {
    const rows = getCalculationRows();
    if (!rows) {
      return res.json({ fornecedores: [] });
    }

    const fornecedores = [...new Set(rows.map(extractFornecedorNome).filter(Boolean))]
      .sort((a, b) => a.localeCompare(b, "pt-BR"));

    res.json({ fornecedores });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao buscar fornecedores." });
  }
});

app.get("/api/filtros", requireAuth, (req, res) => {
  const fornecedor = normalizeText(req.query.fornecedor);
  const centrosParam = String(req.query.centros || "").trim();

  try {
    const rows = getCalculationRows();
    if (!rows) {
      return res.json({ centros: [], cultivares: [] });
    }

    const filtered = rows.filter(r => {
      const fornecedorMatch = normalizeComparable(extractFornecedorNome(r)) === normalizeComparable(fornecedor);
      if (centrosParam) {
        const centros = centrosParam.split(",").map(c => normalizeComparable(c));
        const centroMatch = centros.includes(normalizeComparable(extractCentro(r)));
        return fornecedorMatch && centroMatch;
      }
      return fornecedorMatch;
    });

    const centros = [...new Set(filtered.map(extractCentro).filter(Boolean))];
    const cultivares = [...new Set(filtered.map(extractCultivar).filter(Boolean))];

    res.json({
      centros: centros.sort((a, b) => a.localeCompare(b, "pt-BR")),
      cultivares: cultivares.sort((a, b) => a.localeCompare(b, "pt-BR"))
    });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao buscar filtros." });
  }
});

app.get("/api/pracas", requireAuth, (_req, res) => {
  try {
    res.json({ pracas: getPracaOptions() });
  } catch (error) {
    console.error(`Erro ao buscar praças: ${error.message}`);
    res.status(500).json({ erro: "Erro ao buscar praças." });
  }
});

function responderRankings(_req, res) {
  try {
    const rankings = buildWorkbookRankings();
    if (!rankings) {
      return res.status(500).json({ erro: "Planilha não encontrada." });
    }

    res.json(rankings);
  } catch (error) {
    console.error(`Erro ao gerar rankings: ${error.message}`);
    res.status(500).json({ erro: "Erro ao gerar rankings." });
  }
}

app.get("/api/rankings", requireAuth, requireAdmin, responderRankings);
app.get("/api/rankings-teste", requireAuth, responderRankings);

app.get("/api/unidades-resumo", requireAuth, (_req, res) => {
  try {
    const resumo = buildUnidadesResumo();
    if (!resumo) {
      return res.status(500).json({ erro: "Planilha não encontrada." });
    }

    res.json(resumo);
  } catch (error) {
    console.error(`Erro ao gerar resumo de unidades: ${error.message}`);
    res.status(500).json({ erro: "Erro ao gerar resumo de unidades." });
  }
});

app.get("/api/unidades-resumo/:unidadeId", requireAuth, (req, res) => {
  try {
    const detalhe = buildUnidadeDetalhe(req.params.unidadeId);
    if (!detalhe) {
      return res.status(500).json({ erro: "Planilha não encontrada." });
    }
    if (!detalhe.unidade) {
      return res.status(404).json({ erro: "Centro não encontrado na planilha atual." });
    }

    res.json(detalhe);
  } catch (error) {
    console.error(`Erro ao gerar detalhe da unidade: ${error.message}`);
    res.status(500).json({ erro: "Erro ao gerar detalhe da unidade." });
  }
});

function getUnidadeRows(unidadeId) {
  const rows = getWorkbookRows();
  if (!rows) return null;
  const unitCache = new Map();
  const result = [];
  const entradaRows = rows
    .filter(isEntradaRow)
    .filter(row => !shouldIgnoreCultivar(extractCultivar(row)));
  for (const row of entradaRows) {
    const centro = extractCentro(row);
    if (!centro) continue;
    const unit = getRankingUnitInfo(centro, unitCache);
    const key = unit.centroCode || normalizeComparable(centro);
    if (!isUnidadeMatch(unit, centro, key, unidadeId)) continue;
    result.push(row);
  }
  return result;
}

app.get("/api/unidades-detalhe/:unidadeId/nfs", requireAuth, (req, res) => {
  try {
    const rows = getUnidadeRows(req.params.unidadeId);
    if (!rows) return res.status(500).json({ erro: "Planilha não encontrada." });
    const nfs = rows.map(row => ({
      nf: extractNf(row) || "-",
      romaneio: extractRomaneio(row) || "-",
      produtor: extractFornecedorNome(row) || "-",
      codigo: extractFornecedorCodigo(row) || "-",
      cultivar: extractCultivar(row) || "-",
      pesoLiquido: extractPesoLiquido(row),
      pesoLiquidoFmt: formatNumberPtBr(extractPesoLiquido(row)),
      pesoBruto: extractPesoBruto(row),
      pesoBrutoFmt: formatNumberPtBr(extractPesoBruto(row))
    }));
    nfs.sort((a, b) => b.pesoLiquido - a.pesoLiquido);
    res.json({ nfs, total: nfs.length });
  } catch (error) {
    console.error(`Erro NFs unidade: ${error.message}`);
    res.status(500).json({ erro: "Erro ao buscar NFs da unidade." });
  }
});

app.get("/api/unidades-detalhe/:unidadeId/produtores", requireAuth, (req, res) => {
  try {
    const rows = getUnidadeRows(req.params.unidadeId);
    if (!rows) return res.status(500).json({ erro: "Planilha não encontrada." });
    const groups = new Map();
    for (const row of rows) {
      const nome = extractFornecedorNome(row) || "-";
      const codigo = extractFornecedorCodigo(row) || "-";
      const cultivar = extractCultivar(row) || "-";
      const nf = extractNf(row);
      const key = normalizeComparable(nome);
      const item = groups.get(key) || { nome, codigo, pesoLiquido: 0, pesoBruto: 0, nfs: new Set(), cultivares: new Set(), entradas: 0 };
      item.pesoLiquido += extractPesoLiquido(row);
      item.pesoBruto += extractPesoBruto(row);
      item.entradas += 1;
      item.cultivares.add(normalizeComparable(cultivar));
      if (nf) item.nfs.add(nf);
      groups.set(key, item);
    }
    const produtores = [...groups.values()].map(item => ({
      nome: item.nome,
      codigo: item.codigo,
      pesoLiquido: item.pesoLiquido,
      pesoLiquidoFmt: formatNumberPtBr(item.pesoLiquido),
      pesoBruto: item.pesoBruto,
      pesoBrutoFmt: formatNumberPtBr(item.pesoBruto),
      nfs: item.nfs.size,
      cultivares: item.cultivares.size,
      entradas: item.entradas
    }));
    produtores.sort((a, b) => b.pesoLiquido - a.pesoLiquido);
    res.json({ produtores, total: produtores.length });
  } catch (error) {
    console.error(`Erro Produtores unidade: ${error.message}`);
    res.status(500).json({ erro: "Erro ao buscar produtores da unidade." });
  }
});

app.get("/api/unidades-detalhe/:unidadeId/cultivares", requireAuth, (req, res) => {
  try {
    const rows = getUnidadeRows(req.params.unidadeId);
    if (!rows) return res.status(500).json({ erro: "Planilha não encontrada." });
    const groups = new Map();
    const totalPesoLiquido = rows.reduce((s, r) => s + extractPesoLiquido(r), 0);
    for (const row of rows) {
      const cultivar = extractCultivar(row) || "-";
      const produtor = extractFornecedorNome(row);
      const nf = extractNf(row);
      const key = normalizeComparable(cultivar);
      const item = groups.get(key) || { cultivar, codigo: extractMaterialCodigo(row) || "-", pesoLiquido: 0, pesoBruto: 0, produtores: new Set(), nfs: new Set(), entradas: 0 };
      item.pesoLiquido += extractPesoLiquido(row);
      item.pesoBruto += extractPesoBruto(row);
      item.entradas += 1;
      if (produtor) item.produtores.add(normalizeComparable(produtor));
      if (nf) item.nfs.add(nf);
      groups.set(key, item);
    }
    const cultivares = [...groups.values()].map(item => ({
      cultivar: item.cultivar,
      codigo: item.codigo,
      pesoLiquido: item.pesoLiquido,
      pesoLiquidoFmt: formatNumberPtBr(item.pesoLiquido),
      pesoBruto: item.pesoBruto,
      pesoBrutoFmt: formatNumberPtBr(item.pesoBruto),
      percentual: totalPesoLiquido > 0 ? (item.pesoLiquido / totalPesoLiquido) * 100 : 0,
      produtores: item.produtores.size,
      nfs: item.nfs.size,
      entradas: item.entradas
    }));
    cultivares.sort((a, b) => b.pesoLiquido - a.pesoLiquido);
    res.json({ cultivares, total: cultivares.length, totalPesoLiquido });
  } catch (error) {
    console.error(`Erro Cultivares unidade: ${error.message}`);
    res.status(500).json({ erro: "Erro ao buscar cultivares da unidade." });
  }
});

app.post("/api/calcular", requireAuth, async (req, res) => {
  const { fornecedorNome, centros, cultivares } = req.body;

  if (!fornecedorNome || !centros?.length || !cultivares?.length) {
    return res.status(400).json({ erro: "Dados incompletos." });
  }

  try {
    const profile = await getCurrentProfile(req);
    if (!profile) {
      return res.status(401).json({ erro: "Sessão inválida." });
    }

    const rows = getCalculationRows();
    if (!rows) {
      return res.status(500).json({ erro: "Planilha não encontrada." });
    }

    const selectedCentros = new Set(centros.map(normalizeComparable));
    const selectedCultivares = new Set(cultivares.map(normalizeComparable));
    const groups = new Map();

    for (const row of rows) {
      const rowFornecedor = extractFornecedorNome(row);
      const fornecedorCodigo = extractFornecedorCodigo(row);
      const centro = extractCentro(row);
      const cultivar = extractCultivar(row);

      if (
        normalizeComparable(rowFornecedor) !== normalizeComparable(fornecedorNome) ||
        !selectedCentros.has(normalizeComparable(centro)) ||
        !selectedCultivares.has(normalizeComparable(cultivar))
      ) {
        continue;
      }

      const inscricaoEstadual = extractInscricao(row);
      const key = `${inscricaoEstadual}|||${centro}|||${cultivar}`;
      const current = groups.get(key) || {
        inscricaoEstadual,
        fornecedorCodigo,
        fornecedorNome: rowFornecedor,
        centro,
        cultivar,
        pesoBase: 0,
        nfs: new Set()
      };

      current.pesoBase += extractPesoBase(row);
      const nf = extractNf(row);
      if (nf) current.nfs.add(nf);
      groups.set(key, current);
    }

    const combinacoes = [...groups.values()].map(item => {
      const defaults = getPracaDefaultsForItem(item);
      return {
        inscricaoEstadual: item.inscricaoEstadual,
        fornecedorCodigo: item.fornecedorCodigo,
        fornecedorNome: item.fornecedorNome,
        centro: item.centro,
        cultivar: item.cultivar,
        pesoBase: item.pesoBase,
        pesoBaseFmt: formatNumberPtBr(item.pesoBase),
        nfs: [...item.nfs].join(", ") || "N/A",
        percPadrao: defaults.perc,
        valorPadrao: defaults.valor,
        praca: defaults.praca
      };
    });

    const pracaPorCentroFornecedor = new Map();
    for (const item of combinacoes) {
      const key = `${item.inscricaoEstadual}|||${item.centro}`;
      if (item.praca && !pracaPorCentroFornecedor.has(key)) {
        pracaPorCentroFornecedor.set(key, item.praca);
      }
    }

    for (const item of combinacoes) {
      const key = `${item.inscricaoEstadual}|||${item.centro}`;
      if (!item.praca && pracaPorCentroFornecedor.has(key)) {
        item.praca = pracaPorCentroFornecedor.get(key);
      }

      const pracaDefaults = getPracaOptionDefaults(item.praca);
      item.percPadrao = Number(item.percPadrao) || pracaDefaults.perc || 0;
      item.valorPadrao = Number(item.valorPadrao) || pracaDefaults.valor || 0;
    }

    const detalhes = combinacoes.map(item => {
      const config = getItemConfig(item, req.body);
      const pesoBonif = item.pesoBase * (config.perc / 100);
      const sacas = pesoBonif / 60;
      const subtotal = sacas * config.valor;

      return {
        ...item,
        praca: config.praca || item.praca,
        perc: formatNumberPtBr(config.perc),
        valor: config.valor,
        valorFmt: formatNumberPtBr(config.valor),
        pesoBonif,
        pesoBonifFmt: formatNumberPtBr(pesoBonif),
        sacas,
        sacasFmt: formatNumberPtBr(sacas),
        subtotal,
        subtotalFmt: formatNumberPtBr(subtotal)
      };
    });

    const totalGeral = detalhes.reduce((sum, item) => sum + item.subtotal, 0);
    const totalGeralFmt = formatNumberPtBr(totalGeral);
    let historicoSalvo = false;
    let historicoErro = "";

    if (req.body.salvarHistorico) {
      try {
        await saveHistoricoEntry({
          usuarioId: profile.id,
          usuario: profile.usuario,
          produtorNome: fornecedorNome,
          totalGeralFmt,
          detalhes
        });
        historicoSalvo = true;
      } catch (error) {
        historicoErro = "Cálculo concluído, mas não foi possível salvar no histórico.";
        console.error(`Erro ao salvar historico do calculo: ${error.message}`);
      }
    }

    res.json({
      combinacoes,
      detalhes,
      totalGeral,
      totalGeralFmt,
      historicoSalvo,
      historicoErro,
      payloadPdf: { usuario: profile.usuario, produtorNome: fornecedorNome, detalhes, totalGeralFmt }
    });
  } catch (error) {
    console.error(`Erro ao calcular: ${error.message}`);
    res.status(500).json({ erro: "Erro ao calcular." });
  }
});

// ============= GERAÇÃO DE ESTRUTURA NFe =============

app.post("/api/nfe-estructura", requireAuth, async (req, res) => {
  const { detalhes, fornecedorNome, totalGeralFmt } = req.body;

  if (!detalhes || !Array.isArray(detalhes) || !fornecedorNome) {
    return res.status(400).json({ erro: "Dados incompletos para gerar estrutura de NFe." });
  }

  try {
    const profile = await getCurrentProfile(req);
    if (!profile) {
      return res.status(401).json({ erro: "Sessão inválida." });
    }

    // Tenta extrair CNPJ da planilha
    let cnpjFornecedor = "";
    const rows = getCalculationRows();
    if (rows && rows.length > 0) {
      // Procura a primeira linha do fornecedor selecionado
      const firstRow = rows.find(r => 
        normalizeComparable(extractFornecedorNome(r)) === normalizeComparable(fornecedorNome)
      );
      if (firstRow) {
        cnpjFornecedor = extractCNPJ(firstRow);
      }
    }

    // Gera estrutura de NFe com dados completos
    const nfeStructure = generateNFeStructure(detalhes, fornecedorNome, totalGeralFmt, cnpjFornecedor);
    
    // Formata para exibição visual
    const nfeDisplay = formatNFeForDisplay(nfeStructure);

    res.json({
      estrutura: nfeStructure,
      visualizacao: nfeDisplay
    });
  } catch (error) {
    console.error(`Erro ao gerar estrutura de NFe: ${error.message}`);
    res.status(500).json({ erro: "Erro ao gerar estrutura de NFe." });
  }
});

// ============= CONFIRMAÇÃO E REPLICAÇÃO DE NFe =============

app.post("/api/confirmar-nfe", requireAuth, async (req, res) => {
  const { estrutura, detalhes, fornecedorNome, totalGeralFmt } = req.body;

  if (!estrutura || !detalhes || !fornecedorNome) {
    return res.status(400).json({ erro: "Dados incompletos para confirmar NFe." });
  }

  try {
    const profile = await getCurrentProfile(req);
    if (!profile) {
      return res.status(401).json({ erro: "Sessão inválida." });
    }

    // Prepara dados da NFe confirmada para armazenamento
    const nfeConfirmada = {
      id: `NFe-${estrutura.numero || Math.floor(Math.random() * 1000000)}`,
      numero: estrutura.numero,
      serie: estrutura.serie,
      fornecedor: fornecedorNome,
      cnpj: estrutura.cnpj,
      dataEmissao: estrutura.dataEmissaoISO,
      cfop: estrutura.cfop,
      valorTotal: estrutura.valorTotalNumerico,
      totalItens: estrutura.totalItens,
      infAdic: estrutura.infAdic,
      itens: detalhes.map(d => ({
        centro: d.centro,
        inscricaoEstadual: d.inscricaoEstadual,
        cultivar: d.cultivar,
        quantidade: d.sacas,
        valor: d.valor,
        subtotal: d.subtotal,
        nfs: d.nfs
      })),
      usuarioConfirmacao: profile.usuario,
      dataConfirmacao: new Date().toISOString(),
      status: "confirmada"
    };

    // Salva no histórico com informações de NFe
    await saveHistoricoEntry({
      usuarioId: profile.id,
      usuario: profile.usuario,
      produtorNome: fornecedorNome,
      totalGeralFmt,
      detalhes,
      nfeConfirmada: nfeConfirmada
    });

    res.json({
      mensagem: "NFe confirmada e registrada com sucesso!",
      nfeId: nfeConfirmada.id,
      nfeNumero: nfeConfirmada.numero,
      nfeSerie: nfeConfirmada.serie,
      status: "confirmada"
    });
  } catch (error) {
    console.error(`Erro ao confirmar NFe: ${error.message}`);
    res.status(500).json({ erro: "Erro ao confirmar NFe." });
  }
});

app.post("/api/replicar-nfe", requireAuth, async (req, res) => {
  const { nfeId, estrutura } = req.body;

  if (!nfeId || !estrutura) {
    return res.status(400).json({ erro: "Dados insuficientes para replicar NFe." });
  }

  try {
    const profile = await getCurrentProfile(req);
    if (!profile) {
      return res.status(401).json({ erro: "Sessão inválida." });
    }

    // Gera nova NFe como cópia da original
    const nfeReplicada = {
      id: `NFe-Replica-${Date.now()}`,
      numeroOriginal: estrutura.numero,
      dataOriginal: estrutura.dataEmissao,
      novoNumero: Math.floor(Math.random() * 1000000).toString().padStart(9, '0'),
      fornecedor: estrutura.fornecedor,
      cnpj: estrutura.cnpj,
      dataReplicacao: new Date().toISOString(),
      cfop: "5101",
      itens: estrutura.itens,
      usuarioReplicacao: profile.usuario,
      status: "replicada"
    };

    // Log da replicação
    console.log(`📋 NFe ${nfeReplicada.numeroOriginal} replicada por ${profile.usuario} como ${nfeReplicada.novoNumero}`);

    res.json({
      mensagem: "NFe replicada com sucesso!",
      nfeId: nfeReplicada.id,
      nfeNumero: nfeReplicada.novoNumero,
      dataReplicacao: nfeReplicada.dataReplicacao,
      status: "replicada"
    });
  } catch (error) {
    console.error(`Erro ao replicar NFe: ${error.message}`);
    res.status(500).json({ erro: "Erro ao replicar NFe." });
  }
});

// ============= HISTÓRICO =============

app.post("/api/historico", requireAuth, async (req, res) => {
  const { totalGeralFmt, detalhes, produtorNome } = req.body;

  try {
    const profile = await getCurrentProfile(req);
    if (!profile) {
      return res.status(401).json({ erro: "Sessão inválida." });
    }

    await saveHistoricoEntry({
      usuarioId: profile.id,
      usuario: profile.usuario,
      produtorNome: produtorNome || "Cálculo de Bonificação",
      totalGeralFmt,
      detalhes
    });

    res.json({ mensagem: "Cálculo salvo no histórico." });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao salvar histórico." });
  }
});

app.get("/api/historico", requireAuth, async (req, res) => {
  try {
    const profile = await getCurrentProfile(req);
    if (!profile) {
      return res.status(401).json({ erro: "Sessão inválida." });
    }

    let query = supabaseAdmin
      .from("historico")
      .select("*")
      .order("createdat", { ascending: false })
      .limit(20);

    const isAdmin = isAdminProfile(profile);
    if (!isAdmin) {
      query = query.eq("usuario_id", profile.id);
    }

    const { data, error } = await query;
    if (error) throw error;

    res.json({ historico: (data || []).map(normalizeHistoricoRow) });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao buscar histórico." });
  }
});

app.delete("/api/historico", requireAuth, async (req, res) => {
  try {
    const profile = await getCurrentProfile(req);
    if (!profile) {
      return res.status(401).json({ erro: "Sessão inválida." });
    }

    let query = supabaseAdmin.from("historico").delete();
    const isAdmin = isAdminProfile(profile);
    if (!isAdmin) {
      query = query.eq("usuario_id", profile.id);
    } else {
      query = query.neq("id", "00000000-0000-0000-0000-000000000000");
    }

    const { error } = await query;
    if (error) throw error;

    res.json({ mensagem: "Histórico limpo com sucesso.", historico: [] });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao limpar histórico." });
  }
});

// ============= PDF =============

let unitsRowsCache = null;

function getUnitsRows() {
  if (!fs.existsSync(UNITS_WORKBOOK)) return [];

  try {
    const stat = fs.statSync(UNITS_WORKBOOK);
    const cacheKey = `${UNITS_WORKBOOK}|${stat.mtimeMs}|${stat.size}`;
    if (unitsRowsCache?.cacheKey === cacheKey) return unitsRowsCache.rows;

    const workbook = XLSX.readFile(UNITS_WORKBOOK);
    const worksheet = workbook.Sheets.Unidades || workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    unitsRowsCache = { cacheKey, rows };
    return rows;
  } catch (error) {
    console.error(`Erro ao ler unidades: ${error.message}`);
    unitsRowsCache = null;
    return [];
  }
}

function extractCentroCode(value) {
  const match = String(value || "").toUpperCase().match(/\bBS\s*0?\d{1,2}\b/);
  if (!match) return normalizeComparable(value);
  const digits = match[0].replace(/\D/g, "").padStart(2, "0").slice(-2);
  return `BS${digits}`;
}

function stripCepFromAddress(endereco) {
  return normalizeText(endereco)
    .replace(/\s*,?\s*CEP\s*\d{5}-?\d{3}\s*,?/i, ", ")
    .replace(/\s*,\s*,\s*/g, ", ")
    .replace(/\s{2,}/g, " ")
    .replace(/^\s*,\s*|\s*,\s*$/g, "")
    .trim();
}

function findUnitByCentro(centro) {
  const centroCode = extractCentroCode(centro);
  const unit = getUnitsRows().find(row => extractCentroCode(getRowValue(row, ["CODIGO", "Centro", "Código"])) === centroCode);

  if (!unit) {
    return {
      razaoSocial: "BOA SAFRA SEMENTES S/A",
      cnpj: "",
      inscricaoEstadual: "",
      endereco: "",
      cidade: "",
      uf: "",
      cep: "",
      centro
    };
  }

  return {
    razaoSocial: normalizeText(getRowValue(unit, ["RAZÃO SOCIAL", "Razão Social", "Razao Social"])) || "BOA SAFRA SEMENTES S/A",
    cnpj: normalizeText(getRowValue(unit, ["CNPJ"])),
    inscricaoEstadual: normalizeText(getRowValue(unit, ["ISC. ESTADUAL", "INSCRIÇÃO ESTADUAL", "Inscrição Estadual"])),
    endereco: stripCepFromAddress(getRowValue(unit, ["ENDEREÇO", "Endereco", "Endereço"])),
    cidade: normalizeText(getRowValue(unit, ["CIDADE", "Cidade"])),
    uf: normalizeText(getRowValue(unit, ["UF"])),
    cep: normalizeText(getRowValue(unit, ["CEP"])),
    centro: normalizeText(getRowValue(unit, ["CODIGO", "Centro"])) || centro
  };
}

function parsePtBrNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const text = String(value || "0").replace(/[^\d,.-]/g, "");
  if (!text) return 0;
  if (text.includes(",")) return Number(text.replace(/\./g, "").replace(",", ".")) || 0;
  return Number(text) || 0;
}

function formatCurrencyPdf(value) {
  return `R$ ${formatNumberPtBr(value)}`;
}

function formatNumberPtBrFixed(value, digits = 2) {
  const num = Number(value) || 0;
  return num.toLocaleString("pt-BR", { minimumFractionDigits: digits, maximumFractionDigits: digits });
}

function getItemSubtotal(item) {
  return parsePtBrNumber(item.subtotal ?? item.subtotalFmt);
}

function getItemQuantity(item) {
  return parsePtBrNumber(item.sacas ?? item.sacasFmt);
}

function getItemPesoLiquido(item) {
  return parsePtBrNumber(item.pesoBase ?? item.pesoBaseFmt);
}

function getItemPesoBonificado(item) {
  return parsePtBrNumber(item.pesoBonif ?? item.pesoBonifFmt);
}

function formatItemPesoValorResumo(item) {
  return `Peso Líquido: ${formatNumberPtBr(getItemPesoLiquido(item))} kg | Peso Bonificado: ${formatNumberPtBr(getItemPesoBonificado(item))} kg | Valor: ${formatCurrencyPdf(getItemSubtotal(item))}`;
}

function splitUniqueList(value) {
  return [...new Set(String(value || "")
    .split(",")
    .map(item => item.trim())
    .filter(Boolean))];
}

function getNfePracasResumo(nfe) {
  const pracas = [...new Set((nfe.itens || [])
    .map(item => normalizeText(item.praca))
    .filter(Boolean))];

  return pracas.join(", ");
}

function fitPdfFontSize(doc, text, width, { font = "Helvetica", maxSize = 12, minSize = 8.5, step = 0.5 } = {}) {
  for (let size = maxSize; size >= minSize; size -= step) {
    doc.font(font).fontSize(size);
    if (doc.widthOfString(text) <= width) return size;
  }

  return minSize;
}

function buildBonificacaoNfes(detalhes) {
  const nfesPorGrupo = new Map();

  detalhes.forEach(item => {
    const chave = `${item.inscricaoEstadual}|||${item.centro}`;
    if (!nfesPorGrupo.has(chave)) {
      nfesPorGrupo.set(chave, {
        inscricaoEstadual: item.inscricaoEstadual,
        centro: item.centro,
        destinatario: findUnitByCentro(item.centro),
        itens: []
      });
    }
    nfesPorGrupo.get(chave).itens.push(item);
  });

  return Array.from(nfesPorGrupo.values())
    .sort((a, b) => {
      const ieCompare = String(a.inscricaoEstadual || "").localeCompare(String(b.inscricaoEstadual || ""), "pt-BR");
      if (ieCompare !== 0) return ieCompare;
      return String(a.centro || "").localeCompare(String(b.centro || ""), "pt-BR");
    })
    .map((nfe, index) => ({
      ...nfe,
      numero: index + 1,
      total: nfe.itens.reduce((sum, item) => sum + getItemSubtotal(item), 0),
      nfsReferenciadas: splitUniqueList(nfe.itens.map(item => item.nfs || "").join(", "))
    }));
}

function ensurePdfSpace(doc, heightNeeded, onNewPage) {
  const bottomLimit = doc.page.height - Math.max(doc.page.margins.bottom, REPORT_BOTTOM_MARGIN);
  if (doc.y + heightNeeded <= bottomLimit) return;
  doc.addPage();
  if (onNewPage) onNewPage();
}

function drawReportLetterhead(doc) {
  if (!fs.existsSync(LETTERHEAD_IMAGE)) return;
  doc.image(LETTERHEAD_IMAGE, 0, 0, {
    width: doc.page.width,
    height: doc.page.height
  });
  doc.y = REPORT_CONTENT_TOP;
}

function drawResumoHeader(doc, { produtorNome, usuario, nfesCount }) {
  drawReportLetterhead(doc);
  doc.font("Helvetica-Bold").fontSize(18).fillColor("#2f6b37")
    .text("RELATÓRIO DE BONIFICAÇÃO", { align: "center" });
  doc.moveDown(0.4);
  doc.strokeColor("#cde5bd").lineWidth(1).moveTo(35, doc.y).lineTo(560, doc.y).stroke();
  doc.moveDown(0.5);

  doc.font("Helvetica").fontSize(10).fillColor("#333333");
  doc.text(`Produtor: ${produtorNome || "-"}`);
  doc.text(`Usuário: ${usuario || "-"}`);
  doc.text(`Data Emissão: ${new Date().toLocaleString("pt-BR")}`);
  doc.text(`Bonificações geradas: ${nfesCount}`);
  doc.moveDown(0.6);
}

function drawResumoBonificacao(doc, nfe, renderHeader) {
  const destinatario = nfe.destinatario;
  const cardX = 35;
  const cardWidth = 525;
  const padding = 9;
  const contentWidth = cardWidth - padding * 2;
  const centroLabel = extractCentroCode(nfe.centro) || nfe.centro || "-";
  const pracaLabel = getNfePracasResumo(nfe);
  const titulo = `BONIFICAÇÃO ${nfe.numero} - Inscrição Estadual: ${nfe.inscricaoEstadual || "-"} – Centro ${centroLabel}${pracaLabel ? ` | Pra\u00e7a: ${pracaLabel}` : ""}`;
  const totalLabel = `Valor Total NF: ${formatCurrencyPdf(nfe.total)}`;
  const destinoTexto = ` Destinatário: ${destinatario.razaoSocial || "-"} | CNPJ: ${destinatario.cnpj || "-"}`;
  const linhasItens = nfe.itens.map(item =>
    `Cultivar: ${item.cultivar || "-"} | Notas: ${item.nfs || "N/A"} | ${formatItemPesoValorResumo(item)}`
  );

  const tituloFontSize = fitPdfFontSize(doc, titulo, contentWidth, { font: "Helvetica-Bold", maxSize: 12, minSize: 8.5 });
  doc.font("Helvetica-Bold").fontSize(tituloFontSize);
  const titleHeight = Math.max(29, doc.heightOfString(titulo, { width: contentWidth, align: "center" }) + 13);
  doc.font("Helvetica").fontSize(11);
  const infoHeight = doc.heightOfString(`${totalLabel}${destinoTexto}`, { width: contentWidth });
  const itemsHeight = linhasItens.reduce((sum, linha) => {
    doc.font("Helvetica").fontSize(10.5);
    return sum + doc.heightOfString(linha, { width: contentWidth }) + 2;
  }, 0);
  const cardHeight = titleHeight + padding + infoHeight + 17 + itemsHeight + padding;

  ensurePdfSpace(doc, cardHeight + 8, renderHeader);

  const cardY = doc.y;
  doc.lineWidth(0.8).strokeColor("#111111").roundedRect(cardX, cardY, cardWidth, cardHeight, 10).stroke();
  doc.moveTo(cardX, cardY + titleHeight).lineTo(cardX + cardWidth, cardY + titleHeight).stroke();

  doc.font("Helvetica-Bold").fontSize(tituloFontSize).fillColor("#111111")
    .text(titulo, cardX + padding, cardY + 8, { width: contentWidth, align: "center" });

  let currentY = cardY + titleHeight + padding;
  doc.font("Helvetica-Bold").fontSize(11).fillColor("#111111")
    .text(totalLabel, cardX + padding, currentY, { width: contentWidth, continued: true });
  doc.font("Helvetica").fontSize(11).fillColor("#111111")
    .text(destinoTexto, { width: contentWidth });

  currentY = doc.y + 14;
  linhasItens.forEach(linha => {
    doc.font("Helvetica").fontSize(10.5).fillColor("#111111")
      .text(linha, cardX + padding, currentY, { width: contentWidth });
    currentY = doc.y + 2;
  });

  doc.y = cardY + cardHeight + 8;
}

function drawResumoBonificacoes(doc, { usuario, produtorNome, totalGeralFmt, nfes }) {
  const renderHeader = () => drawResumoHeader(doc, { produtorNome, usuario, nfesCount: nfes.length });
  renderHeader();

  let lastIe = null;
  nfes.forEach(nfe => {
    if (lastIe === nfe.inscricaoEstadual) {
      ensurePdfSpace(doc, 58, renderHeader);
    }
    drawResumoBonificacao(doc, nfe, renderHeader);
    lastIe = nfe.inscricaoEstadual;
  });

  ensurePdfSpace(doc, 36, renderHeader);
  doc.moveDown(0.4);
  doc.font("Helvetica-Bold").fontSize(13).fillColor("#2f6b37")
    .text(`VALOR TOTAL GERAL: R$ ${totalGeralFmt || formatNumberPtBr(nfes.reduce((sum, nfe) => sum + nfe.total, 0))}`, { align: "right" });
}

function drawFieldBox(doc, x, y, w, h, label, value, options = {}) {
  doc.lineWidth(0.5).strokeColor("#222222").rect(x, y, w, h).stroke();
  doc.font("Helvetica-Bold").fontSize(options.labelSize || 5.5).fillColor("#222222")
    .text(label, x + 2, y + 2, { width: w - 4, height: 8 });
  if (value !== undefined && value !== null && value !== "") {
    doc.font(options.bold ? "Helvetica-Bold" : "Helvetica").fontSize(options.valueSize || 7).fillColor("#111111")
      .text(String(value), x + 2, y + 11, { width: w - 4, height: h - 13, align: options.align || "left" });
  }
}

function makeSimulatedAccessKey(nfe, produtorNome) {
  const seed = crypto
    .createHash("sha256")
    .update(`${produtorNome}|${nfe.inscricaoEstadual}|${nfe.centro}|${nfe.numero}|${nfe.total}`)
    .digest("hex");
  return Array.from(seed)
    .map(char => (parseInt(char, 16) % 10).toString())
    .join("")
    .padEnd(44, "0")
    .slice(0, 44);
}

function formatAccessKey(key) {
  return String(key || "").replace(/(.{4})/g, "$1 ").trim();
}

function drawSimulatedBarcode(doc, x, y, w, h, key) {
  doc.lineWidth(0.5).strokeColor("#222222").rect(x, y, w, h).stroke();
  const digits = String(key || "").padEnd(44, "0");
  let cursor = x + 5;
  for (let i = 0; i < digits.length && cursor < x + w - 4; i++) {
    const digit = Number(digits[i]) || 0;
    const lineWidth = digit % 3 === 0 ? 1.4 : digit % 2 === 0 ? 0.8 : 0.4;
    doc.lineWidth(lineWidth).moveTo(cursor, y + 4).lineTo(cursor, y + h - 4).stroke();
    cursor += digit % 2 === 0 ? 3 : 2;
  }
}

function drawDanfeProducts(doc, nfe, x, y, w) {
  const columns = [
    { label: "CÓDIGO", width: 38 },
    { label: "DESCRIÇÃO DO PRODUTO / SERVIÇO", width: 168 },
    { label: "NCM/SH", width: 44 },
    { label: "O/CST", width: 32 },
    { label: "CFOP", width: 30 },
    { label: "UN", width: 22 },
    { label: "QUANT", width: 42 },
    { label: "VALOR UNIT", width: 52 },
    { label: "VALOR TOTAL", width: 54 },
    { label: "B.CÁLC ICMS", width: 42 },
    { label: "VALOR ICMS", width: 38 },
  ];

  const headerHeight = 18;
  doc.font("Helvetica-Bold").fontSize(5).fillColor("#222222");
  let cursor = x;
  columns.forEach(col => {
    doc.lineWidth(0.5).strokeColor("#222222").rect(cursor, y, col.width, headerHeight).stroke();
    doc.text(col.label, cursor + 2, y + 3, { width: col.width - 4, align: "center" });
    cursor += col.width;
  });

  let rowY = y + headerHeight;
  const maxBottom = 640;
  const visibleItems = [];
  let hiddenItems = 0;

  for (const item of nfe.itens) {
    const descriptionHeight = doc.heightOfString(item.cultivar || "-", { width: columns[1].width - 4 });
    const rowHeight = Math.max(18, Math.min(34, descriptionHeight + 8));
    if (rowY + rowHeight > maxBottom) {
      hiddenItems++;
      continue;
    }
    visibleItems.push({ item, rowHeight });
    rowY += rowHeight;
  }

  rowY = y + headerHeight;
  visibleItems.forEach(({ item, rowHeight }, idx) => {
    const quantity = 1;
    const subtotal = getItemSubtotal(item);
    const valorUnitario = subtotal;
    const values = [
      String(idx + 1).padStart(2, "0"),
      item.cultivar || "-",
      "12019000",
      "040",
      "5101",
      "KG",
      formatNumberPtBrFixed(quantity, 4),
      formatNumberPtBrFixed(valorUnitario, 4),
      formatNumberPtBr(subtotal),
      "0,00",
      "0,00"
    ];

    cursor = x;
    columns.forEach((col, colIndex) => {
      doc.lineWidth(0.5).strokeColor("#222222").rect(cursor, rowY, col.width, rowHeight).stroke();
      doc.font("Helvetica").fontSize(colIndex === 1 ? 5.5 : 5).fillColor("#111111")
        .text(values[colIndex], cursor + 2, rowY + 4, {
          width: col.width - 4,
          height: rowHeight - 6,
          align: colIndex >= 6 ? "right" : "left"
        });
      cursor += col.width;
    });
    rowY += rowHeight;
  });

  if (hiddenItems > 0) {
    doc.font("Helvetica-Bold").fontSize(6).fillColor("#111111")
      .text(`+ ${hiddenItems} item(ns) detalhado(s) nas informações complementares.`, x + 2, rowY + 3, { width: w - 4 });
  }
}

function drawSimulatedDanfe(doc, nfe, { produtorNome }) {
  const margin = 18;
  const width = 559;
  const today = new Date();
  const dataEmissao = today.toLocaleDateString("pt-BR");
  const horaEmissao = today.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
  const accessKey = makeSimulatedAccessKey(nfe, produtorNome);
  const destinatario = nfe.destinatario;
  const nfNumero = String(nfe.numero).padStart(9, "0");
  const totalFmt = formatNumberPtBr(nfe.total);
  const cultivares = nfe.itens.map(item => item.cultivar).filter(Boolean).join("; ");
  const notas = nfe.nfsReferenciadas.join(", ") || "N/A";
  const pesoLiquidoTotal = nfe.itens.reduce((sum, item) => sum + getItemPesoLiquido(item), 0);
  const pesoBonificadoTotal = nfe.itens.reduce((sum, item) => sum + getItemPesoBonificado(item), 0);

  doc.font("Helvetica").fontSize(6).fillColor("#555555")
    .text(`Impresso em ${today.toLocaleString("pt-BR")} - DANFE simulado pelo Sistema de Bonificação Boa Safra`, margin, 12, { width });
  doc.font("Helvetica-Bold").fontSize(8).fillColor("#9a3412")
    .text("SIMULAÇÃO DE DANFE - SEM VALIDADE FISCAL", margin, 23, { width, align: "center" });

  drawFieldBox(doc, margin, 38, 150, 42, "DATA DE RECEBIMENTO", "");
  drawFieldBox(doc, margin + 150, 38, 409, 42, "IDENTIFICAÇÃO E ASSINATURA DO RECEBEDOR", "");

  doc.lineWidth(0.5).strokeColor("#222222").rect(margin, 84, 260, 82).stroke();
  doc.font("Helvetica-Bold").fontSize(5.5).fillColor("#222222")
    .text("IDENTIFICAÇÃO DO EMITENTE", margin + 2, 86, { width: 256, height: 8 });
  doc.font("Helvetica").fontSize(7).fillColor("#111111")
    .text(`${produtorNome || "-"}\nIE: ${nfe.inscricaoEstadual || "-"}\nCNPJ/CPF: não informado`, margin + 2, 95, { width: 256, height: 32 });
  doc.font("Helvetica-Bold").fontSize(20).fillColor("#000000")
    .text(`BONIFICAÇÃO ${nfe.numero}`, margin + 2, 130, { width: 256, align: "center" });
  doc.lineWidth(0.5).strokeColor("#222222").rect(margin + 260, 84, 105, 82).stroke();
  doc.font("Helvetica-Bold").fontSize(17).fillColor("#111111").text("DANFE", margin + 260, 92, { width: 105, align: "center" });
  doc.font("Helvetica").fontSize(7).text("Documento Auxiliar da Nota Fiscal Eletrônica", margin + 268, 114, { width: 89, align: "center" });
  doc.font("Helvetica-Bold").fontSize(8).text("0 - ENTRADA", margin + 276, 138);
  doc.text("1 - SAÍDA", margin + 276, 151);
  doc.rect(margin + 345, 144, 12, 12).stroke();
  doc.font("Helvetica-Bold").fontSize(8).text("1", margin + 349, 146);

  drawFieldBox(doc, margin + 365, 84, 212, 28, "CONTROLE DO FISCO", "SIMULADO", { align: "center", bold: true, valueSize: 8 });
  drawSimulatedBarcode(doc, margin + 365, 112, 212, 28, accessKey);
  drawFieldBox(doc, margin + 365, 140, 212, 26, "CHAVE DE ACESSO", formatAccessKey(accessKey), { valueSize: 6, align: "center" });

  drawFieldBox(doc, margin, 166, 260, 30, "NATUREZA DA OPERAÇÃO", "BONIFICAÇÃO");
  drawFieldBox(doc, margin + 260, 166, 105, 30, "NF-e", `Nº ${nfNumero}\nSérie 001`, { bold: true, align: "center", valueSize: 7 });
  drawFieldBox(doc, margin + 365, 166, 212, 30, "PROTOCOLO DE AUTORIZAÇÃO DE USO", "SIMULADO - SEM AUTORIZAÇÃO FISCAL", { valueSize: 6 });

  doc.font("Helvetica-Bold").fontSize(7).fillColor("#111111").text("DESTINATÁRIO / REMETENTE", margin, 205);
  drawFieldBox(doc, margin, 216, 290, 34, "NOME / RAZÃO SOCIAL", destinatario.razaoSocial || "-");
  drawFieldBox(doc, margin + 290, 216, 110, 34, "CNPJ / CPF", destinatario.cnpj || "-");
  drawFieldBox(doc, margin + 400, 216, 80, 34, "DATA DA EMISSÃO", dataEmissao, { align: "center" });
  drawFieldBox(doc, margin + 480, 216, 79, 34, "Nº DA NF-e", nfNumero, { align: "center", bold: true });
  drawFieldBox(doc, margin, 250, 290, 34, "ENDEREÇO", destinatario.endereco || "-");
  drawFieldBox(doc, margin + 290, 250, 110, 34, "BAIRRO / DISTRITO", "-");
  drawFieldBox(doc, margin + 400, 250, 80, 34, "CEP", destinatario.cep || "-");
  drawFieldBox(doc, margin + 480, 250, 79, 34, "DATA SAÍDA/ENTRADA", dataEmissao, { align: "center" });
  drawFieldBox(doc, margin, 284, 165, 34, "MUNICÍPIO", destinatario.cidade || "-");
  drawFieldBox(doc, margin + 165, 284, 50, 34, "UF", destinatario.uf || "-");
  drawFieldBox(doc, margin + 215, 284, 185, 34, "INSCRIÇÃO ESTADUAL", destinatario.inscricaoEstadual || "-");
  drawFieldBox(doc, margin + 400, 284, 80, 34, "CENTRO", destinatario.centro || nfe.centro || "-");
  drawFieldBox(doc, margin + 480, 284, 79, 34, "HORA SAÍDA/ENTRADA", horaEmissao, { align: "center" });

  doc.font("Helvetica-Bold").fontSize(7).fillColor("#111111").text("PAGAMENTO", margin, 328);
  drawFieldBox(doc, margin, 338, 180, 30, "FORMA", "Sem pagamento");
  drawFieldBox(doc, margin + 180, 338, 110, 30, "VALOR", formatCurrencyPdf(nfe.total), { bold: true });

  doc.font("Helvetica-Bold").fontSize(7).fillColor("#111111").text("CÁLCULO DO IMPOSTO", margin, 378);
  const calcY = 388;
  const calcW = width / 6;
  [
    ["BASE DE CÁLC. DO ICMS", "0,00"],
    ["VALOR DO ICMS", "0,00"],
    ["BASE CÁLC. ICMS S.T.", "0,00"],
    ["VALOR ICMS SUBST.", "0,00"],
    ["V. TOTAL PRODUTOS", totalFmt],
    ["V. TOTAL DA NOTA", totalFmt]
  ].forEach(([label, value], index) => {
    drawFieldBox(doc, margin + calcW * index, calcY, calcW, 34, label, value, { align: "right", bold: index >= 4 });
  });

  doc.font("Helvetica-Bold").fontSize(7).fillColor("#111111").text("TRANSPORTADOR / VOLUMES TRANSPORTADOS", margin, 432);
  drawFieldBox(doc, margin, 442, 220, 30, "NOME / RAZÃO SOCIAL", "");
  drawFieldBox(doc, margin + 220, 442, 70, 30, "FRETE", "9-Sem Transporte");
  drawFieldBox(doc, margin + 290, 442, 80, 30, "PLACA DO VEÍCULO", "");
  drawFieldBox(doc, margin + 370, 442, 45, 30, "UF", "");
  drawFieldBox(doc, margin + 415, 442, 144, 30, "CNPJ / CPF", "");

  doc.font("Helvetica-Bold").fontSize(7).fillColor("#111111").text("DADOS DOS PRODUTOS / SERVIÇOS", margin, 484);
  drawDanfeProducts(doc, nfe, margin, 494, width);

  doc.font("Helvetica-Bold").fontSize(7).fillColor("#111111").text("DADOS ADICIONAIS", margin, 666);
  drawFieldBox(doc, margin, 676, 385, 128, "INFORMAÇÕES COMPLEMENTARES", "", { valueSize: 6 });
  const complemento = `COMPLEMENTO DE VALOR REFERENTE BONIFICAÇÃO DAS SEGUINTES CULTIVARES: ${cultivares || "-"}.\nNOTAS FISCAIS REFERENCIADAS: ${notas}.\nPESO LÍQUIDO: ${formatNumberPtBr(pesoLiquidoTotal)} KG | PESO BONIFICADO: ${formatNumberPtBr(pesoBonificadoTotal)} KG | VALOR: ${formatCurrencyPdf(nfe.total)}.\nOperação simulada com CFOP 5101. Documento sem validade fiscal; use apenas para conferência interna antes da emissão oficial.`;
  doc.font("Helvetica").fontSize(6).fillColor("#111111").text(complemento, margin + 3, 690, { width: 379, height: 105 });
  drawFieldBox(doc, margin + 385, 676, 174, 128, "RESERVADO AO FISCO", "SIMULAÇÃO\nSEM VALIDADE FISCAL", { align: "center", bold: true, valueSize: 11 });
}

app.post("/api/relatorio", requireAuth, (req, res) => {
  const { usuario, produtorNome, totalGeralFmt } = req.body;
  const detalhes = Array.isArray(req.body.detalhes) ? req.body.detalhes : [];

  if (detalhes.length === 0) {
    return res.status(400).json({ erro: "Nenhum detalhe para gerar NFes." });
  }

  const doc = new PDFKitDocument({ size: "A4", margin: 35 });
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

  const nfes = buildBonificacaoNfes(detalhes);

  drawResumoBonificacoes(doc, { usuario, produtorNome, totalGeralFmt, nfes });

  nfes.forEach((nfe, index) => {
    doc.addPage();
    drawSimulatedDanfe(doc, nfe, { produtorNome });
  });

  doc.end();
});

// ============= UPLOAD =============

app.post("/api/planilha", requireAuth, requireAdmin, upload.single("planilha"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ erro: "Nenhum arquivo enviado." });
  }

  try {
    const profile = await getCurrentProfile(req);
    const savedWorkbook = saveCurrentUploadedWorkbook(req.file.buffer, req.file.originalname);
    const meta = {
      storedName: savedWorkbook.currentName,
      currentName: savedWorkbook.currentName,
      ext: savedWorkbook.ext,
      size: savedWorkbook.size,
      sha256: savedWorkbook.sha256,
      originalName: req.file.originalname,
      uploadedAt: new Date().toISOString(),
      uploadedBy: profile?.usuario || req.session.user || null,
      uploadedByEmail: profile?.email || null
    };

    writeWorkbookMeta(meta);
    await saveWorkbookSnapshot(meta, req.file.buffer);
    clearWorkbookCache();

    res.json({ mensagem: "Planilha atualizada com sucesso.", planilha: getWorkbookInfo() });
  } catch (error) {
    console.error(`Erro ao fazer upload da planilha: ${error.message}`);
    res.status(500).json({ erro: "Erro ao fazer upload." });
  }
});

app.get("/api/planilha", requireAuth, requireAdmin, (_req, res) => {
  const info = getWorkbookInfo();
  res.json(info);
});

// ============= CONFIGURACOES E USUARIOS =============

app.get("/api/dashboard-link", requireAuth, async (_req, res) => {
  try {
    const link = await getConfigValue("dashboard_link") || getLocalDashboardLink();
    res.json({ link });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao buscar link do dashboard." });
  }
});

app.put("/api/dashboard-link", requireAuth, requireAdmin, async (req, res) => {
  const link = String(req.body.link || "").trim();

  try {
    await setConfigValue("dashboard_link", link);
    res.json({ mensagem: "Link do dashboard atualizado.", link });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao salvar link do dashboard." });
  }
});

app.get("/api/usuarios", requireAuth, requireAdmin, async (_req, res) => {
  try {
    res.json({ usuarios: await listUserProfiles() });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao listar usuários." });
  }
});

app.post("/api/usuarios", requireAuth, requireAdmin, async (req, res) => {
  const usuario = sanitizeUsuario(req.body.usuario);
  const email = normalizeEmail(req.body.email);
  const senha = String(req.body.senha || "").trim();
  const admin = normalizeBoolean(req.body.admin);

  if (!usuario || !email || !senha) {
    return res.status(400).json({ erro: "Informe usuário, e-mail e senha." });
  }

  if (!isEmailValid(email)) {
    return res.status(400).json({ erro: "E-mail deve ser @boasafrasementes.com.br" });
  }

  if (senha.length < 4) {
    return res.status(400).json({ erro: "A senha deve ter no mínimo 4 caracteres." });
  }

  try {
    const existingByUsuario = await getUserProfileByUsuario(usuario);
    if (existingByUsuario) {
      return res.status(409).json({ erro: "Usuário já existe." });
    }

    const existingByEmail = await supabaseAdmin
      .from("usuarios")
      .select("id")
      .eq("email", email)
      .maybeSingle();

    if (existingByEmail.error) throw existingByEmail.error;

    if (existingByEmail.data) {
      return res.status(409).json({ erro: "E-mail já existe." });
    }

    const { user, error: authError } = await createSupabaseAuthUser({ email, senha, usuario });
    if (authError) {
      if (isSupabaseDuplicateAuthError(authError)) {
        return res.status(409).json({ erro: "E-mail já existe no Supabase Auth." });
      }
      return res.status(500).json({ erro: "Erro ao criar usuário no Supabase Auth." });
    }

    const profile = await findOrCreateUserProfile(email, usuario, senha, user?.id, { admin, confirmado: true });
    if (!profile) {
      return res.status(500).json({ erro: "Erro ao criar perfil do usuário." });
    }

    res.json({ mensagem: "Usuário criado com sucesso.", usuarios: await listUserProfiles() });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao criar usuário." });
  }
});

app.put("/api/usuarios/:usuario/senha", requireAuth, async (req, res) => {
  const usuario = sanitizeUsuario(req.params.usuario);
  const senha = String(req.body.senha || "").trim();

  if (senha.length < 4) {
    return res.status(400).json({ erro: "A senha deve ter no mínimo 4 caracteres." });
  }

  try {
    const currentProfile = await getCurrentProfile(req);
    if (!currentProfile) {
      return res.status(401).json({ erro: "Sessão inválida." });
    }

    if (currentProfile.usuario !== usuario && !isAdminProfile(currentProfile)) {
      return res.status(403).json({ erro: "Sem permissão para alterar esta senha." });
    }

    const profile = await getUserProfileByUsuario(usuario);
    if (!profile) {
      return res.status(404).json({ erro: "Usuário não encontrado." });
    }

    let authUser = await getAuthUserForProfile(profile);
    if (!authUser && isEmailValid(profile.email)) {
      const created = await createSupabaseAuthUser({ email: profile.email, senha, usuario: profile.usuario });
      if (created.error && !isSupabaseDuplicateAuthError(created.error)) {
        return res.status(500).json({ erro: "Erro ao sincronizar usuário no Supabase Auth." });
      }
      authUser = created.user || await getAuthUserByEmail(profile.email);
    }

    if (authUser) {
      const { error: authError } = await supabaseAdmin.auth.admin.updateUserById(authUser.id, { password: senha });
      if (authError) throw authError;
    }

    const { error } = await supabaseAdmin
      .from("usuarios")
      .update({ senha })
      .eq("usuario", usuario);

    if (error) throw error;

    res.json({ mensagem: "Senha atualizada com sucesso." });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao alterar senha." });
  }
});

app.put("/api/usuarios/:usuario/admin", requireAuth, requireAdmin, async (req, res) => {
  const usuario = sanitizeUsuario(req.params.usuario);
  const email = normalizeEmail(req.body.email);
  const admin = normalizeBoolean(req.body.admin);

  try {
    const profile = await getUserProfileByUsuario(usuario);
    if (!profile) {
      return res.status(404).json({ erro: "Usuário não encontrado." });
    }

    const updates = { admin };
    if (email) {
      if (!isEmailValid(email)) {
        return res.status(400).json({ erro: "E-mail deve ser @boasafrasementes.com.br" });
      }
      updates.email = email;
    }

    const authUser = await getAuthUserForProfile({ ...profile, email: email || profile.email });
    if (authUser) {
      const authUpdates = { user_metadata: { usuario } };
      if (email && normalizeEmail(authUser.email) !== email) authUpdates.email = email;
      const { error: authError } = await supabaseAdmin.auth.admin.updateUserById(authUser.id, authUpdates);
      if (authError) throw authError;
    }

    const { error } = await supabaseAdmin
      .from("usuarios")
      .update(updates)
      .eq("usuario", usuario);

    if (error) throw error;

    res.json({
      mensagem: admin ? "Usuário promovido a administrador." : "Permissão de administrador removida.",
      usuarios: await listUserProfiles()
    });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao atualizar permissão de administrador." });
  }
});

app.delete("/api/usuarios/:usuario", requireAuth, requireAdmin, async (req, res) => {
  const usuario = sanitizeUsuario(req.params.usuario);

  try {
    const currentProfile = await getCurrentProfile(req);
    if (currentProfile?.usuario === usuario) {
      return res.status(400).json({ erro: "Não é possível excluir o usuário logado." });
    }

    const profile = await getUserProfileByUsuario(usuario);
    if (!profile) {
      return res.status(404).json({ erro: "Usuário não encontrado." });
    }

    const authUser = await getAuthUserForProfile(profile);
    if (authUser) {
      await supabaseAdmin.auth.admin.deleteUser(authUser.id);
    }

    const { error } = await supabaseAdmin
      .from("usuarios")
      .delete()
      .eq("usuario", usuario);

    if (error) throw error;

    res.json({ mensagem: "Usuário excluído com sucesso.", usuarios: await listUserProfiles() });
  } catch (error) {
    res.status(500).json({ erro: "Erro ao excluir usuário." });
  }
});

// ============= LOGOUT =============

app.get("/logout", (req, res) => {
  req.session.destroy(() => {
    res.redirect("/");
  });
});

app.post("/api/logout", requireAuth, (req, res) => {
  req.session.destroy((err) => {
    if (err) {
      return res.status(500).json({ erro: "Erro ao fazer logout." });
    }
    res.json({ mensagem: "Logout realizado com sucesso." });
  });
});

// ============= START =============

async function startServer() {
  try {
    const restored = await restoreWorkbookSnapshot();
    if (restored) {
      console.log("✅ Planilha atual restaurada do armazenamento persistente.");
    }
  } catch (error) {
    console.warn(`Não foi possível restaurar a planilha atual: ${error.message}`);
  }

  try {
    await restoreAllExtraWorkbooks();
  } catch (error) {
    console.warn(`Não foi possível restaurar planilhas auxiliares: ${error.message}`);
  }

  try {
    preparePracaWorkbookIfNeeded(findPracaWorkbookPath());
    watchPracaWorkbookChanges();
  } catch (error) {
    console.warn(`Não foi possível iniciar o tratamento automático de praças: ${error.message}`);
  }

  try {
    watchZmm113Changes();
    watchExtraWorkbookChanges();
  } catch (error) {
    console.warn(`Não foi possível iniciar o monitoramento das planilhas: ${error.message}`);
  }

  startSnapshotSyncPolling();

  const server = app.listen(PORT, "0.0.0.0", () => {
    console.log(`🚀 Servidor rodando em http://localhost:${PORT}`);
    console.log(`✅ Conectado ao Supabase: ${SUPABASE_URL}`);
  });

  server.on("error", (error) => {
    if (error.code === "EADDRINUSE") {
      console.error(`❌ A porta ${PORT} já está em uso.`);
      console.error(`Feche o servidor antigo ou rode: npm run restart`);
      process.exit(1);
    }

    console.error(`❌ Erro ao iniciar servidor: ${error.message}`);
    process.exit(1);
  });
}

startServer();
