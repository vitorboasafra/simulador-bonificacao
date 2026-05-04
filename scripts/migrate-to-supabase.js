/**
 * Script de Migração de Dados Locais para Supabase
 * Uso: node scripts/migrate-to-supabase.js
 */

const fs = require("fs");
const path = require("path");
require("dotenv").config();

const { createClient } = require("@supabase/supabase-js");

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;

if (!SUPABASE_URL || !SUPABASE_KEY) {
  console.error("❌ SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY são obrigatórios!");
  process.exit(1);
}

const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
const ROOT_DIR = path.join(__dirname, "..");

async function migrateUsers() {
  console.log("📋 Migrando usuários...");

  const usuariosPath = path.join(ROOT_DIR, "usuarios.json");
  if (!fs.existsSync(usuariosPath)) {
    console.log("⚠️  Arquivo usuarios.json não encontrado.");
    return;
  }

  try {
    const users = JSON.parse(fs.readFileSync(usuariosPath, "utf8"));
    let count = 0;

    for (const [usuario, data] of Object.entries(users)) {
      const { error } = await supabase.from("usuarios").insert({
        usuario,
        email: data.email || `${usuario}@boasafrasementes.com.br`,
        senha: data.senha, // TODO: Hash com bcrypt!
        admin: Boolean(data.admin),
        confirmado: Boolean(data.confirmado)
      });

      if (error) {
        if (error.code === "23505") {
          console.log(`⚠️  ${usuario} já existe`);
        } else {
          console.error(`❌ Erro ao migrar ${usuario}:`, error.message);
        }
      } else {
        console.log(`✅ ${usuario} migrado`);
        count++;
      }
    }

    console.log(`\n✅ ${count} usuários migrados!`);
  } catch (error) {
    console.error("❌ Erro ao migrar usuários:", error.message);
  }
}

async function migrateHistory() {
  console.log("\n📊 Migrando histórico...");

  const historyPath = path.join(ROOT_DIR, "data", "historico.json");
  if (!fs.existsSync(historyPath)) {
    console.log("⚠️  Arquivo historico.json não encontrado.");
    return;
  }

  try {
    const history = JSON.parse(fs.readFileSync(historyPath, "utf8"));
    let count = 0;

    for (const item of history) {
      const { error } = await supabase.from("historico").insert({
        usuario: item.usuario,
        produtorNome: item.produtorNome,
        totalGeralFmt: item.totalGeralFmt,
        detalhes: item.detalhes,
        itens: item.itens,
        createdAt: item.createdAt
      });

      if (error) {
        console.error(`❌ Erro ao migrar item:`, error.message);
      } else {
        count++;
      }
    }

    console.log(`✅ ${count} itens de histórico migrados!`);
  } catch (error) {
    console.error("❌ Erro ao migrar histórico:", error.message);
  }
}

async function run() {
  console.log("🔄 Iniciando migração para Supabase...\n");

  await migrateUsers();
  await migrateHistory();

  console.log("\n✨ Migração concluída!");
  process.exit(0);
}

run();
