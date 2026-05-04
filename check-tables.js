const { createClient } = require('@supabase/supabase-js');
require('dotenv').config();

const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY);

async function checkTables() {
  try {
    console.log('🔍 Verificando tabelas no Supabase...');

    // Testar se tabela usuarios existe
    const { error: usuariosError } = await supabase
      .from('usuarios')
      .select('id')
      .limit(1);

    if (usuariosError && usuariosError.message.includes('relation "public.usuarios" does not exist')) {
      console.log('❌ Tabela "usuarios" NÃO existe!');
      console.log('📋 Execute este SQL no painel do Supabase (SQL Editor):');
      console.log(`
CREATE TABLE usuarios (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  usuario TEXT UNIQUE NOT NULL,
  email TEXT UNIQUE NOT NULL,
  senha TEXT NOT NULL,
  admin BOOLEAN DEFAULT FALSE,
  confirmado BOOLEAN DEFAULT TRUE,
  created_at TIMESTAMP DEFAULT NOW()
);

ALTER TABLE usuarios ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Usuários podem ler seus dados"
ON usuarios FOR SELECT
USING (auth.uid() = id);
      `);
    } else {
      console.log('✅ Tabela "usuarios" existe');
    }

    // Testar se tabela historico existe
    const { error: historicoError } = await supabase
      .from('historico')
      .select('id')
      .limit(1);

    if (historicoError && historicoError.message.includes('relation "public.historico" does not exist')) {
      console.log('❌ Tabela "historico" NÃO existe!');
      console.log('📋 Execute este SQL no painel do Supabase (SQL Editor):');
      console.log(`
CREATE TABLE historico (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  usuario_id UUID NOT NULL,
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
      `);
    } else {
      console.log('✅ Tabela "historico" existe');
    }

    // Testar se tabela configuracoes existe
    const { error: configError } = await supabase
      .from('configuracoes')
      .select('id')
      .limit(1);

    if (configError && configError.message.includes('relation "public.configuracoes" does not exist')) {
      console.log('❌ Tabela "configuracoes" NÃO existe!');
      console.log('📋 Execute este SQL no painel do Supabase (SQL Editor):');
      console.log(`
CREATE TABLE configuracoes (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  chave TEXT UNIQUE NOT NULL,
  valor TEXT,
  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW()
);

ALTER TABLE configuracoes ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Qualquer um pode ler"
ON configuracoes FOR SELECT USING (true);
      `);
    } else {
      console.log('✅ Tabela "configuracoes" existe');
    }

    console.log('🎯 Verificação concluída! Se alguma tabela não existir, execute o SQL mostrado acima.');

  } catch (err) {
    console.log('❌ Erro geral:', err.message);
  }
}

checkTables();