const { createClient } = require('@supabase/supabase-js');
require('dotenv').config();

const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY);

async function initTables() {
  try {
    console.log('🔄 Criando tabelas no Supabase...');

    // Criar extensão pgcrypto
    const { error: extError } = await supabase.rpc('exec_sql', {
      sql: 'CREATE EXTENSION IF NOT EXISTS "pgcrypto";'
    });

    if (extError) {
      console.log('⚠️ Erro na extensão (pode ser normal):', extError.message);
    }

    // Criar tabela usuarios
    const createUsuariosSQL = `
      CREATE TABLE IF NOT EXISTS usuarios (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        usuario TEXT UNIQUE NOT NULL,
        email TEXT UNIQUE NOT NULL,
        senha TEXT NOT NULL,
        admin BOOLEAN DEFAULT FALSE,
        confirmado BOOLEAN DEFAULT TRUE,
        created_at TIMESTAMP DEFAULT NOW()
      );

      ALTER TABLE usuarios ENABLE ROW LEVEL SECURITY;

      DROP POLICY IF EXISTS "Usuários podem ler seus dados" ON usuarios;
      CREATE POLICY "Usuários podem ler seus dados"
      ON usuarios FOR SELECT
      USING (auth.uid() = id);

      DROP POLICY IF EXISTS "Admin pode ler todos" ON usuarios;
      CREATE POLICY "Admin pode ler todos"
      ON usuarios FOR SELECT
      USING (
        EXISTS (
          SELECT 1 FROM usuarios WHERE id = auth.uid() AND admin = true
        )
      );
    `;

    const { error: error1 } = await supabase.rpc('exec_sql', { sql: createUsuariosSQL });

    if (error1) {
      console.log('❌ Erro ao criar tabela usuarios:', error1.message);
    } else {
      console.log('✅ Tabela usuarios criada/atualizada');
    }

    // Criar tabela historico
    const createHistoricoSQL = `
      CREATE TABLE IF NOT EXISTS historico (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        usuario_id UUID NOT NULL,
        produtorNome TEXT NOT NULL,
        usuario TEXT NOT NULL,
        totalGeralFmt TEXT NOT NULL,
        detalhes JSONB NOT NULL,
        itens INTEGER DEFAULT 0,
        createdAt TIMESTAMP DEFAULT NOW()
      );

      CREATE INDEX IF NOT EXISTS idx_historico_usuario ON historico(usuario_id);
      CREATE INDEX IF NOT EXISTS idx_historico_data ON historico(createdAt DESC);

      ALTER TABLE historico ENABLE ROW LEVEL SECURITY;

      DROP POLICY IF EXISTS "Usuários veem seu histórico" ON historico;
      CREATE POLICY "Usuários veem seu histórico"
      ON historico FOR SELECT
      USING (usuario_id = auth.uid());

      DROP POLICY IF EXISTS "Admin vê todos" ON historico;
      CREATE POLICY "Admin vê todos"
      ON historico FOR SELECT
      USING (
        EXISTS (
          SELECT 1 FROM usuarios WHERE id = auth.uid() AND admin = true
        )
      );
    `;

    const { error: error2 } = await supabase.rpc('exec_sql', { sql: createHistoricoSQL });

    if (error2) {
      console.log('❌ Erro ao criar tabela historico:', error2.message);
    } else {
      console.log('✅ Tabela historico criada/atualizada');
    }

    // Criar tabela configuracoes
    const createConfigSQL = `
      CREATE TABLE IF NOT EXISTS configuracoes (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        chave TEXT UNIQUE NOT NULL,
        valor TEXT,
        created_at TIMESTAMP DEFAULT NOW(),
        updated_at TIMESTAMP DEFAULT NOW()
      );

      ALTER TABLE configuracoes ENABLE ROW LEVEL SECURITY;

      DROP POLICY IF EXISTS "Qualquer um pode ler" ON configuracoes;
      CREATE POLICY "Qualquer um pode ler"
      ON configuracoes FOR SELECT USING (true);

      DROP POLICY IF EXISTS "Apenas admin edita" ON configuracoes;
      CREATE POLICY "Apenas admin edita"
      ON configuracoes FOR UPDATE USING (
        EXISTS (
          SELECT 1 FROM usuarios WHERE id = auth.uid() AND admin = true
        )
      );
    `;

    const { error: error3 } = await supabase.rpc('exec_sql', { sql: createConfigSQL });

    if (error3) {
      console.log('❌ Erro ao criar tabela configuracoes:', error3.message);
    } else {
      console.log('✅ Tabela configuracoes criada/atualizada');
    }

    console.log('🎉 Inicialização do banco de dados concluída!');

  } catch (err) {
    console.log('❌ Erro geral:', err.message);
  }
}

initTables();