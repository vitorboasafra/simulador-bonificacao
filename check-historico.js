const { createClient } = require('@supabase/supabase-js');
require('dotenv').config();

const supabaseAdmin = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY,
  { auth: { persistSession: false } }
);

async function createHistoricoTable() {
  try {
    console.log('📝 Verificando tabela historico...');
    
    // Tenta fazer uma query simples para ver se a tabela existe
    const { error: checkError } = await supabaseAdmin
      .from('historico')
      .select('count(*)', { count: 'exact', head: true })
      .limit(1);

    if (checkError && checkError.code === 'PGRST116') {
      console.log('❌ Tabela não existe no Supabase!');
      console.log('');
      console.log('📋 SOLUÇÃO: Execute este SQL no Supabase SQL Editor:');
      console.log('');
      console.log(`
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

CREATE INDEX IF NOT EXISTS idx_historico_usuario ON historico(usuario_id);
CREATE INDEX IF NOT EXISTS idx_historico_data ON historico(createdAt DESC);

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
      `);
      console.log('');
      console.log('Passos:');
      console.log('1. Acesse https://app.supabase.com');
      console.log('2. Vá para "SQL Editor"');
      console.log('3. Cole o SQL acima');
      console.log('4. Execute (Ctrl + Enter)');
      console.log('5. Teste novamente o servidor');
      process.exit(1);
    } else if (checkError) {
      throw checkError;
    } else {
      console.log('✅ Tabela historico existe!');
    }
  } catch (error) {
    console.error('❌ Erro:', error.message);
    process.exit(1);
  }
}

createHistoricoTable();
