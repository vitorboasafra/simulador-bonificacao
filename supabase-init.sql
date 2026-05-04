-- Supabase SQL de criação de tabelas para o Simulador de Bonificação

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

-- Inserir usuário inicial para login
-- Use o valor de usuario abaixo para fazer login no app.
INSERT INTO usuarios (id, usuario, email, senha, admin, confirmado)
VALUES (
  gen_random_uuid(),
  'vitor.boasafra',
  'vitor.boasafra@gmail.com',
  '@Boasafra2026',
  true,
  true
);
