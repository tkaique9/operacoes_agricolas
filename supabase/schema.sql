-- =====================================================================
-- Operações agrícolas: alterações salvas das páginas de escalas
-- Schema para Supabase (Postgres), nomes abreviados como nos outros projetos.
--
-- Como rodar: painel do Supabase → SQL Editor → New query → cole este arquivo → Run.
-- Pode rodar de novo sem problema. O server.js também roda este arquivo ao conectar.
--
--   esc_est = estado das páginas de escalas
--   pag = página (relatorio, departamentos), doc = {estado, quando, anteriores[]}
--         (anteriores = histórico do Desfazer, até 100 versões), atu_em = atualizado em
-- =====================================================================
create table if not exists esc_est (
  pag text primary key,
  doc jsonb not null,
  atu_em timestamptz not null default now()
);

-- RLS ligado: sem isso a API REST do Supabase deixaria ler e gravar a tabela com a chave anon.
-- O servidor não é afetado, porque conecta como dono da tabela.
alter table esc_est enable row level security;
