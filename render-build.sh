#!/usr/bin/env bash
# Monta public/ com as páginas publicadas no Render (ver render.yaml).
set -euo pipefail
rm -rf public
mkdir -p public/arquivo_rsc0819 public/escalas_out2026
cp index.html public/
cp arquivo_rsc0819/index.html public/arquivo_rsc0819/
cp escalas_out2026/Escalas_por_Departamento_OUT2026.html \
   escalas_out2026/Relatorio_Escalas_EC141_OUT2026.html \
   escalas_out2026/Relatorio_Escalas_EC141_OUT2026.pdf \
   escalas_out2026/salvar.js public/escalas_out2026/
# /escalas/: atalho para o relatório de escalas (página de redirecionamento)
mkdir -p public/escalas
cat > public/escalas/index.html <<'HTML'
<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Escalas</title>
<meta http-equiv="refresh" content="0; url=/escalas_out2026/Relatorio_Escalas_EC141_OUT2026.html">
</head><body><a href="/escalas_out2026/Relatorio_Escalas_EC141_OUT2026.html">Relatório de escalas</a></body></html>
HTML
echo "public/ pronto:"; find public -type f
