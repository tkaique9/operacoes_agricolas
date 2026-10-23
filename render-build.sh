#!/usr/bin/env bash
# Monta public/ com as páginas publicadas no Render (ver render.yaml).
set -euo pipefail
rm -rf public
mkdir -p public/arquivo_rsc0819 public/escalas_out2026
cp index.html public/
cp arquivo_rsc0819/index.html public/arquivo_rsc0819/
cp escalas_out2026/Escalas_por_Departamento_OUT2026.html \
   escalas_out2026/Relatorio_Escalas_EC141_OUT2026.html \
   escalas_out2026/Relatorio_Escalas_EC141_OUT2026.pdf public/escalas_out2026/
echo "public/ pronto:"; find public -type f
