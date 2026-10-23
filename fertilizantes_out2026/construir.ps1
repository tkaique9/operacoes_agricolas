# Refaz todos os entregáveis a partir da planilha de origem.
#
#   powershell -ExecutionPolicy Bypass -File construir.ps1 [-Planilha "<caminho do .xlsx>"]
#
# 1. extrair_dados.py          lê a planilha, confere as fórmulas e grava dados_consolidados.json
# 2. gerar_apresentacao.js     monta os slides (build\base.pptx) e o roteiro de animações
# 3. animar_e_exportar.ps1     aplica animações/transições no PowerPoint, salva o .pptx final e os PNGs
# 4. gerar_planilha_auxiliar.py + recalcular_planilha.ps1   planilha auxiliar com as contas em fórmulas
param([string]$Planilha = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Necessidade Fertilizante OUT 2026.xlsx'))
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$node = Join-Path $env:LOCALAPPDATA 'Programs\node-v24.20.0-win-x64\node.exe'
if (-not (Test-Path $node)) { $node = 'node' }

function Passo([string]$nome, [scriptblock]$acao) {
	Write-Output "== $nome"
	& $acao
	if ($LASTEXITCODE -ne 0) { throw "$nome falhou ($LASTEXITCODE)" }
}

Passo 'Extração e conferência' { python extrair_dados.py $Planilha }
Passo 'Slides' { & $node gerar_apresentacao.js }
Passo 'Animações e PNGs' { powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'animar_e_exportar.ps1') }
Passo 'Planilha auxiliar' { python gerar_planilha_auxiliar.py }
Passo 'Recálculo no Excel' { powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'recalcular_planilha.ps1') }
Write-Output '== Pronto'
