# Aplica animações e transições no PowerPoint e gera a apresentação final.
#
#   powershell -ExecutionPolicy Bypass -File animar_e_exportar.ps1
#
# Fase 1 (PowerPoint): abre build\base.pptx e aplica o roteiro de build\animacoes.json —
#   agrupa os cartões, faz os ajustes finos dos gráficos, cria as animações e transições.
# Fase 2 (Python): graficos_por_etapa.py grava no XML a entrada dos gráficos por categoria/série.
# Fase 3 (PowerPoint): reabre o arquivo, confere o que o PowerPoint leu em cada slide, salva a
#   apresentação final e exporta um PNG de cada slide para conferência.
#
# Precisa do Microsoft PowerPoint instalado (automação COM) e de Python com lxml.
param(
	[string]$Base = (Join-Path $PSScriptRoot 'build\base.pptx'),
	[string]$Roteiro = (Join-Path $PSScriptRoot 'build\animacoes.json'),
	[string]$Saida = (Join-Path $PSScriptRoot 'Necessidade_Fertilizantes_OUT2026.pptx'),
	[string]$Previews = (Join-Path $PSScriptRoot 'build\previews'),
	[switch]$SemPreviews
)
$ErrorActionPreference = 'Stop'
$spec = Get-Content -Raw -Encoding UTF8 $Roteiro | ConvertFrom-Json
$fase1 = Join-Path $PSScriptRoot 'build\fase1_animado.pptx'
$fase2 = Join-Path $PSScriptRoot 'build\fase2_graficos.pptx'

$EFEITO = @{ fade = 10; wipe = 22; zoom = 23; appear = 1 }           # MsoAnimEffect
$GATILHO = @{ clique = 1; com = 2; apos = 3 }                         # MsoAnimTriggerType
# Wipe: lado de onde o efeito parte (PowerPoint: "De baixo" = 3, "Da esquerda" = 4)
$DIRECAO = @{ cima = 1; direita = 2; baixo = 3; esquerda = 4 }        # MsoAnimDirection
$TRANSICAO = @{ fade = 3849; morph = 3954 }                           # ppEffectFadeSmoothly, ppEffectMorphByObject
$DURACAO_TRANSICAO = @{ fade = 0.7; morph = 1.0 }

function Cor([string]$hex) {
	$r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
	$g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
	$b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
	return $r + 256 * $g + 65536 * $b
}

$ppt = New-Object -ComObject PowerPoint.Application
$jaAberto = $ppt.Presentations.Count
try {
	# ------------------------------------------------------------------ fase 1
	$pres = $ppt.Presentations.Open((Resolve-Path $Base).Path, 0, 0, 0)
	try {
		foreach ($g in $spec.grupos) {
			$sl = $pres.Slides.Item([int]$g.slide)
			$grupo = $sl.Shapes.Range([object[]]@($g.membros)).Group()
			$grupo.Name = $g.nome
		}

		foreach ($c in $spec.graficos) {
			$ch = $pres.Slides.Item([int]$c.slide).Shapes.Item($c.forma).Chart
			if ($c.ocultarSerie) {
				$s = $ch.SeriesCollection([int]$c.ocultarSerie)
				$s.Format.Fill.Visible = 0
				$s.Format.Line.Visible = 0
				$s.HasDataLabels = $false
			}
			if ($c.corPontos) {
				$s = $ch.SeriesCollection([int]$c.corPontos.serie)
				for ($i = 0; $i -lt $c.corPontos.cores.Count; $i++) {
					$p = $s.Points($i + 1)
					$p.Format.Fill.Visible = -1
					$p.Format.Fill.Solid()
					$p.Format.Fill.ForeColor.RGB = (Cor $c.corPontos.cores[$i])
				}
			}
			if ($c.corRotulos) {
				$s = $ch.SeriesCollection([int]$c.corRotulos.serie)
				for ($i = 0; $i -lt $c.corRotulos.cores.Count; $i++) {
					$s.Points($i + 1).DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = (Cor $c.corRotulos.cores[$i])
				}
			}
			if ($c.tracejado) {
				foreach ($k in $c.tracejado) { $ch.SeriesCollection([int]$k).Format.Line.DashStyle = 4 }   # msoLineDash
			}
			if ($c.espessura) {
				foreach ($p in $c.espessura.PSObject.Properties) { $ch.SeriesCollection([int]$p.Name).Format.Line.Weight = [single]$p.Value }
			}
			if ($c.rotuloUltimo) {
				foreach ($k in $c.rotuloUltimo) {
					$s = $ch.SeriesCollection([int]$k)
					$p = $s.Points($s.Points().Count)
					$p.HasDataLabel = $true
					$p.DataLabel.ShowValue = $true
					$p.DataLabel.ShowSeriesName = $false
					$p.DataLabel.ShowCategoryName = $false
					$p.DataLabel.Position = 0                          # xlLabelPositionAbove
					$p.DataLabel.NumberFormat = '#,##0" t"'           # fonte herdada dos rótulos da série (12 pt, negrito)
				}
			}
		}

		foreach ($e in $spec.efeitos) {
			$sl = $pres.Slides.Item([int]$e.slide)
			$ef = $sl.TimeLine.MainSequence.AddEffect($sl.Shapes.Item($e.forma), $EFEITO[$e.efeito], 0, $GATILHO[$e.gatilho])
			$ef.Timing.Duration = [single]$e.duracao
			if ($e.atraso) { $ef.Timing.TriggerDelayTime = [single]$e.atraso }
			if ($e.efeito -eq 'wipe' -and $e.direcao) { $ef.EffectParameters.Direction = $DIRECAO[$e.direcao] }
		}

		foreach ($p in $spec.transicoes.PSObject.Properties) {
			$t = $pres.Slides.Item([int]$p.Name).SlideShowTransition
			$t.EntryEffect = $TRANSICAO[$p.Value]
			$t.Duration = $DURACAO_TRANSICAO[$p.Value]
		}
		$pres.SaveAs($fase1, 24)                                  # ppSaveAsOpenXMLPresentation
	}
	finally { $pres.Close() }

	# ------------------------------------------------------------------ fase 2
	& python (Join-Path $PSScriptRoot 'graficos_por_etapa.py') $fase1 $fase2 $Roteiro
	if ($LASTEXITCODE -ne 0) { throw "graficos_por_etapa.py falhou ($LASTEXITCODE)" }

	# ------------------------------------------------------------------ fase 3
	$pres = $ppt.Presentations.Open((Resolve-Path $fase2).Path, 0, 0, 0)
	try {
		$rel = @()
		foreach ($sl in $pres.Slides) {
			$seq = $sl.TimeLine.MainSequence
			$cliques = 0; $passosGrafico = 0
			for ($k = 1; $k -le $seq.Count; $k++) {
				$ef = $seq.Item($k)
				if ($ef.Timing.TriggerType -eq 1) { $cliques++ }
				if ($ef.Shape.HasChart) { $passosGrafico++ }        # fundo + uma etapa por categoria/série
			}
			$tr = $sl.SlideShowTransition.EntryEffect
			$nomeTr = 'outra'
			if ($tr -eq 3849) { $nomeTr = 'fade' } elseif ($tr -eq 3954) { $nomeTr = 'morph' }
			$rel += [pscustomobject]@{ slide = $sl.SlideIndex; efeitos = $seq.Count; noClique = $cliques; passosDeGrafico = $passosGrafico; transicao = $nomeTr }
		}
		$rel | Format-Table -AutoSize | Out-String | Write-Output
		$rel | ConvertTo-Json | Set-Content -Encoding UTF8 (Join-Path $PSScriptRoot 'build\relatorio_animacoes.json')

		$pres.SaveAs($Saida, 24)
		Write-Output "Salvo: $Saida"

		if (-not $SemPreviews) {
			New-Item -ItemType Directory -Force $Previews | Out-Null
			Get-ChildItem $Previews -Filter 'slide-*.png' | Remove-Item -Force
			foreach ($sl in $pres.Slides) {
				$sl.Export((Join-Path $Previews ('slide-{0:D2}.png' -f $sl.SlideIndex)), 'PNG', 1920, 1080)
			}
			Write-Output "Previews: $Previews"
		}
	}
	finally { $pres.Close() }
}
finally {
	if ($ppt.Presentations.Count -eq 0 -and $jaAberto -eq 0) { $ppt.Quit() }
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
}
