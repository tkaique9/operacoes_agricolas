# Abre a planilha auxiliar no Excel, recalcula todas as fórmulas, procura células com erro
# (#DIV/0!, #REF!, #NOME?...) e salva com os resultados gravados.
#
#   powershell -ExecutionPolicy Bypass -File recalcular_planilha.ps1
param([string]$Arquivo = (Join-Path $PSScriptRoot 'Base_Consolidada_Fertilizantes_OUT2026.xlsx'))
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
	$wb = $xl.Workbooks.Open((Resolve-Path $Arquivo).Path)
	$xl.CalculateFull()
	$formulas = 0; $erros = @()
	foreach ($ws in $wb.Worksheets) {
		foreach ($c in $ws.UsedRange.Cells) {
			if ($c.HasFormula) {
				$formulas++
				if ($c.Text -match '^#') { $erros += "$($ws.Name)!$($c.Address($false, $false)) = $($c.Text)" }
			}
		}
	}
	$wb.Save()
	$wb.Close($true)
	[pscustomobject]@{ formulas = $formulas; erros = $erros.Count; onde = ($erros -join '; ') } | ConvertTo-Json
}
finally {
	$xl.Quit()
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
