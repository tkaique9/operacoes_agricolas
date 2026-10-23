# Auditoria de calibragem dos reboques canavieiros — out/2026

Relatório para a diretoria que cruza a conferência de calibragem (20 frotas, 280 posições de
pneu) com o peso por pneu e o sucateamento dos conjuntos (pesagens de 15/07 a 25/09/2026) e
com a tabela de carga x pressão (ETRTO) do 295/80R22,5 148 em rodado duplo.

| Arquivo | O que é |
|---|---|
| `Relatorio_Calibragem_Reboques_OUT2026.pdf` | 8 páginas A4 paisagem, no padrão dos relatórios de pneu |
| `Base_Auditoria_Calibragem_OUT2026.xlsx` | Pneu a pneu (pressão, faixa, capacidade, ação), rodado duplo, conjuntos e tabela ETRTO |
| `analise.py` | Lê a planilha, cruza com peso/sucateamento e a tabela ETRTO, grava `analise.json` |
| `gerar_relatorio.py` | Gráficos (matplotlib), HTML e PDF (Edge headless) e a base xlsx |

Para refazer: colocar `CONFERENCIA DE CALIBRAGEM.xlsx` em `dados/` (não sobe para o git) e rodar
`python gerar_relatorio.py`. Os números por conjunto (kg/pneu, sucateados, taxa) estão em
`analise.CONJUNTOS`, copiados do *Relatório Peso por Pneu x Sucateamento*.

## Resultado

- 83 de 280 pneus (29,6%) em 115–125 psi; só 20 em 120 psi. Média 108,2 psi. 34 sem leitura.
- 52 de 108 rodados duplos com diferença ≥ 10 psi entre externo e interno.
- Pelo peso: 9 eixos precisam de 118 psi na carga máxima (120 psi certo, sem folga);
  11 eixos precisam de 86 psi (110 psi bastam).
- Pneus perdidos x pressão média do conjunto: correlação −0,06. Pneus perdidos x % de pneus com
  pressão abaixo do que a carga exige: 0,76 (0,81 só nos 9 eixos; 10 conjuntos, indicativo).
