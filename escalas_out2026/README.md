# Escalas de trabalho — fim da 6x1 (EC 141/2026) — out/2026

Simulação para a diretoria de como ficam as escalas da CRV Industrial com a Emenda Constitucional 141/2026
(promulgada em 16/10/2026): **2 folgas por semana e 42 h a partir de 15/12/2026**, **40 h a partir de dez/2027**.
Os turnos 24 h saem da 5x1, a 6x1 de dia vira 5x2 e o horário ADM com sábado vai para segunda a sexta, 8 h.

| Arquivo | O que é |
|---|---|
| `Relatorio_Escalas_EC141_OUT2026.html` | Relatório dinâmico: sugestões lado a lado, escala de cada uma, risco no MTE e simulador (escala por área, custo médio, hora extra). Abre direto no navegador, sem internet. Só números agregados, sem nomes |
| `Relatorio_Escalas_EC141_OUT2026.pdf` | 9 páginas A4 paisagem, no padrão dos relatórios da diretoria |
| `Simulacao_Escalas_EC141_OUT2026.xlsx` | Premissas editáveis (amarelo), simulação por função com fórmulas, turma G por horário, lista pessoa a pessoa com a escala proposta, códigos de escala e as grades de cada escala. **Não vai para o Git** (o repositório é público e a planilha tem nome e matrícula): é gerada na máquina pelo `gerar_relatorio.py` |
| `analise.py` | Lê as planilhas, classifica cada código de escala e calcula as sugestões; grava `analise.json` |
| `gerar_relatorio.py` | Gráficos, PDF (Edge headless), xlsx e, no fim, chama `gerar_html.py` |
| `gerar_html.py` + `relatorio_template.html` | Monta o HTML dinâmico com os dados embutidos |
| `recalcular_planilha.ps1` | Abre a planilha no Excel e grava os resultados das fórmulas |
| `Escalas_por_Departamento_OUT2026.html` | Validação por departamento (começa pela oficina): cada setor em cada escala, pessoas por turno editáveis, contratar ou sobra, escala sugerida e a grade da escala escolhida |
| `oficina.py` | Lê `dados/oficina/ativos_oficina.xlsx` (anotações por departamento e linhas em azul), cruza com a escala de hoje e calcula cada escala com turma fixa por turno |
| `gerar_departamentos.py` + `departamentos_template.html` | Monta a página por departamento |

### Operacional (frota 24 h)

Aba **Operacional** da mesma apresentação: os 84 equipamentos do *Detalhamento da frota* (lista em
`gerar_departamentos.FROTA`), cada um com 1 operador por turno nas 24 h, hoje em 5x1. Como o folguista roda entre
equipamentos, a conta é por equipamento: 5x1 = 3,6 pessoas por equipamento (hoje 303 operadores), 5x2 = 4,2 (+50),
4x2 de 8 h = 4,5 (+75), 12x36 e 3x3 = 4,0 (+33), 4x2 de 12 h = 3,0 (−51). Cada tipo fica fracionado e o total da frota
arredonda para cima. Quantidade e operadores por turno são editáveis na página.

A 4x2 de 12 h entra na simulação com **risco alto** (botão "4x2 12 h em todos" e opção de entrar na sugerida): no
cadastro, operadores de máquinas (335) e motoristas (250) são CLT urbano e, com a hora noturna reduzida, a noite de
12 h passa de 10 h contadas. Só vale com norma coletiva que troque a hora reduzida por adicional maior.

### Oficina por departamento

`python gerar_departamentos.py`. Regra de cada departamento em `oficina.REGRAS`, tirada das anotações da planilha.
"Garantir por turno" começa na média de hoje arredondada (2,5 vira 2), no mínimo 1 nos departamentos 24 h; PCM e
caldeiraria com 1 por turno. Com turma fixa por turno: 5x2 em todos dá +32; a escala sugerida (a de menos contratação
dentro da lei) dá +20, quase toda em 12x36. A 4x2 de 12 h entra com risco alto (CLT urbano, noite com hora reduzida).

## Para refazer

Colocar em `dados/` (não sobe para o git) o **Relatório de Revezamento** (`escalas.xlsx`) e a **Relação de
Funcionários** (`ativoss.xlsx`) e rodar, **no PowerShell** (no Bash da sessão do Claude o Edge não grava o PDF):

```powershell
python gerar_relatorio.py
powershell -ExecutionPolicy Bypass -File recalcular_planilha.ps1
```

## Como a conta é feita

- **5x1** = horários com 6 códigos de escala, um por dia de folga do ciclo de 6 dias (ex.: 566 a 571 = 07:00–15:20).
  Os códigos com "6x1" no nome, o 300 e os ADM com sábado (807, 814, 815, 801) são 6x1; 910, 76 e 921 são sábado
  compensado. A tabela fica na aba Códigos da planilha.
- A cobertura de hoje é mantida: o mesmo número de pessoas trabalhando por dia em cada turno.
- **Frota 24 h** = posto que roda de manhã, de tarde e de noite (o menor dos três turnos, por função, área e grupo de
  horários): 232 frotas. Os 177 postos-turno que sobram (quase todos de manhã) ficam em 5x2 de 8 h nas sugestões de 12 h.
- Pessoas por frota = turnos por dia × dias do ciclo ÷ dias trabalhados × (1 + rendição).
- Hora extra = horas por mês ÷ 200 × (1 + 50%) × (1 + 20% de reflexos de 13º e férias).

## Resultado (turnos 24 h, 1.047 pessoas em 5x1 hoje)

| Sugestão | Pessoas por frota | Contratar | Folha x hoje | Risco no MTE |
|---|---|---|---|---|
| 5x2 rotativa (3 turnos de 8 h, 7 turmas) | 4,2 | +175 | +16,7% | Baixo |
| 4x2 de 8 h | 4,5 | +262 | +25,0% | Baixo |
| 12x36 / 3x3 (11 h de trabalho) | 4,0 | +129 | +12,2% | Médio |
| 4x2 de 12 h com 2 h de intervalo (10 h = 8 h + 2 h extras) | 3,0 | −103 | +14,4% | Alto |
| 4x2 de 12 h + rendição (1 para 5 frotas) | 3,6 | +36 | +32,5% | Alto |

Mais: 6x1 de dia (122 pessoas), de 0 a +25, conforme a empresa mantenha ou não o sábado; horário ADM (303 pessoas),
sem contratação. A 4x2 de 12 h não é proibida, mas tem 10 h por dia (o teto da CLT), hora extra fixa, exigência de
licença do MTE ou norma coletiva em atividade insalubre (CLT art. 60 / 611-A, XIII) e o turno da noite do trabalhador
urbano passa de 10 h contadas. Validar com o jurídico antes de implantar.
