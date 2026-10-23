# Necessidade de fertilizantes — posição 16/10/2026

Apresentação para a diretoria sobre a necessidade de fertilizantes para concluir a adubação
da safra da Usina CRV Industrial (Unidade Capinópolis-MG), feita a partir da planilha
**Necessidade Fertilizante OUT 2026.xlsx**. Nenhum valor da planilha de origem foi alterado.

## Entregáveis

| Arquivo | O que é |
|---|---|
| `Necessidade_Fertilizantes_OUT2026.pptx` | 17 slides 16:9 (16 + anexo), editáveis, com gráficos nativos, animações, transições e notas do apresentador |
| `Base_Consolidada_Fertilizantes_OUT2026.xlsx` | Premissas com a célula de origem, todas as contas em fórmulas (451, sem erro), curva, sensibilidade editável, indicadores, inconsistências e conferência |
| `dados_consolidados.json` | Números consolidados que alimentam a apresentação e a planilha auxiliar |
| `build/previews/slide-NN.png` | Imagem de cada slide, exportada pelo próprio PowerPoint (conferência) |

## Diagnóstico da planilha de origem

As duas abas têm a mesma estrutura e são **os dois cenários** da apresentação. A única
premissa que muda é C13, os dias em que a vinhaça localizada ainda atende o nível de potássio:

| | Cenário A — `LOC. 1.000ha` | Cenário B — `LOC. 450ha` |
|---|---|---|
| Dias de vinhaça com K adequado (C13) | 20 | 0 |
| Área de localizada (C14 = 20 × 50,78 ha/dia) | 1.015,6 ha | 0 ha (só os 446,63 ha da Cargill) |

Quadros de cada aba: 1) projeção da vinhaça localizada (datas, ritmo); 2) áreas adubadas;
3) áreas remanescentes (área, fórmula, dose, volume); 4) resumo (necessidade × estoque ×
comprado a receber) e estoque por fórmula em 16/10/2026.

| Informado na planilha | Calculado pela planilha | Ausente |
|---|---|---|
| Datas, ritmo da localizada, dias de K, áreas colhidas/a colher/realizadas, cana bisada, doses, estoques, comprado a receber | Área a adubar, necessidade por fórmula, saldo, necessidade de aquisição, área equivalente do estoque | Preços, datas de entrega, prazo de fornecimento, demanda por mês, ritmo da adubação mineral, detalhe por fazenda/bloco, necessidade do plantio |

As 85 fórmulas e totais da planilha foram refeitos em Python e conferem com os valores
gravados pelo Excel; mais 2 conferências mostram que a sensibilidade por dias reproduz
exatamente os cenários A e B (aba **Conferência** da planilha auxiliar).

## Resumo executivo (cenário A; cenário B entre parênteses)

- **Área ainda a adubar: 11.859,27 ha**, igual nos dois cenários — 50% da área de tratos já
  foi adubada e 3.076,93 ha ficam como cana bisada. O cenário muda só o produto.
- **Necessidade: 5.535,6 t** (5.706,3 t) — 16-00-20 = 5.421,8 t (5.706,3 t), 98% do volume;
  46-00-00 = 113,8 t (0 t).
- **16-00-20:** estoque 1.132 t + comprado a receber 2.030 t = 3.162 t, cobertura de 58,3%
  (55,4%). **Compra necessária: 2.259,8 t** (2.544,3 t).
- **46-00-00 sobra:** 1.506,97 t disponíveis para 113,8 t de uso → excedente de 1.393,2 t
  (1.507,0 t).
- **Quando:** com a necessidade distribuída por igual até o término da safra (premissa
  ilustrativa), o estoque de 16-00-20 acaba por volta de 28/10 e o disponível por volta de
  20/11 (27/10 e 18/11 no B). A compra adicional precisa chegar antes disso.
- **Alavanca:** cada dia a mais de vinhaça localizada com K adequado (50,78 ha) troca 25,4 t
  de 16-00-20 por 10,2 t de 46-00-00 — o produto em falta pelo produto que sobra.
- **Financeiro:** sem preços na base. Cada R$ 100/t no 16-00-20 vale R$ 226 mil na compra
  (R$ 254 mil no B) e R$ 50/ha no custo por hectare.
- **Alternativa a avaliar com a agronomia:** o excedente de 46-00-00 contém 640,9 t de N, acima
  das 361,6 t de N do déficit de 16-00-20; faltaria uma fonte de K para 452,0 t de K₂O.
- **Estoque total:** 2.600,3 t em 10 fórmulas (62% de socaria); o plantio projetado de 2.400 ha
  não tem necessidade calculada na planilha.

**Decisões propostas:** 1) aprovar a compra de 16-00-20 no cenário escolhido; 2) definir o
destino do excedente de 46-00-00; 3) obter o cronograma de entrega das 3.180 t já compradas e
os preços.

## Inconsistências e dados faltantes

| # | Gravidade | Onde | Achado |
|---|---|---|---|
| 1 | Alta | Ambas, linhas 26–29 | Quadro de vinhaça localizada não fecha: total 8.517,09 ha × 6.081,74 ha nas linhas (diferença 2.435,35 ha); B27+B28 = 5.328,37 ≠ B22 = 5.795,19. Linhas sem descrição na coluna A |
| 2 | Média | `LOC. 450ha`!H38 | 0 digitado; na aba A é `=C14-H36` (que daria −446,63 ha). Não se atualiza se C13 mudar |
| 3 | Baixa | Ambas, H36 | `=1200-F36` usa 1.200 digitado em vez de E36 |
| 4 | Média | Ambas, D38 | `=11593+9354+8779` sem identificar as parcelas (blocos? fazendas?) |
| 5 | Média | Ambas, K58 × K60:K61 | Oxyfertil (142,19 t; 284,38 ha) fora dos subtotais SOCARIA e PLANTIO |
| 6 | Baixa | Rótulos | "16-00-20." e "16-00-20" (idem 46-00-00) — quebram PROCV/tabela dinâmica |
| 7 | Média | Estoque × resumo | 21-00-00 (126 t), 13-00-34 (4 t), 25-00-00 (0,85 t) e Oxyfertil fora do balanço |
| 8 | Média | Nota A42 | Plantio de 2.400 ha sem necessidade calculada; estoque de plantio cobre 1.397 ha |
| 9 | Alta | Planilha inteira | Sem preços, datas de entrega, prazo de fornecimento, demanda mensal, ritmo da adubação mineral |
| 10 | Baixa | `LOC. 450ha`, seções | Numeração "1." repetida; grafias "Fertlizante" e "Rítmo" |

## Premissas da apresentação

- Cenários = abas da planilha; nenhum cenário foi inventado.
- Nutrientes pela garantia da fórmula (N-P₂O₅-K₂O em %); o teor de K da vinhaça não foi estimado.
- Curva de abastecimento: distribuição uniforme da necessidade até 15/12, só para ordem de
  grandeza — está marcada como premissa ilustrativa no slide.
- Sensibilidade por dias (10, 30, 45): mesma lógica da aba A, só variando C13; não é previsão.
- Nenhum preço, data de entrega ou responsável nominal foi assumido. As áreas do plano de ação
  são sugestões; nomes e prazos ficam a definir.

## Como apresentar

- **Transições:** Fade suave (0,7 s) em todos os slides; **Morph** (1 s) na entrada do slide 11 —
  o cartão "2.260 t" desliza do slide 10 para a conta do orçamento. Morph exige PowerPoint 2019
  ou Microsoft 365; em versões anteriores o arquivo cai para Fade.
- **Entradas automáticas:** depois da transição, o conteúdo entra sozinho na ordem da narrativa
  (Fade nos cartões, Wipe nos gráficos e na linha do tempo). Os gráficos de barras entram por
  categoria (por série no slide 4): primeiro eixos e legenda, depois cada barra.
- **Cliques** — o que fica guardado para o apresentador revelar:

| Slide | No clique |
|---|---|
| 2 | Quadro "Decisões para a diretoria" |
| 3 | Nota: a área a adubar é a mesma nos dois cenários |
| 8 | Cartão "2.260 t precisam chegar antes de ~20/11" |
| 9 | Alerta: áreas equivalentes não se somam; plantio sem necessidade calculada |
| 12 | Alternativa técnica (N do excedente + fonte de K) |
| 16 | As quatro conclusões, uma por clique; no 5º, o quadro de decisões |

- **Notas do apresentador** em todos os slides, com o roteiro e as células de origem.
- Sem animações (PDF, impressão ou "Apresentar sem animação" em Configurar Apresentação), todo
  o conteúdo aparece completo.

## Como refazer

Requisitos: Windows com PowerPoint e Excel, Python 3 com `openpyxl` e `lxml`, e Node.js. O Node
foi instalado em modo portátil (sem administrador) em
`%LOCALAPPDATA%\Programs\node-v24.20.0-win-x64` e **não** está no PATH; o `construir.ps1` o
encontra lá. Se a pasta `node_modules` não existir, rode `npm install` nesta pasta.

```bash
powershell -ExecutionPolicy Bypass -File construir.ps1
```

Para outra planilha: `construir.ps1 -Planilha "<caminho do .xlsx>"` (mesma estrutura de abas).

| Etapa | Script | Saída |
|---|---|---|
| 1 | `extrair_dados.py` | `dados_consolidados.json`; para se alguma fórmula da origem divergir |
| 2 | `gerar_apresentacao.js` | `build/base.pptx` e `build/animacoes.json` (roteiro de animações) |
| 3 | `animar_e_exportar.ps1` + `graficos_por_etapa.py` | `.pptx` final, `build/previews/*.png`, `build/relatorio_animacoes.json` |
| 4 | `gerar_planilha_auxiliar.py` + `recalcular_planilha.ps1` | `.xlsx` auxiliar recalculado pelo Excel |

## Limitações conhecidas

- O PowerPoint não aceita, por automação, animação de gráfico por categoria ("Invalid request",
  inclusive em gráfico nativo). Ela é gravada direto no XML (`graficos_por_etapa.py`) e o
  arquivo é reaberto no PowerPoint, que reconhece cada etapa e a mantém ao salvar.
- A capa e o fechamento usam uma ilustração vetorial de canavial gerada pelo script, não
  fotografia. Para usar foto da usina: Exibir > Slide Mestre > layout CAPA > Formatar plano de fundo.
- Os gráficos de cascata (slides 3 e 10) usam uma série invisível "apoio da cascata (oculto)"; ao
  editar os dados no PowerPoint, mantenha-a.
