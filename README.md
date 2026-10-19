# Tratos Culturais — Reunião Semanal

Apresentação da reunião semanal de tratos culturais no padrão **Indicadores Grupo
Japungu**: mesmo visual, navegação e convenções do `comparativo_moagem` (página única,
uma tela por tópico, sem rolagem de documento). É feita para a **TV da diretoria**: sem
parágrafos de explicação nem faixas de texto, o espaço vai para gráficos e números; em tela
de 1600 px ou mais tudo cresce (CSS e a função `fz()` do script das páginas).

A versão anterior, montada sobre o relatório RSC0819R, está guardada em `arquivo_rsc0819/`.

## Como abrir

Basta abrir o `index.html` no navegador — não precisa de servidor nem de internet. Para
apresentar, use F11. Navegação: setas ←/→, barra de espaço, as bolinhas do rodapé, os
botões ou o menu de módulos do topo. `?fundo=off` desliga as partículas do fundo.

## Atualizar a semana

1. Coloque na pasta `dados/` as duas planilhas da semana (substituindo as antigas):
   - **Operações Tratos Culturais - 2025-2026** (`.xlsb`) — aba `Semanal-Acumulado`;
   - **Operações Tratos Culturais - SAFRA 2026** (`.xlsx`) — abas dos programas;
   - para o detalhe por fazenda e talhão: o **RSC0819R** da mesma semana (`.xlsx`), com o
     período de segunda até o DIA REF. Relatório de outra semana pode ficar na pasta: só
     entram as OS com data dentro da semana.
2. Rode:

```bash
python atualizar_tratos.py
```

O script acha cada planilha pelas abas que ela tem (não pelo nome do arquivo) e regrava o
bloco entre `DADOS-INICIO` e `DADOS-FIM` do `index.html`. A semana é a do **DIA REF** da aba
`Semanal-Acumulado`: de segunda até o dia de referência e, como a planilha soma, mais o
domingo (ou outro dia depois do DIA REF) que teve produção. Precisa de
`pip install openpyxl pyxlsb`. Os avisos que ele imprime também aparecem na página
"Pontos de atenção".

## Páginas

| Módulo | # | Página | O que mostra |
|---|---|---|---|
| — | — | Capa | Realizado, % da meta, eficiência e % da meta do mês; realizado acumulado da semana × meta acumulada e semana anterior |
| **Semana** | 1 | Semana por operação | Uma tabela por operação e atividade: meta, realizado, % da meta, semana anterior e variação, eficiência (% da meta por unidade) e o mês (% da meta, com a cor pelo ritmo esperado). Clicar na operação (›) abre o **modal da operação**; clicar no Total abre o mesmo modal para todas; clicar na atividade abre o dia a dia dela |
| **Programas da safra** | 2 | Andamento dos programas | Realizado e saldo de pré-emergente (soca e área de queima), 3ª molécula, 1ª e 2ª catação e adubação (localizada e cocho) |
| | 3 | Saldo por bloco | Mapa de calor bloco × programa, em hectares a realizar |
| **Atenção** | 4 | Pontos de atenção | Só o que pede ação: abaixo de 50% da meta, eficiência abaixo de 50%, programas abaixo de 50%, 1ª catação com mais de 20 dias e o que conferir na planilha (cartão sem item não aparece) |

O **modal da operação** tem três abas:

- **Dia a dia** — gráfico do realizado por dia × meta do dia, empilhado por atividade (no de
  todas, por operação), e a tabela dia a dia com a semana e a % da meta. Clicar numa barra
  abre as atividades daquele dia;
- **Eficiência** — unidades planejadas e em campo, dias com operação, hectares por unidade
  por dia × meta e % da meta;
- **Mês e histórico** — os cinco meses anteriores, o mês até a data, a meta do mês e a % da
  meta, com o traço no ritmo esperado.

Cada número tem um lugar só: na página, o resumo da semana, da eficiência e do mês; no modal,
o detalhe da operação (o total de cada dia em cima das barras e o da semana na última linha
da tabela); a data da planilha só na capa. Uma frente nova da aba `Semanal-Acumulado` entra
sozinha na tabela e ganha o seu modal.
Toda tabela rola dentro do próprio cartão, então nada passa da altura da tela, seja a TV ou
o navegador com a barra do claude.ai em cima.

Clicar num cartão, barra, linha ou bloco abre o **painel de detalhe** com a conta por trás
do número (dia a dia da atividade, atividades da frente, blocos do programa…).

### Até fazenda e talhão (só da semana)

No modal de cada operação (aba Dia a dia), o painel desce em níveis, com **‹ Voltar** entre
eles: operação (linha "Total") ou atividade → **fazendas da semana** (talhões, área,
aplicado, dias) → **talhões da fazenda** (dia, talhão, área, aplicado, operação, OS, corte,
produto e dose real × recomendada). A etiqueta "N faz. · N talh." marca as linhas que têm
esse detalhe. Entram só fazendas e talhões trabalhados nos dias da semana, vindos de:

- **RSC0819R** (relatório de operações do sistema, `.xlsx`) colocado em `dados/`: o script
  reconhece pelo cabeçalho "Período Acumulado", usa só as OS com data dentro da semana e
  manda cada operação para a frente pela descrição (drone → aérea, Uniport, cocho, manual;
  formiga → Catação de Formiga; drone → Drone Inseticida, o antigo Drone Terceiro). O sistema não separa Uniport
  próprio de terceiro nem as equipes do cocho, então esses ficam no nível da operação;
- aba **Realizado - 3ª Mol** da planilha SAFRA: OS da 3ª molécula (Uniport bordadura); a
  coluna "Recom." é a área de borda aplicada no talhão. Entra quando a OS não veio no RSC.

O painel de fazendas mostra quanto do realizado da semana tem fazenda e talhão
identificados. Operação sem esse detalhe na semana não mostra etiqueta nem botão.

## As contas

| Indicador | Conta |
|---|---|
| Realizado da semana | como na planilha: soma dos dias de segunda até o DIA REF na aba `Semanal-Acumulado`, mais o domingo trabalhado (ele aparece como mais uma coluna só quando teve produção) |
| Meta da semana | meta dia × dias da semana (a coluna "Meta Semana" da planilha); domingo não tem meta |
| Meta da frente | soma só das atividades que **operaram** na semana (realizado > 0), como faz a planilha. Atividade ou operação **sem produção na semana não aparece em nenhuma tela** |
| Operações fora | `FRENTES_FORA`, no `atualizar_tratos.py`, tira uma frente inteira da apresentação e dos totais; hoje: Quebra lombo (não é aplicação desta operação) |
| Semana anterior | os mesmos dias da semana anterior, na mesma aba |
| Mês | do dia 1 até o último dia da semana (com o domingo trabalhado, como o "Realizado do Mês" da planilha); ritmo esperado = dias trabalhados até aqui (segunda a sábado) ÷ dias da meta do mês (meta mês ÷ meta dia) |
| Eficiência por equipamento/pessoa | realizado ÷ unidades-dia nos dias em que a atividade operou. Unidade = pessoa apontada no dia (Catação Manual, Cata Mato, Quadriciclo) ou, sem apontamento, o nº de equipamentos/pessoas planejado. Comparada com a meta por unidade da planilha; o planejado de pessoas e a meta por pessoa vêm da linha "Pessoas Dia" quando ela tem quadro de meta. `META_PESSOA`, no `atualizar_tratos.py`, fixa a meta por pessoa onde a da planilha não vale; hoje: Cata Mato, 4 ha/pessoa (a planilha traz 8). `EQUIP` fixa o nº de equipamentos e refaz a meta do dia, da semana e do mês; hoje: Quadriciclo, 2 × 10 ha = 20 ha/dia (a planilha traz 1), e Catação de Formiga, 2 equipamentos. `META_EQUIP` fixa a meta por equipamento antes disso; hoje: Catação de Formiga, 10 ha/equipamento (a planilha traz 1 de 20 ha; a meta do dia continua 20 ha) |
| Eficiência da operação | Σ realizado ÷ Σ (meta por unidade × unidades-dia) das atividades que operaram |
| % realizado do programa | realizado ÷ (realizado + a realizar) |
| Saldo por bloco | o quadro BLOCO \| Á REALIZAR da aba (pré-emergente cana soca), o resumo `ABERTO-ENCERRADO_BLOCO` (adubação) ou a soma das fazendas por bloco (demais) |

Nomes de bloco escritos de jeitos diferentes entre as abas são unificados
(`SHEKINHA-I`/`SHEKINAH l` → SHEKINAH I, `MARCO-AMUY` → MARCO AMUY, `BR-154` → 154,
`SCALON 7 PLACAS` → SCALON SETE PLACAS) em `SINONIMOS_BLOCO`, no `atualizar_tratos.py`.

## O que o script confere na planilha

- **Linha colada fora do lugar:** se uma atividade repete, por 3 dias ou mais, os números
  da linha "Pessoas Dia" ou "Rendimento" de outra, esses dias saem da conta e viram aviso
  (na semana de 28/09 a 03/10, Colheitabilidade e Plantio Manual repetem pessoas e
  rendimento da Catação Manual).
- **Rendimento invertido:** se o "Rendimento H/Dia" da planilha for pessoas ÷ ha (Cata
  Mato), a apresentação usa ha ÷ pessoa e avisa.

## Convenções mantidas do `comparativo_moagem`

- Mesmo CSS base, tema, cromo, painel de detalhe e helpers de SVG, copiados de lá.
- Cor de destaque por módulo: dourado (Semana), verde-água (Programas), creme (Atenção).
- Cor fixa por frente: Manual azul `#2E86C1`, Uniport oliva `#8A6D14`, Aérea verde-água
  `#3FA79C`, Cocho navy `#12224A`.
- Farol da meta: verde na meta ou acima, âmbar de 50% a 99%, vermelho abaixo de 50% — sempre
  com o número escrito ao lado.
