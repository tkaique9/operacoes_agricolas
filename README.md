# Tratos Culturais — Reunião Semanal

Apresentação da reunião semanal de tratos culturais no padrão **Indicadores Grupo
Japungu**: mesmo visual, navegação e convenções do `comparativo_moagem` (página única,
uma tela por tópico, sem rolagem de documento). É feita para a **TV da diretoria**: sem
parágrafos de explicação nem faixas de texto, o espaço vai para gráficos e números; em tela
de 1600 px ou mais tudo cresce (CSS e a função `fz()` do script das páginas).

A versão anterior, montada sobre o relatório RSC0819R, está guardada em `arquivo_rsc0819/`.

## Como abrir

Basta abrir o `index.html` no navegador — não precisa de servidor nem de internet. Para
apresentar, use F11. Para imprimir, os botões **Imprimir página** e **Imprimir
relatório** da barra de baixo (ou Ctrl+P, que imprime o relatório): A4 deitado, uma
página por folha, fundo branco, tabelas inteiras (sem rolagem; a letra diminui só se a
tabela não couber na largura) e gráficos redesenhados no tamanho da folha. A capa não sai
no papel. Para salvar em PDF, escolha "Salvar como PDF" na janela de impressão.
No papel as tabelas saem com letra de 8 pt (unidades e subtítulos com 6 pt), o resumo
gráfico em 2 colunas e o título "Indicadores" sempre junto dos gráficos; o relatório
inteiro dá 13 folhas. As regras de tamanho de tela (celular, tela baixa, TV) valem só na
tela (`@media screen and ...`): o Chrome mede a folha como A4 em pé ao imprimir e, sem
isso, aplicava as de celular no papel. Navegação: setas ←/→, barra de espaço, as bolinhas do rodapé, os
botões ou o menu de módulos do topo. `?fundo=off` desliga as partículas do fundo.

## Atualizar a semana

1. Coloque na pasta `dados/` as duas planilhas da semana (substituindo as antigas):
   - **Operações Tratos Culturais - 2025-2026** (`.xlsb`) — aba `Semanal-Acumulado`;
   - **Operações Tratos Culturais - SAFRA 2026** (`.xlsx`) — abas dos programas;
   - para o detalhe por fazenda e talhão: o **RSC0819R** da mesma semana (`.xlsx`), com o
     período de segunda até o DIA REF. Relatório de outra semana pode ficar na pasta: só
     entram as OS com data dentro da semana;
   - para as páginas dos programas e da adubação: a **impressão das planilhas em PDF**
     (ex.: `19.10.2026.pdf`), com as abas de pré-emergência, reforço bordadura, 1ª catação
     e adubação de soqueira. Vale o PDF mais novo da pasta.
2. Rode:

```bash
python atualizar_tratos.py
```

O script acha cada planilha pelas abas que ela tem (não pelo nome do arquivo) e regrava o
bloco entre `DADOS-INICIO` e `DADOS-FIM` do `index.html`. A semana é a do **DIA REF** da aba
`Semanal-Acumulado`: de segunda até o dia de referência e, como a planilha soma, mais o
domingo (ou outro dia depois do DIA REF) que teve produção. Precisa de
`pip install openpyxl pyxlsb pymupdf`.

Quando chega o PDF (a impressão das planilhas, ex.: `20.10.2026.pdf`), os números são
atualizados por ele, sem mexer no layout nem nas regras:

- as **tabelas dos programas** e da adubação (sempre);
- os **números da semana** — realizado por dia, pessoas, quadro de meta (nº de
  equipamentos/pessoas, meta por unidade, meta do dia, da semana e do mês) e o realizado
  do mês e do mês anterior de cada atividade, achada pelo nome —, **só se o PDF for da
  mesma semana** (mesmo DIA REF) que já está na apresentação. Semana anterior, histórico
  dos meses e atividades que a impressão não traz ficam como estão; as correções de meta
  (`META_EQUIP`, `EQUIP`, `META_PESSOA`) valem de novo; o domingo que a impressão esconde
  mas a planilha soma entra pela diferença entre o realizado da semana e os dias impressos.
  Semana nova pede as planilhas e o `atualizar_tratos.py` completo.

```bash
python atualizar_tratos.py --pdf dados/20.10.2026.pdf
```

A leitura do PDF fica em `programas_pdf.py`: cada número vai para a coluna cujo título
termina mais perto dele (os números são alinhados à direita), e as linhas são lidas como
estão — fazendas, subtotais e totais da planilha, sem refazer conta. O nome da
propriedade que a impressão gruda no bloco ("…ESPERANÇABEBEDOURO") é separado pelos nomes
de bloco dos subtotais. Região com erro de fórmula da planilha (`#N/D`, `#REF!`) fica com a
região da fazenda de cima no mesmo quadro (ex.: a 9849 São João, no meio das áreas
Cargill, sai como CARGILL). Os valores da 1ª catação por região só existem no gráfico da
impressão e são lidos dos rótulos das barras.

## Páginas

| Módulo | # | Página | O que mostra |
|---|---|---|---|
| — | — | Capa | Realizado, % da meta, eficiência e % da meta do mês; realizado acumulado da semana × meta acumulada e semana anterior |
| **Semana** | 1 | Semana por operação | Uma tabela por operação e atividade, em quatro grupos de colunas, cada um com título e linha forte à esquerda: **Semana** (meta, realizado e % da meta); **Equipamento ou homem · por dia** (quantidade **prevista** — o planejado da planilha — e **realizada** — a média de pessoas apontadas em campo por dia operado; para equipamento a planilha não aponta a quantidade do dia e o realizado é o que a conta do rendimento usa, o previsto nos dias com operação, em cinza —, meta e rendimento por equipamento/homem e a diferença, por dia operado; a meta e o rendimento vêm com a unidade: **ha/h/dia** quando a atividade é medida por pessoa apontada e **ha/equip/dia** quando é por equipamento — o Quadriciclo, apontado por operador, conta como equipamento, `ATIV_EQUIPAMENTO` no `index.html`); **Semana anterior** (realizado e variação); e **o mês até a data** (o que já foi produzido em ha, a meta do mês e a %, com a cor pelo ritmo esperado). Clicar na operação (›) abre o **modal da operação**; clicar no Total abre o mesmo modal para todas; clicar na atividade abre o dia a dia dela. Os títulos das colunas são centralizados (em todas as tabelas) |
| | 2 | Semana por dia | A mesma tabela aberta por dia: meta do dia, o realizado de cada dia (segunda a sábado e o domingo, se houve produção) com a % da meta do dia e o fundo verde, âmbar ou vermelho, e o fechamento da semana (meta, realizado e % da meta), por operação e atividade, com o total. Clicar abre os mesmos modais da página 1 |
| | 3 | Resumo gráfico | Para a diretoria, no padrão dos gráficos do comparativo: um painel por operação (grade 2 × 2; com mais de quatro operações, 3 colunas, e a vaga que sobra é o painel de todas as operações). À esquerda, os dois indicadores em barra — **meta da semana** (de 0 a 150%, traço da meta e marca cinza da semana anterior) e **o mês até a data** (de 0 a 100% da meta do mês, traço do ritmo esperado), na cor da faixa (vermelho abaixo de 50%, âmbar de 50% a 99%, verde na meta ou acima; no mês, em relação ao ritmo). À direita, as colunas de cada dia: o realizado (dourado) ao lado do mesmo dia da semana anterior (cinza), a linha tracejada da meta do dia e a pílula da % da meta do dia. No topo, a semana anterior (▲/▼) e a meta do dia; embaixo, a % da meta de cada atividade. Clicar no painel abre o **modal da operação** |
| **Programas da safra** | 4 | Pré-emergência | Uma página com rolagem: a tabela da aba como na impressão (fazendas, subtotal por região e total da safra) e, no fim, os **indicadores** em gráfico (total da safra, por região e ordens encerradas × abertas) |
| | 5 | Reforço bordadura | Igual: tabela da 3ª molécula e o gráfico do total (aberta, realizada, a realizar) |
| | 6 | 1ª catação química | Igual: as 71 fazendas numa página com rolagem e, no fim, o total da safra e por região |
| **Adubação** | 7 | Adubação localizada | Região 1 + Cargill e área de queima, com subtotal por bloco, total e observação; no fim, o resumo da localizada e da área de queima |
| | 8 | Adubação com cocho | Alencar (regiões 1 e 2) e Fábio do Prado (região 3), com subtotal por bloco e totais; no fim, as partes do total geral — o total da localizada (sem Cargill) e o da área de queima, que vêm da página 7, o total com cocho e o total geral — e a Cargill à parte. Nos indicadores, Alencar, Fábio do Prado, localizada + área de queima (sem Cargill) e o geral, que é a soma dos três. Nas duas páginas da adubação cada tipo tem as suas cores — cocho: azul-marinho (área), dourado (aplicada) e vermelho (a realizar); localizada: azul-acinzentado, bronze e vinho — e a barra empilhada traz o nome e a área de cada parte (dentro dela ou, quando a parte é fina, ao lado); no geral, a área, a aplicada e o a realizar vêm divididos em cocho e localizada |

As tabelas dos programas seguem um padrão só: as colunas vêm em grupos com título
(**Fazenda** — região, cód., bloco, propriedade, colheita e dias de atraso, nessa ordem em
todas; depois **Área**, **Aplicação**/**Aplicada**, **Saldo** ou **A realizar**, conforme o
programa) e uma linha entre cada coluna, mais forte entre os grupos; as linhas são
zebradas; subtotal em cinza, total em amarelo; dias de atraso acima de 20 em vermelho e a
**% concluída** de cada linha em barra, na régua de andamento de programa (verde a partir de 90%, âmbar de 50% a 89%, vermelho abaixo de 50% — não é a régua da meta da semana). O total da safra sai como **Total geral**; nas duas páginas de adubação a última linha (azul-marinho) é o **total geral da adubação** do resumo final da planilha e, na do cocho, antes dela, o total das duas equipes (Alencar + Fábio do Prado). O cabeçalho fica preso no alto enquanto a tabela
rola, e a letra só diminui se a tabela não couber na largura. Coluna sem nenhum valor
nos quadros da página não aparece (ex.: cana muda, na adubação).

Os indicadores do fim de cada página seguem o padrão dos gráficos da diretoria (o mesmo
do `comparativo_moagem`): título dentro do gráfico, grade com eixo, colunas com o topo
arredondado, valor em cima e a pílula da % sobre a área do programa (verde, âmbar ou
vermelho, como a meta). Cores fixas: azul-marinho é a área do programa (colhida, aberta,
a tratar), dourado o que foi feito (realizado/aplicado; dourado-claro a parte manual ou
a outra aplicação), vermelho o que falta (laranja: a realizar com menos de 20 dias) e
cinza-azulado o que é contexto (a colher, reforma, fornecedor, cata mato, catação).
A data de atualização da planilha vai no alto de cada página dos programas.

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
  manda cada operação para a frente pela descrição (drone → aplicação com drone, Uniport, cocho, manual;
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
| Dias sem meta | `SEM_META_DIAS`, no `atualizar_tratos.py`, tira a meta de dias da semana de uma atividade, além do domingo; hoje: Área Monitoramento sem meta no sábado (o drone de monitoramento não voa no sábado). A meta da semana conta só os dias com meta (5 × 300 = 1.500 ha) e a do mês também (dias do mês sem sábado e domingo); a meta do dia da operação e do total é a soma das atividades naquele dia |
| Atividades não planejadas | `NAO_PLANEJADAS`, no `atualizar_tratos.py`: entram com o realizado e as pessoas apontadas, mas sem meta (nem na operação, nem no total) e com a etiqueta "não planejada"; hoje: Roço |
| Áreas Cargill | ficam à parte (`cargill_a_parte`, no `atualizar_tratos.py`): os totais da localizada saem **sem Cargill** (a planilha traz "Total região 1 + Cargill" e "Total localizada" com a Cargill em todas as colunas, menos no saldo, que ela já soma sem a Cargill), a linha "Áreas Cargill · à parte" vem logo depois e a Cargill tem gráfico próprio |
| Total geral da adubação | a soma dos totais que as páginas mostram: localizada (sem Cargill) + área de queima + Alencar + Fábio do Prado. O "resumo final" da planilha não é usado: a linha do Alencar nele tem fórmula quebrada (#REF!) e não bate com o total do Alencar |
| Operações separadas | `SEPARA`, no `atualizar_tratos.py`, tira atividades da frente da planilha e faz delas uma operação própria, logo depois; hoje: a aplicação aérea fica **Aplicação com drone** e a Área Monitoramento vira **Monitoramento com drone** |
| Semana anterior | os mesmos dias da semana anterior, na mesma aba |
| Mês | do dia 1 até o último dia da semana (com o domingo trabalhado, como o "Realizado do Mês" da planilha); ritmo esperado = dias trabalhados até aqui (segunda a sábado) ÷ dias da meta do mês (meta mês ÷ meta dia) |
| Eficiência por equipamento/pessoa | realizado ÷ unidades-dia nos dias em que a atividade operou. Unidade = pessoa apontada no dia (Catação Manual, Cata Mato, Quadriciclo) ou, sem apontamento, o nº de equipamentos/pessoas planejado. Comparada com a meta por unidade da planilha; o planejado de pessoas e a meta por pessoa vêm da linha "Pessoas Dia" quando ela tem quadro de meta. `META_PESSOA`, no `atualizar_tratos.py`, fixa a meta por pessoa onde a da planilha não vale; hoje: Cata Mato, 4 ha/pessoa (a planilha traz 8). `EQUIP` fixa o nº de equipamentos e refaz a meta do dia, da semana e do mês; hoje: Quadriciclo, 2 × 10 ha = 20 ha/dia (a planilha traz 1), e Catação de Formiga, 2 equipamentos. `META_EQUIP` fixa a meta por equipamento antes disso; hoje: Catação de Formiga, 10 ha/equipamento (a planilha traz 1 de 20 ha; a meta do dia continua 20 ha) |
| Eficiência da operação | Σ realizado ÷ Σ (meta por unidade × unidades-dia) das atividades que operaram |
| % realizado do programa | realizado ÷ (realizado + a realizar) |
| % concluído (programas) | 1 − saldo ÷ área base da linha: a realizar ÷ área colhida (pré-emergência), ÷ área aberta (bordadura), (a realizar > 20 + < 20 dias) ÷ área aberta (1ª catação), a realizar ÷ área a tratar (adubação) |

Nomes de bloco escritos de jeitos diferentes entre as abas são unificados
(`SHEKINHA-I`/`SHEKINAH l` → SHEKINAH I, `MARCO-AMUY` → MARCO AMUY, `BR-154` → 154,
`SCALON 7 PLACAS` → SCALON SETE PLACAS) em `SINONIMOS_BLOCO`, no `atualizar_tratos.py`.

## Conferência das contas do PDF

```bash
python conferir_pdf.py dados/20.10.26.pdf
```

Lista o que não fecha na impressão das planilhas, sem mudar nada: realizado da semana ×
soma dos dias, meta do dia = quantidade × meta por unidade, meta da semana = meta do
dia × 6, total de cada operação = soma das atividades; nos programas, saldo = área −
realizado em cada linha, total = soma dos subtotais e o resumo final da adubação = soma
dos totais. A impressão esconde as fazendas e os blocos sem saldo, então em área a tratar
o total pode ser maior que a soma do que aparece; o saldo a realizar tem de fechar.

## O que o script confere na planilha

- **Linha colada fora do lugar:** se uma atividade repete, por 3 dias ou mais, os números
  da linha "Pessoas Dia" ou "Rendimento" de outra, esses dias saem da conta e viram aviso
  (na semana de 28/09 a 03/10, Colheitabilidade e Plantio Manual repetem pessoas e
  rendimento da Catação Manual).
- **Rendimento invertido:** se o "Rendimento H/Dia" da planilha for pessoas ÷ ha (Cata
  Mato), a apresentação usa ha ÷ pessoa e avisa.

## Convenções mantidas do `comparativo_moagem`

- Mesmo CSS base, tema, cromo, painel de detalhe e helpers de SVG, copiados de lá.
- Cor de destaque por módulo: dourado (Semana), verde-água (Programas da safra), azul (Adubação).
- Cor fixa por frente: Manual azul `#2E86C1`, Uniport oliva `#8A6D14`, Aplicação com drone
  verde-água `#3FA79C`, Monitoramento com drone roxo `#8E6CB8`, Cocho navy `#12224A`.
- Farol da meta: verde na meta ou acima, âmbar de 50% a 99%, vermelho abaixo de 50% — sempre
  com o número escrito ao lado.

## Publicar no Render

O site roda como **Web Service** Node (`server.js`, sem dependências; `render.yaml` traz a
configuração). O servidor publica as páginas e guarda as alterações das páginas de escalas.

- **Build Command:** `./render-build.sh` — copia só as páginas (`index.html`,
  `arquivo_rsc0819/index.html`, os relatórios de `escalas_out2026/` e o `salvar.js`) para
  `public/`. Scripts Python, templates e planilhas não vão para o site. Página nova? Inclua-a aqui.
- **Start Command:** `node server.js`
- **Variável `GITHUB_TOKEN`:** token *fine-grained* do GitHub só para este repositório, com
  **Contents: Read and write**.

Endereços: `/` (tratos culturais), `/arquivo_rsc0819/`, `/escalas/` (relatório de escalas) e
`/escalas_out2026/Escalas_por_Departamento_OUT2026.html`.

### Alterações salvas e Desfazer

No relatório de escalas e na página por departamento, tudo o que se muda (escala de cada
departamento, pessoas por turno, equipamentos, opções do simulador) é salvo sozinho, menos de
1 s depois, e vale para todo mundo que abrir o site. **↶ Desfazer** (ou Ctrl+Z) volta uma
alteração por vez, até a página como veio; o histórico guarda as últimas 100.

Cada alteração é um commit em `estado/<página>.json` no branch **`dados`** (criado sozinho na
primeira gravação). Fica fora da `main` para não refazer o deploy a cada clique; o histórico
completo está no Git. Para começar do zero, apague o arquivo nesse branch e reinicie o serviço.
Aberta do disco, no GitHub Pages ou no claude.ai a página funciona como antes, sem salvar.
Para testar na máquina: `bash render-build.sh && node server.js` (sem token, grava em `estado-local/`).
