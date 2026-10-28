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
| `Escalas_por_Departamento_OUT2026.html` | Validação por departamento (começa pela oficina): cada setor em cada escala, pessoas por turno editáveis, contratar ou sobra, escala sugerida e a grade da escala escolhida. No painel de cada departamento ou equipamento dá para trocar a escala, salvar e passar para o próximo. As tabelas grandes têm **Tela cheia**. **Imprimir página** e **Imprimir relatório** (ou Ctrl+P) saem em A4 deitado, fundo branco, uma página por folha, reduzida até 68% quando precisa; a matriz da frota vai em duas folhas com o cabeçalho repetido. **Escolher páginas** abre a lista das páginas por módulo para marcar o que vai para o papel (a última escolha fica guardada no navegador). A capa não sai no papel. No claude.ai a página roda num quadro que bloqueia a impressão: os botões abrem a cópia do GitHub Pages numa aba nova, com as escolhas feitas, e ela já abre a janela de impressão (por isso a main precisa estar publicada). No painel de cada departamento escolhe-se também quantos turnos ele precisa: 3 (24 h), 2 (sem noite) ou 1 (só de dia) |
| `oficina.py` | Lê `dados/oficina/ativos_oficina.xlsx` (anotações por departamento e linhas em azul), cruza com a escala de hoje e calcula cada escala com turma fixa por turno |
| `gerar_departamentos.py` + `departamentos_template.html` | Monta a página por departamento |
| `adm.py` → `adm.json` | Lê a planilha **ESCALA ADM** (`dados/adm/escala_adm.xlsx`, fora do git: tem nome e matrícula) e grava só as contagens por departamento e por central 24 h, que entram na aba ADM |
| `rural.py` → `rural.json` | Lê a planilha **ESCALA RURAL** (`dados/rural/escala_rural.xlsx`, fora do git) e grava só as contagens por departamento (pela coluna Dpto, o agrupamento novo) e do grupo que roda 3 turnos, que entram na aba Rural |
| `pessoal_frota.py` → `pessoal_frota.json` | Lê a planilha **Quadro motorista operador** (`dados/frota/quadro_motorista_operador.xlsx`, fora do git: tem nome, matrícula e CPF) e grava só quantos motoristas e operadores estão ativos, no total e por atividade |
| `fat.py` → `fat.json` + `FAT_lista_nominal.html` | Lê o **Painel FAT & Férias** do RH (`dados/fat/painel_fat_ferias.xlsx`, fora do git: tem nome, matrícula e idade) e grava em `fat.json` só a fila de quem tem direito ao FAT, com papel, atividade e mês do limite das férias, e as contagens dos que podem ter direito e dos contratos safra, para a página FAT. A lista com nomes (`FAT_lista_nominal.html`, e a cópia em `build/` para a página privada do claude.ai) fica fora do git |

### Operacional (frota 24 h)

Aba **Operacional** da mesma apresentação: os 84 equipamentos do *Detalhamento da frota* (lista em
`gerar_departamentos.FROTA`), cada um com 1 operador por turno nas 24 h, hoje em 5x1. Como o folguista roda entre
equipamentos, a conta é por equipamento: 5x1 = 3,6 pessoas por equipamento (hoje 303 operadores), 5x2 = 4,2 (+50),
4x2 de 8 h = 4,5 (+75), 12x36 e 3x3 = 4,0 (+33), 4x2 de 12 h = 3,0 (−51). Cada tipo fica fracionado e o total da frota
arredonda para cima. Quantidade e operadores por turno são editáveis na página.

A frente **Atividades de tratos** (pedido de 23/10/2026) entra com o turno e o folguista anotados de cada equipamento:
12 tratores de cata mato, 1 caminhão, 1 retroescavadeira e 1 uniport em 1 turno sem folguista; 3 tratores de
adubação em 3 turnos e 5 bombeiros em 2 turnos, com folguista; 3 tratores de manutenção da irrigação, 2 quadriciclos e
2 tratores de formiga no ADM. O hoje conta o mesmo regime na 5x1 (sem folguista e ADM, 1 por turno; com folguista,
1,2 por turno).

No quadro por equipamento, cada equipamento tem **turnos** (3 · 24 h, 2 · sem noite, 1 · só de dia ou ADM, que é
segunda a sexta, de dia, sempre em 5x2) e **folguista** (com ou sem). Sem folguista ficam só as turmas inteiras da
escala e o equipamento para na folga do operador: na 5x2 roda 71% dos dias e na 4x2 67%; na 12x36 e na 3x3 as duas
turmas se revezam e ele roda todo dia. Com 2 turnos as escalas de 12 h não se aplicam, como na oficina.

Na oficina, a opção **Todos na mesma escala** (na matriz e no painel do departamento) tira o reforço de 8 h em 5x2
das escalas de 12 h: o dia de 12 h cobre o maior entre manhã e tarde e a noite, o maior entre tarde e noite
(Caminhões / Veículos em 12x36: 6 no dia e 4 na noite, em vez de 4 + 4 e 2 de reforço).

A 4x2 de 12 h entra na simulação com **risco alto** (botão "4x2 12 h em todos" e opção de entrar na sugerida): no
cadastro, operadores de máquinas (335) e motoristas (250) são CLT urbano e, com a hora noturna reduzida, a noite de
12 h passa de 10 h contadas. Só vale com norma coletiva que troque a hora reduzida por adicional maior.

### Resumo para a diretoria

Página 14 (módulo Resumo, antes da Referência), para imprimir numa folha A4 e apresentar **como fica o quadro**: no topo,
quadro hoje, com as escalas escolhidas, a diferença e o rural no fim da safra; embaixo, uma coluna por área (oficina,
frota, ADM e rural) com cada departamento (na frota, cada tipo de equipamento), a escala escolhida e as pessoas hoje →
com a escala. Tudo sai das escolhas feitas na página (`resumoFinal()`). Na frota a base é de operadores necessários
(3,6 por equipamento na 5x1), não o quadro.

A capa traz o título **Estudo comparativo de escalas** e o símbolo da CRV em branco (`logo_crv_branco_datauri.txt`,
feito do logo azul: o azul vira branco e o fundo branco, transparente).

### Motoristas e operadores (painel da frota)

No topo do painel da frota, o resumo compara os ativos do Quadro motorista operador de 26/10/2026 (591: 328 operadores e
263 motoristas; 1 com rescisão anotada fica fora; o papel vem da função) com o que a frota precisa hoje (5x1), na escala
escolhida e em cada escala. O que sobra fica **sem atividade**. Contam como operador trator, plantadora, transbordo,
colhedora, retroescavadeira, quadriciclo, uniport, pá carregadeira, carregadeira, motoniveladora, escavadeira e rolo;
os outros equipamentos (caminhões, bombeiro, caçamba, prancha, IMAVE) contam como motorista (`papelEq` no script). Com
a escala escolhida: 245 operadores e 139 motoristas, 207 sem atividade. Muitos ativos estão em atividades que ainda não
estão na lista da frota (ônibus, pipas, comboio, caminhões e tratores da irrigação, máquinas de herbicida) e aparecem
como sem atividade; a lista por atividade abre em "Ativos por atividade".

### FAT (página 8)

Pedido de 28/10/2026: quem fica **sem atividade** na escala escolhida (os ativos do Quadro motorista operador menos o que
a frota precisa, como no painel da frota) pode ir para o **FAT**. A fonte é o **Painel FAT & Férias** do RH (27/10/2026):
320 com direito (efetivos, 12 meses ou mais, não aposentados), 199 operadores e 121 motoristas. A primeira lista
(*validar fat*, 237) deixava de fora os 83 com direito e férias vencidas. Vão os primeiros da fila de cada papel até a
necessidade; a fila põe a **colheita primeiro** (trator transbordo, colhedora e caminhão canavieiro, que param na
entressafra), depois as outras atividades, e dentro de cada grupo o **menor tempo de casa** primeiro. Com a escolhida dos
dados: 83 operadores e 124 motoristas sem atividade.

**Sem férias antes do FAT** (decisão de 28/10/2026). O FAT suspende o contrato (CLT art. 476-A): o tempo parado não
conta para as próximas férias, mas as férias já adquiridas têm de começar até o "Limite férias" do painel e não dá para
dar férias com o contrato suspenso. Como não haverá férias antes do FAT, quem vence a **segunda férias na safra de 2027**
(limite de abril a novembro de 2027) **não vai para o FAT**: voltaria e teria de tirar as férias na safra. Com a escolhida
dos dados ficam fora 107 dos 320 com direito (69 operadores e 38 motoristas). Vão 83 operadores (todos os que faltam) e
83 motoristas (todos os que podem ir); **41 motoristas ficam sem atividade e sem FAT**. Para reduzir: os 38 motoristas
fora pelas férias podem tirar as férias na entressafra, 15 motoristas podem ter direito e estão a conferir (60 anos ou
mais com aposentadoria a verificar, ou fora do levantamento) e 12 de contrato safra saem no fim da safra. Quem vai com o
limite até mar/27 (144) entra no FAT sem tirar as férias e o limite cai durante o FAT: confirmar com o DP se o prazo
fica parado durante o FAT, senão as férias vencem em dobro. (A página 9, Férias e FAT, saiu a pedido em 28/10/2026.)

A página só tem números e o botão **Lista nominal**, que abre a página privada do claude.ai (só para quem tem acesso;
o endereço, https://claude.ai/artifact/VYcCeLG9VKAHniU3EFaaLh, fica em `LISTA_PRIVADA`, no `fat.py`). Os nomes ficam na **lista nominal** (`FAT_lista_nominal.html`, gerada pelo `fat.py`; desde 28/10/2026 está no
repositório, que ficou privado, e o Render não a publica, ver `render-build.sh`; a cópia em `build/` é a página privada
do claude.ai): a mesma fila, com caixa para marcar quem vai (quem vence a segunda férias na safra fica sem caixa), a
necessidade de cada papel editável, contador, a situação das férias de cada um, a lista dos que podem ter direito e "Copiar
marcados" (cola no Excel). Na lista, a necessidade de partida é a da escolhida dos dados (`NEC_PADRAO` no `fat.py`); se a
escolhida mudar, vale o número da página FAT.

### ADM agrícola

Aba **ADM** (páginas 8 a 10: painel, quadro e matriz), no mesmo modelo da Operacional, a partir da planilha ESCALA
ADM de 26/10/2026: 97 pessoas em 24 departamentos. Quem está em 5x2 (85: 7 ADM e 78 sem cartão de ponto) fica de
segunda a sexta e não muda em nenhuma escala. As três centrais que rodam 24 h, hoje em 12x36 com 4 pessoas cada (1 por
turno), entram na simulação com turma fixa por turno, como a oficina: **Monitoramento** (em amarelo na planilha; 3 da
Central de Operações Agrícolas e 1 da Central de Rastreamento), **COA** (idem) e **COI** (4 da Central de Operações de
Irrigação). Na 12x36, na 3x3 e na 4x2 de 12 h fecham com as 4 de hoje; na 5x2 e na 4x2 de 8 h precisam de 6 (+2 cada,
+6 no total). Pessoas por turno, turnos e escala mudam na matriz e no painel de cada central. Para atualizar: trocar a
planilha em `dados/adm/`, rodar `python adm.py` e `python gerar_departamentos.py`.

### Rural

Aba **Rural** (páginas 11 a 13), no mesmo modelo do ADM (o código das duas abas é um só: `AREAS` no script da página),
a partir da planilha ESCALA RURAL de 26/10/2026: 163 pessoas, 101 trabalhadores rurais (contrato rural) e 62 auxiliares
agrícolas (CLT urbano). Os departamentos seguem a coluna Dpto da planilha (Capinópolis I recebe 8 da Zona 3;
Fertirrigação I, 5 da Irrigação Localizada). Os 145 marcados ADM ficam de segunda a sexta, em 5x2, e não mudam. Os 18
marcados **3t** (atreladores de cana picada e de pátio, lavagem de colhedora e apoio regiões, todos auxiliares agrícolas)
rodam 3 turnos nas 24 h, juntos em Atreladores Pátio: hoje 18 em 5x2, o que dá 4 por turno com turma fixa. Na 5x2
rotativa e na 4x2 de 8 h fecham com os 18; na 12x36 e na 3x3 sobram 2; na 4x2 de 12 h, 6. Para atualizar: trocar a
planilha em `dados/rural/`, rodar `python rural.py` e `python gerar_departamentos.py`.

Fim da safra: a planilha **quadro rural** (`dados/rural/quadro_rural.xlsx`, fora do git) tem os 197 ativos de hoje. Pela
matrícula, 163 estão na escala e **34 saem no fim da safra** (22 atreladores de cana picada, 4 queimadores, 3 da
irrigação localizada, 3 do apoio regiões, 1 da adubação pátio e 1 do pátio/outros): 197 − 34 = 163 depois da safra. As
3 linhas "CONTRATAR" da escala não entram (a conta é só com quem já está no quadro). Esses números ficam nos quadros do
painel Rural (o detalhe aparece ao passar o mouse).

### 5x2 com folga desencontrada

A apresentação mostra a 5x2 com **folga separada** (`analise.modelo_5x2_desenc`): as 2 folgas da semana ficam 3 ou 4
dias uma da outra e mudam de dia a cada semana, no máximo 4 dias seguidos e um domingo de folga a cada 3 ou 4 semanas.
A folga junta (`modelo_5x2`, 2 dias seguidos que andam um dia por semana, até 5 dias seguidos) fica no botão ao lado
para comparar. A quantidade de gente é a mesma nas duas: todo dia trabalham 5 das 7 turmas, então continua 4,2 por
equipamento.

### Oficina por departamento

`python gerar_departamentos.py`. Regra de cada departamento em `oficina.REGRAS`, tirada das anotações da planilha.
"Garantir por turno" começa na média de hoje arredondada (2,5 vira 2), no mínimo 1 nos departamentos 24 h; PCM e
caldeiraria com 1 por turno; Colhedoras de Cana com **5 frentes × 2 pessoas por turno** (10 por turno nas 24 h; frentes
em `gerar_departamentos.FRENTES_DEP`, editáveis no painel do departamento). Com turma fixa por turno: 5x2 em todos dá
+36; a escala sugerida (a de menos contratação dentro da lei) dá +24, quase toda em 12x36. A 4x2 de 12 h entra com risco alto (CLT urbano, noite com hora reduzida).

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
