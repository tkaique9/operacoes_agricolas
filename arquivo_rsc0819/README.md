# Operações Agrícolas — Relatório Semanal

Apresentação semanal das operações agrícolas do Condomínio Agrícola Paulo Fernando e
Outros, no padrão **Indicadores Grupo Japungu**: mesmo visual, navegação e convenções do
`comparativo_moagem` (página única, uma tela por tópico, sem rolagem de documento).

É apresentada na **TV da diretoria**: as páginas não têm parágrafo de explicação nem faixa
de texto embaixo — o espaço vai para gráficos e números. Em tela de 1600 px ou mais
(TV 1920×1080), títulos, números e tabelas crescem pelo CSS, e o texto dos gráficos cresce
pela função `fz()` do script das páginas (cerca de 1,3× na TV, nunca menor que no
notebook).

## Como abrir

Basta abrir o `index.html` no navegador — não precisa de servidor nem de internet.
Para apresentar, use F11 (tela cheia). Navegação: setas ←/→, barra de espaço, as
bolinhas do rodapé, os botões ou o menu de módulos do topo. `?fundo=off` desliga as
partículas do fundo (projetor fraco).

## Atualizar a semana

1. Exporte o **RSC0819R — Relatório de Operações** da semana em `.xlsx`.
2. Coloque o arquivo na pasta `dados/` (os das semanas anteriores ficam lá).
3. Rode:

```bash
python atualizar_dados.py
```

O script lê todos os `.xlsx` de `dados/`, um por semana (o período sai do cabeçalho do
relatório), e regrava o bloco entre `DADOS-INICIO` e `DADOS-FIM` no `index.html`. Todo
o resto — totais, indicadores, gráficos e textos — é calculado no navegador. Com mais de
uma semana nos dados, aparece um seletor de semana no cabeçalho.

Operação nova sem rótulo curto cai na descrição do sistema: para dar um nome curto,
acrescente o código em `ROTULOS`, no topo do `atualizar_dados.py`.

O **tipo de aplicação** sai da descrição da operação pela primeira
palavra-chave encontrada, na ordem de `PALAVRAS_TIPO`: DRONE → Drone, UNIPORT →
Uniport, TRATOR → Trator, ARRASTAO → Arrastão, MANUAL → Manual, COCHO → Cocho,
VINHACA → Vinhaça enriquecida; sem palavra-chave, vira "Outros". Para forçar o tipo de
um código, acrescente-o em `TIPOS`.

## Páginas

| Módulo | # | Página | O que responde |
|---|---|---|---|
| — | — | Capa | Área da semana por dia e acumulada; área, OS, fazendas e aderência de dose |
| **Execução** | 1 | Resumo da semana | Oito números (área, média por dia, melhor dia, OS, fazendas, operações, aderência, OS para conferir) e a tabela por grupo com dias apontados e média por dia; clicar leva à página do número |
| | 2 | Área por grupo | Quanto foi executado em cada grupo, aberto por tipo de aplicação (drone, manual, Uniport…); clicar no tipo abre as operações dele |
| | 3 | Execução dia a dia | Área por dia, empilhada por grupo, contra a média dos dias apontados |
| | 4 | Área por fazenda | Ranking das fazendas, com filtro por grupo |
| **Por operação** | 5–8 | Uma página por grupo | Controle de pragas, Herbicida, Adubação e Irrigação localizada: área por dia, tipos e operações, fazendas e produtos do grupo |
| **Qualidade da aplicação** | 9 | Aderência de dose | Parte da área tratada abaixo, dentro e acima da tolerância, por operação |
| | 10 | Dose por produto | Dose real × recomendada de cada produto e o consumo da semana |
| | 11 | Pontos de atenção | OS com desvio acima de 20%, com o que conferir |

As páginas **por operação** são montadas pelo script a partir dos grupos que existem nos
dados (`montaPaginasGrupo`), como as páginas de frota do comparativo: um grupo novo num
relatório futuro ganha página sozinho, e a numeração das páginas e do índice se ajusta. Os
tipos de aplicação aparecem em tons da cor do grupo.

Em todas as páginas, clicar numa barra ou linha abre o **painel de detalhe** com a conta
por trás do número, e a linha de total fecha com o gráfico.

## Indicadores — as contas

O relatório traz uma linha por produto aplicado; a área é a do apontamento (OS × talhão ×
data) e aparece só na primeira linha de cada apontamento.

| Indicador | Conta |
|---|---|
| Área executada | Σ área dos apontamentos (cada apontamento conta uma vez, não uma vez por produto) |
| Média por dia | área ÷ dias com apontamento (no grupo: só os dias em que o grupo apontou) |
| Melhor dia | o dia de maior área executada |
| Dose real | Σ quantidade aplicada ÷ Σ área tratada |
| Desvio de dose | Σ quantidade ÷ Σ (área × dose recomendada) − 1 |
| Aderência de dose | área tratada com desvio dentro da tolerância ÷ área tratada |

- **Área tratada é por produto:** uma OS de 13 ha com 4 produtos conta 4 × 13 ha na
  aderência. Cada aplicação pesa pela sua área — nunca é média de percentuais.
- **Tolerância:** ±5% (padrão) ou ±10%, no alternador das páginas 4 e 5. A capa
  acompanha a escolha.
- **Pontos de atenção:** entra a OS com algum produto mais de 20% fora da dose
  (`ATENCAO`, no script das páginas). A leitura sugerida sai de três regras:
  OS na mesma fazenda e dia com desvio no sentido oposto (produto lançado na OS
  errada?); quantidade redonda (múltiplo de 5) com dose acima do dobro (total rateado
  pela área?); senão, abaixo ou acima da recomendação.

## Convenções mantidas do `comparativo_moagem`

- Mesmo CSS base, tema (Archivo, Source Serif 4, IBM Plex Mono, fundo de partículas),
  cromo, painel de detalhe e helpers de SVG, copiados de lá.
- Cor de destaque por módulo: dourado (Execução) e azul (Qualidade da aplicação).
- Cor fixa por grupo de operação em todas as telas: Controle de pragas azul `#2E86C1`,
  Herbicida oliva `#8A6D14`, Adubação navy `#12224A`, Irrigação localizada verde-água
  `#3FA79C`. Grupo novo entra no fim, com cor de reserva.
- Dose: dentro da tolerância em verde, abaixo em âmbar, acima em vermelho — sempre com o
  número escrito ao lado.

## Pontos em aberto

- **Meta × realizado:** o relatório traz "Área a executar da operação", mas com valores
  negativos em alguns talhões; falta confirmar o que o campo significa antes de usar.
- **Códigos 1xxxx × 2xxxx:** a mesma operação aparece com os dois prefixos (ex.: 14330 e
  24330, Uniport barra total). Se for cana planta × soca, dá para abrir a área por ciclo.
- **Controle de formiga:** a dose recomendada de Bifentol é 0,001 L/ha; com quantidades
  tão pequenas o arredondamento do apontamento pesa no desvio.
- **Comparativo semana a semana:** entra quando houver a segunda semana em `dados/`.
