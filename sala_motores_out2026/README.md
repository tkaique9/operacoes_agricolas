# Oficina de motores e transmissões com lavagem — estudo de layout (out/2026)

Estudo conceitual de layout para a sala de motores, transmissões e lavagem de componentes no galpão do desenho
**CRV(MG)20102026 rev. 00** ("Galpão para motor a diesel", 7.850 x 20.000 mm, CRV Açúcar & Álcool, Capinópolis/MG).
Não é projeto executivo: serve para validação interna, levantamento em campo, seleção de equipamentos e orçamento.

| Arquivo | O que é |
|---|---|
| `Pranchas_Sala_Motores_OUT2026.pdf` | 7 pranchas A3 vetoriais: SM-01 existente e interferências; SM-02 planta proposta (Desenho 1); SM-03 cotas e corte (Desenho 2); SM-04 fluxo e fluxograma (Desenho 3); SM-05 lavagem (Desenho 4); SM-06 infraestrutura (Desenho 5); SM-07 Alternativa 2 e comparação |
| `Planta_Sala_Motores_OUT2026.dxf` | CAD (DXF R2010, mm, escala 1:1) com camadas: existente, Alternativa 1 com cotas, fluxos e infraestrutura, Alternativa 2 |
| `Memorial_Tecnico_Sala_Motores_OUT2026.pdf` | Memorial técnico A4 (30 páginas), capítulos 1–18 |
| `Quadros_Sala_Motores_OUT2026.xlsx` | Quadro de áreas, equipamentos e posições, riscos, pendências, premissas, normas, plano, orçamento (a cotar), verificação, checklist |
| `build/SM-0x.svg`, `build/SM-0x.png` | Cada prancha em SVG (editável) e PNG |

## Como refazer

O PDF original fica em `dados/SALA DE MOTORES.pdf` (não sobe para o git). Na ordem:

```
python extrair_desenho.py   # lê as linhas do PDF, confere as cotas, grava desenho_original.json
python verificacao.py       # quadro de áreas e verificações de consistência (só imprime)
python pranchas.py          # pranchas PDF/SVG/PNG
python exportar_dxf.py      # planta em DXF
python memorial.py          # memorial PDF (Edge headless) e planilha
```

Fonte única dos números: `galpao.py` (galpão existente, com a fonte de cada medida), `layout.py` (setores, equipamentos,
fluxos e infraestrutura das duas alternativas) e `conteudo.py` (premissas, pendências, riscos, normas, plano, orçamento).
Mudou uma posição em `layout.py`, rodar de novo `pranchas.py`, `exportar_dxf.py` e `memorial.py`.

## Resultado

- Medidas tiradas das linhas do PDF batem com as cotas do desenho (diferença ≤ 2 mm). Forro ≈ 5,38 m do piso (medido).
- O pórtico de 2 t sobre trilhos do projeto deixa 5,01 m de largura útil (Y 1.418–6.432); as faixas das pernas ocupam
  37,47 m² (24,5% da área interna).
- **Recomendada — Alternativa 1:** pórtico mantido, curso limitado por batentes (pernas X 5.300–19.250), fluxo em U
  (lado A sujo, lado B limpo, lavagem no fundo do lado A). Setores produtivos: 53,76 m².
- Alternativa 2 (ponte rolante): 93,48 m² de setores produtivos (+74%) e 2ª posição de montagem, mas exige viga de
  rolamento, pilares e fundações novos.
- Trilhos, caixas de piso, dreno, canaleta e contenção são embutidos: decidir **antes da concretagem**.
- Pendências que bloqueiam execução: P01 pesos e dimensões, P02 geometria do pórtico, P03 situação da obra,
  P04 origem do pórtico, P07 energia, P09 destino dos efluentes, P16 divergências estruturais, P18 CBMMG.

## Limitações dos formatos

- **PDF**: vetorial, A3; plantas em 1:75 e detalhe da lavagem em 1:40. Impresso reduzido ou ampliado, vale a escala gráfica.
- **DXF**: abre no AutoCAD, BricsCAD, DraftSight, LibreCAD e QCAD; para DWG, abrir no AutoCAD e "Salvar como". Origem no
  cruzamento dos eixos 1 e A; Alternativa 1 em Y 0, existente deslocado +16.000 e Alternativa 2 −16.000. As camadas
  `FLUXOS`, `INFRA-*` e `00-ORIGINAL-PDF` vêm desligadas. Cotas são associadas ao estilo `SM`; textos em MTEXT (Arial).
- **SVG**: textos ficam como texto em Arial Narrow; sem essa fonte, o programa substitui por outra.
- Envelopes de equipamento são reservas de espaço, não especificação; itens em magenta dependem de confirmação.
