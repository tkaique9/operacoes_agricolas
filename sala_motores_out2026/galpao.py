"""Galpão existente (desenho CRV(MG)20102026, rev. 00) em mm, com a fonte de cada medida.

Origem: eixo da parede de fundos = X 0 (portão em X 20.000); eixo da parede do lado da porta
de 1.000 = Y 0 ("lado A"); lado do escritório = Y 7.850 ("lado B").
Fonte "cota": valor escrito no desenho. "medido": tirado das linhas (extrair_desenho.py),
precisão de ±5 mm na escala 1:50. "premissa": adotado aqui, precisa de confirmação.
"""

DESENHO = {
    "numero": "CRV(MG)20102026",
    "titulo": "Galpão para motor a diesel — dimensão 7850 x 20.000 mm — conjunto",
    "revisao": "00 — emissão inicial (16/10/2026)",
    "autoria": "Desenho F. Monteiro; projeto, aprovação e revisão Ludeir; visto M. Coutinho",
    "local": "CRV Açúcar & Álcool — Capinópolis/MG",
    "formato": "A1, escala 1:50 (detalhes 1:10)",
    "vistas": "Vista superior — piso; vista superior — estrutura; secção S1-S1; fachada; vista posterior; "
              "vista lateral (sem título); detalhe do piso; detalhe da base dos pilares",
}

L, B = 20000, 7850            # entre eixos (cota)
PAREDE = 150                  # alvenaria centrada no eixo (cota "150", medido 152)
XI0, XI1 = PAREDE / 2, L - PAREDE / 2       # faces internas: 75 .. 19.925
YI0, YI1 = PAREDE / 2, B - PAREDE / 2       # 75 .. 7.775

PILARES_X = [0, 5000, 10000, 15000, 20000]  # cota 5.000 x 4
PILAR = (150, 180)            # perfil em planta, X x Y (medido)
BLOCO = 700                   # bloco em planta (medido 702; cotas 600 e 700; detalhe 600 x 600)
BLOCO_SOBRE_PISO = 150        # topo do bloco acima do piso acabado (medido no corte; detalhe "150")

ESCRITORIO = (75, 4075, 3925, 6925)        # faces externas x0, x1, y0, y1 (cotas 4.000 e 3.000; medido 4.074)
ESCRITORIO_PORTA = (5676, 6476)            # vão de 800 na face X 4.074 (cotas 450 / 800 / 1.750)
PORTA_FUNDOS = (850, 1851)                 # vão de 1.000 na parede X 0, abre para dentro
PORTAO = (1925, 5925)                      # 4.000 x 5.000 na parede X 20.000 (cotas 1.925/4.000/1.925)
ABERTURA_LATERAL = (17499, 19749)          # 2.250 na parede Y 0, elemento delgado de ~36 mm (a identificar)

TRILHOS_Y = (658, 7192)                    # eixos dos trilhos (vão 6.534, cota); trilho desenhado com 50 mm
TRILHOS_X = (75, 19925)                    # como desenhados: o comprimento interno todo
PERNA_LARGURA = 520                        # largura da perna/truque do pórtico no corte (medido ~513)

AREAS_DEMARCADAS = [                       # retângulos com borda de 100 mm (medido)
    ("ÁREA-02", 5360, 9641), ("ÁREA-02", 10360, 14639), ("ÁREA-01", 15359, 19638)]
AREAS_Y = (693, 7158)

ALTURAS = {   # acima do piso acabado, mm
    "topo dos blocos": (150, "medido no corte (detalhe: 150 acima do terreno)"),
    "base da viga do pórtico": (4150, "medido no corte (viga W 310 x 44,5)"),
    "topo da viga do pórtico": (4460, "cota 4.460"),
    "forro": (5380, "medido no corte; a cota 5.259 parte do topo do bloco + 16"),
    "beiral (fachada)": (5420, "1.959 + 1.630 + 1.685 = 5.274 acima do bloco, + 150"),
    "cumeeira": (6270, "6.121 acima do bloco, + 150"),
    "alvenaria das laterais": (2060, "1.910 + 16 acima do bloco (vista lateral)"),
}

PISO = "Laje de 100 mm, concreto 25 MPa, tela soldada 15 x 15 fio 3,8 a 30 mm do fundo, juntas a cada 3 m (40 mm)"
FUNDACAO = "Bloco 600 x 600 x 500 (30 MPa) sobre estaca Ø 250 x 4.000 (20 MPa); malha dupla CA-50 Ø 10"
PORTICO = "Pórtico móvel 2.000 kg sobre trilhos no piso, viga W 310 x 44,5 kg/m, talha elétrica 2.000 kg"

AREA_EIXOS = L * B / 1e6                                   # 157,00 m²
AREA_CONSTRUIDA = (L + PAREDE) * (B + PAREDE) / 1e6        # 161,20 m²
AREA_INTERNA = (XI1 - XI0) * (YI1 - YI0) / 1e6             # 152,85 m²
AREA_ESCRITORIO = (ESCRITORIO[1] - ESCRITORIO[0]) * (ESCRITORIO[3] - ESCRITORIO[2]) / 1e6  # 12,00 m²

# ---------------------------------------------------------------- achados da leitura do desenho
DIVERGENCIAS = [
    ("D01", "Bloco de fundação: a planta cota 600 num sentido e 700 no outro; a linha desenhada mede 700 x 700; "
            "o detalhe diz 600 x 600.", "Projetista estrutural"),
    ("D02", "Detalhe da base dos pilares: cotas parciais 150 + 250 = 400, cota total 500.", "Projetista estrutural"),
    ("D03", "Piso: volume indicado de 45 m³; a laje de 100 mm sobre 152,9 m² internos dá ≈ 15,3 m³.",
            "Projetista / orçamento"),
    ("D04", "Blocos: 10 x 0,6 x 0,6 x 0,5 = 1,80 m³ (2,45 m³ com 0,7 x 0,7); indicado 2,80 m³.", "Projetista estrutural"),
    ("D05", "Estacas: 10 x Ø 0,25 x 4,0 m = 1,96 m³; indicado 1,40 m³.", "Projetista estrutural"),
    ("D06", "Pé-direito livre até o forro não está cotado a partir do piso acabado (5.259 parte do topo do bloco). "
            "Medido no corte: ≈ 5,38 m.", "Projetista"),
    ("D07", "Portão de arraste de 4.000: só há 1.925 de parede de cada lado; o recolhimento da folha "
            "(duas folhas? trilho externo além da fachada?) não está representado.", "Projetista / fornecedor do portão"),
    ("D08", "Parede do lado A (Y 0), X 17.499 a 19.749: vão de 2.250 com elemento delgado (~36 mm). "
            "Não aparece em nenhuma elevação: porta, janela ou outro elemento?", "Projetista"),
    ("D09", "Vista posterior: vão tracejado de ~2,0 x 1,2 m (peitoril ~0,9 m) do lado do escritório e de "
            "~0,9 x 2,1 m do lado da porta, e um montante central no oitão; a planta não mostra janela nem pilar ali.",
            "Projetista"),
    ("D10", "Pórtico x escritório: a perna do pórtico (≈520 mm no corte, centrada no trilho Y 7.192) passa a "
            "≈10 mm da parede do escritório (Y 6.925) — ponto de esmagamento se o pórtico correr até o fundo.",
            "Projetista / fornecedor do pórtico"),
    ("D11", "As ÁREAS 01/02 (Y 693 a 7.158) cobrem as faixas por onde passam as pernas do pórtico; "
            "por isso foram lidas como demarcação de piso, não como base ou rebaixo.", "Projetista"),
    ("D12", "Vista lateral sem título e sem indicar o lado; os vãos superiores das laterais (1.630 + 1.706 acima "
            "da alvenaria) não têm fechamento indicado (tela, veneziana, policarbonato ou aberto?).", "Projetista"),
    ("D13", "Norte não indicado; altura e forro do escritório não indicados.", "Projetista"),
]

INTERFERENCIAS = [
    ("I01", "Pilares (perfil ~150 x 180) a cada 5 m nas laterais, com blocos de ~700 x 700 cujo topo fica "
            "150 mm acima do piso: ressalto de ~275 mm para dentro da face interna em 10 pontos."),
    ("I02", "Trilhos do pórtico embutidos no piso nos eixos Y 658 e Y 7.192, ao longo de todo o comprimento."),
    ("I03", "Pernas do pórtico (~520 mm): varrem uma faixa de cada lado; entre a perna e a parede sobram ~320 mm."),
    ("I04", "Viga do pórtico entre 4,15 e 4,46 m: dentro do curso do pórtico não cabe nada suspenso abaixo de "
            "~4,6 m (dutos, luminárias, linhas de ar, eletrocalhas)."),
    ("I05", "Escritório de 4,0 x 3,0 m no fundo, lado B, afastado 850 mm da parede lateral."),
    ("I06", "Forro a ~5,38 m: luminárias e qualquer instalação aérea ficam entre ~4,6 e ~5,3 m."),
    ("I07", "Juntas serradas do piso a cada 3 m: precisam de selagem química na área de lavagem."),
    ("I08", "Porta de fundos de 1,0 m abrindo para dentro, junto ao canto do lado A."),
    ("I09", "Portão de 4,0 x 5,0 m: único acesso para motores, transmissões e empilhadeira."),
    ("I10", "Vão de 2,25 m na lateral A junto ao portão (D08), dentro da faixa de varredura da perna do pórtico."),
]
