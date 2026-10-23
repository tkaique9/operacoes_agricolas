"""Alternativas de layout da oficina de motores e transmissões (mm, origem como em galpao.py).

Fonte única para as pranchas, o DXF, o memorial e a planilha: setores, equipamentos, fluxos e
pontos de infraestrutura têm a mesma identificação em todos os arquivos.
Retângulos: (x0, x1, y0, y1). Dimensões de equipamento são envelopes de reserva, não especificação.
"""
import galpao as G

# ---------------------------------------------------------------- pórtico (Alternativa 1)
# premissas a confirmar com o fornecedor: truque da perna 1.000 (X) x 520 (Y); gancho a 750 do eixo do trilho
TRUQUE_X = 1000
FOLGA_SEG = 500                     # folga mínima entre a perna em movimento e objeto fixo/pessoa
PERNA_FAIXA = [(y - G.PERNA_LARGURA / 2, y + G.PERNA_LARGURA / 2) for y in G.TRILHOS_Y]  # 398-918 / 6932-7452
FAIXA_UTIL_Y = (PERNA_FAIXA[0][1] + FOLGA_SEG, PERNA_FAIXA[1][0] - FOLGA_SEG)            # 1.418 .. 6.432
PERNAS_X = (5300, 19250)            # curso proposto das pernas (batentes); hoje os trilhos vão de 75 a 19.925
GANCHO_X = (PERNAS_X[0] + TRUQUE_X / 2, PERNAS_X[1] - TRUQUE_X / 2)                      # 5.800 .. 18.600
GANCHO_Y = (G.TRILHOS_Y[0] + 750, G.TRILHOS_Y[1] - 750)                                   # 1.408 .. 6.442
TRILHOS_PROPOSTOS_X = (4800, 19600)

# ponte rolante (Alternativa 2): aproximações do gancho a confirmar com o fornecedor
PONTE_GANCHO_X = (700, 19300)
PONTE_GANCHO_Y = (900, 6950)

ZONAS = {   # cor de preenchimento, cor de borda, rótulo
    "suja": ("#f6dcc4", "#b8652a", "Zona suja"),
    "lavagem": ("#cfe3f5", "#2c6ea8", "Lavagem"),
    "limpa": ("#d5ecd6", "#2f7d3c", "Zona limpa"),
    "qualidade": ("#e4dcf1", "#6a4c9c", "Qualidade e expedição"),
    "apoio": ("#e6e6e6", "#666666", "Apoio"),
    "escritorio": ("#f2efe2", "#8a7d4f", "Escritório / metrologia"),
    "exclusao": ("#fff3b0", "#c9a400", "Faixa do pórtico (acesso restrito)"),
}

# ---------------------------------------------------------------- catálogo de equipamentos
# tag: nome, função, qtd base, dimensão de referência (mm), espaço operacional, utilidades,
#      requisitos de instalação, restrições de segurança, prioridade, dados a levantar
EQUIP = {
    "E01": ("Pórtico móvel 2 t com talha elétrica 2 t", "Içar e transportar motores, transmissões e peças pesadas entre as estações",
            1, "vão 6.534 (trilhos); viga W 310; altura 4.460", "faixa das pernas + 500 de folga; curso proposto X 5.300–19.100",
            "Energia trifásica (cabo festonado/enrolador ao longo do trilho)", "Trilhos nivelados e ancorados no piso; batentes de fim de curso; aterramento",
            "Plano de içamento; operador capacitado (NR-11/NR-12); inspeção periódica; ninguém sob carga suspensa", "Alta (já previsto no projeto)",
            "Peso máximo dos conjuntos + dispositivos; dimensão do truque das pernas; aproximação do gancho; altura útil do gancho"),
    "E02": ("Posição de recebimento (berço/palete com bandeja de contenção)", "Receber o conjunto da empilhadeira, conferir, fotografar e drenar fluidos",
            1, "1.200 x 1.400 (posição)", "800 livres ao redor; acesso da empilhadeira pelo portão", "—",
            "Bandeja estanque sob o conjunto; marcação de piso", "Calçar o conjunto; área de manobra da empilhadeira sinalizada", "Alta",
            "Dimensões do maior motor e da maior transmissão; meio de transporte de chegada"),
    "E03": ("Carrinho coletor de fluidos (óleo e arrefecimento)", "Drenar óleo e líquido de arrefecimento antes da desmontagem",
            1, "600 x 600", "Junto à posição de recebimento", "—", "Recipientes separados por tipo de fluido",
            "Óleo usado é resíduo perigoso (OLUC): coleta por empresa autorizada", "Alta", "Volume típico de óleo e de arrefecimento por conjunto"),
    "E04": ("Cavalete rotativo para motor — desmontagem", "Fixar e girar o motor durante a desmontagem",
            1, "1.600 x 1.000 (com motor)", "≥ 800 nas laterais de trabalho; içamento pelo pórtico", "—",
            "Piso nivelado; bandeja de gotejamento sob o motor", "Capacidade ≥ peso do maior motor (a confirmar); travamento da rotação",
            "Alta", "Peso e furação de fixação dos motores atendidos (adaptadores)"),
    "E05": ("Carrinho de peças — lado sujo (cestos por OS)", "Separar e identificar as peças desmontadas por ordem de serviço",
            2, "1.000 x 600", "Estacionam junto às estações de desmontagem", "—", "Cor própria do lado sujo; etiqueta da OS",
            "Capacidade de carga compatível; freio nos rodízios", "Alta", "—"),
    "E06": ("Bancada reforçada de desmontagem de transmissões com morsa", "Desmontar caixas, redutores, eixos e engrenagens",
            1, "2.000 x 900", "900 livres na frente", "Tomada 220 V; ar comprimido", "Fixação no piso se houver esforço de torque",
            "Molas e conjuntos sob tensão: dispositivos próprios; peças > 20 kg com auxílio mecânico", "Alta",
            "Peso da maior carcaça; ferramentas especiais por modelo"),
    "E07": ("Lavadora industrial de peças por aspersão (cabine, mesa giratória, carro extraível)", "Lavar blocos, cabeçotes, cárteres, carcaças e cestos de peças",
            1, "envelope de reserva 2.500 x 2.200 x 2.600 (A); carro extraível 1.600", "Zona de carga 1.600 na frente; acesso de serviço atrás pelo corredor",
            "Trifásica dedicada; água para enchimento/reposição; exaustão de vapor; esgotamento do banho", "Piso impermeável com contenção; nível; duto de exaustão até a parede de fundos",
            "Superfícies e solução quentes; produto alcalino; intertravamento da porta; bloqueio (LOTO) na manutenção", "Alta",
            "Maior peça (dimensões e peso) → diâmetro e carga da mesa; temperatura e química recomendadas pelos fabricantes dos componentes"),
    "E08": ("Lavadora de peças pequenas (tanque com bomba e pincel, solução aquosa)", "Pré-lavar e lavar peças pequenas, parafusos e componentes sensíveis",
            1, "1.000 x 700", "1.000 livres na frente", "Tomada 220 V", "Tampa com fechamento automático; sobre bacia",
            "Não usar diesel, querosene ou solvente inflamável", "Alta", "Volume de peças pequenas por OS"),
    "E10": ("Mesa de escorrimento e secagem com anteparo de sopro", "Escorrer, secar com ar e inspecionar visualmente as peças lavadas",
            1, "2.000 x 800", "800 livres na frente", "Ar comprimido (bico de segurança); bandeja para a água escorrida",
            "Anteparo contra respingo e névoa; grelha inox", "Bico de sopro com pressão limitada; protetor facial; nunca soprar na pele",
            "Alta", "—"),
    "E12": ("Bancada de inspeção dimensional com desempeno", "Medir blocos, virabrequins, eixos e carcaças depois de estabilizados à temperatura ambiente",
            1, "1.800 x 750 (desempeno 1.000 x 630 sobre a bancada)", "800 livres na frente", "Tomada 220 V; iluminação localizada",
            "Longe de respingo e de sopro; cobertura quando fora de uso", "—", "Alta",
            "Lista de instrumentos (faixa e resolução) e tolerâncias dos manuais"),
    "E14": ("Carrinho kit de OS — lado limpo", "Guardar peças aprovadas e peças novas por ordem de serviço até a montagem",
            3, "1.000 x 600", "Estacionam nas estações de montagem", "—", "Cor própria do lado limpo; cobertura contra poeira",
            "Freio nos rodízios", "Alta", "—"),
    "E15": ("Bancada de montagem de transmissões", "Montar caixas, redutores e conjuntos de eixos",
            1, "2.000 x 900", "900 livres na frente", "Tomada 220 V; ar comprimido", "Tampo limpo e resistente a óleo",
            "Peças > 20 kg com auxílio mecânico", "Alta", "—"),
    "E16": ("Prensa hidráulica de piso", "Prensar e sacar rolamentos, buchas e engrenagens",
            1, "700 x 800 (capacidade a definir)", "800 livres na frente", "—", "Fixação no piso",
            "Proteção frontal (NR-12); dispositivos de apoio adequados", "Média", "Maior força de prensagem exigida"),
    "E17": ("Cavalete rotativo para motor — montagem", "Montar o motor e fazer as verificações de montagem",
            1, "1.600 x 1.000 (com motor)", "≥ 800 nas laterais de trabalho", "Tomada 220 V; ar comprimido (CP próxima)",
            "Piso nivelado; bandeja de gotejamento", "Travamento da rotação; capacidade compatível", "Alta", "Idem E04"),
    "E18": ("Posição de inspeção final e expedição (skid/palete com capa)", "Inspeção final, registro e espera da retirada",
            1, "1.200 x 1.400 (posição)", "Acesso da empilhadeira pelo portão", "—", "Capa de proteção; etiqueta de liberação",
            "Conjunto calçado; área de manobra sinalizada", "Alta", "—"),
    "E19": ("Armário de ferramentas especiais e dispositivos", "Guardar ferramentas especiais, dispositivos e torquímetros",
            2, "1.000 x 600", "800 livres na frente", "—", "Fixado na parede", "—", "Média", "Inventário das ferramentas especiais existentes"),
    "E20": ("Rack de quarentena (peças reprovadas ou aguardando avaliação)", "Segregar peças reprovadas até decisão (descarte, garantia, recuperação)",
            1, "1.000 x 600", "800 livres na frente", "—", "Identificação vermelha", "—", "Média", "—"),
    "E21": ("Armário de produtos químicos com bacia de contenção", "Guardar detergente da lavadora, protetivo e demais produtos",
            1, "600 x 1.200", "800 livres na frente", "Ventilação", "Bacia ≥ volume do maior recipiente (a confirmar com a FDS/IT)",
            "Separar incompatíveis; FDS disponíveis; inflamáveis só em armário próprio", "Média", "Lista de produtos e volumes (FDS)"),
    "E22": ("Lava-olhos e chuveiro de emergência + lavatório", "Primeiro atendimento a respingo químico; higiene das mãos",
            1, "600 x 750", "Acesso livre e sinalizado", "Água; drenagem para a caixa de contenção", "A até ~10 s da lavadora",
            "Inspeção e acionamento periódicos", "Alta", "—"),
    "E23": ("Quadro de distribuição da oficina (QDO)", "Alimentar e proteger os circuitos da oficina",
            1, "800 x 250", "1.000 livres na frente (NR-10)", "Alimentação da rede da usina", "Fora da área úmida; fechadura",
            "Somente pessoal autorizado (NR-10)", "Alta", "Tensão e potência disponíveis no ponto de entrega"),
    "E24": ("Desempeno de granito e armário de instrumentos (no escritório)", "Medições de peças pequenas e guarda dos instrumentos",
            1, "1.000 x 630 (desempeno)", "800 livres", "Tomada 220 V; ar-condicionado", "Ambiente fechado, sem poeira e com temperatura estável",
            "—", "Média", "Instrumentos existentes e certificados de calibração"),
    "E26": ("Enrolador automático de mangueira de ar (15 m)", "Ar comprimido nas estações sem cruzar as faixas do pórtico",
            4, "fixado na parede/coluna", "—", "Ar comprimido", "Fora do curso das pernas do pórtico", "Mangueira com retorno controlado",
            "Média", "Demanda de ar das ferramentas"),
    "E29": ("Púlpito de registro (OS, fotos, etiqueta)", "Abrir a OS e registrar o estado de chegada",
            1, "500 x 550", "—", "Tomada 220 V", "Fora do curso das pernas", "—", "Baixa", "Sistema de OS utilizado"),
    "E31": ("Carro de testes funcionais (pré-lubrificação e teste de pressão)", "Pré-lubrificar e testar estanqueidade antes da liberação",
            1, "600 x 600", "Junto ao cavalete ou à posição final", "Ar comprimido; tomada 220 V", "—",
            "Teste com motor em funcionamento NÃO previsto no galpão", "Média", "Procedimentos de teste dos fabricantes"),
}

# ---------------------------------------------------------------- Alternativa 1 (recomendada)
A1_SETORES = [
    # id, nome, zona, retângulos, justificativa
    ("S1", "Recebimento, identificação e drenagem", "suja", [(17300, 19925, 1418, 3300)],
     "Junto ao portão, na metade do vão do lado A: a empilhadeira deixa o conjunto sem entrar no galpão e o pórtico o retira."),
    ("S2", "Desmontagem de motores", "suja", [(13400, 17300, 1418, 3300)],
     "Logo atrás do recebimento, no mesmo lado: o motor anda em linha reta e as peças seguem para o fundo, longe da área limpa."),
    ("S3", "Desmontagem de transmissões e redutores", "suja", [(11000, 13400, 1418, 3300)],
     "Bancada no mesmo lado sujo, entre a desmontagem de motores e a lavagem."),
    ("S4", "Lavagem de peças pequenas e espera de sujos", "lavagem", [(10000, 11000, 1418, 3300)],
     "Na entrada da lavagem: peças pequenas são lavadas aqui; as grandes esperam a lavadora nos carrinhos."),
    ("S5", "Lavagem industrial (lavadora e zona de carga)", "lavagem", [(5700, 10000, 1418, 3700)],
     "No fundo do lado A, na curva do fluxo em U: o duto de vapor e a drenagem saem pela parede de fundos sem cruzar o pórtico."),
    ("S6", "Escorrimento, secagem e inspeção visual", "lavagem", [(5700, 8100, 4900, 6432)],
     "Em frente à lavadora, já no lado limpo: a peça sai da lavadora e atravessa só o corredor."),
    ("S7", "Inspeção dimensional e estabilização", "limpa", [(8100, 10000, 4900, 6432)],
     "Entre a secagem e a montagem, perto do escritório/metrologia; peças medidas só depois de frias."),
    ("S8", "Montagem de transmissões e redutores", "limpa", [(10000, 13400, 4550, 6432)],
     "Lado B, espelhando a desmontagem: peças aprovadas chegam pelo lado limpo."),
    ("S9", "Montagem de motores", "limpa", [(13400, 17300, 4550, 6432)],
     "Lado B, mais longe da lavagem e perto da saída."),
    ("S10", "Inspeção final (CQ) e expedição", "qualidade", [(17300, 19925, 4550, 6432)],
     "Junto ao portão, na metade do vão do lado B, separada do recebimento pelo corredor central."),
    ("S11", "Escritório técnico e metrologia fina (sala existente)", "escritorio", [G.ESCRITORIO],
     "Única sala fechada do galpão: guarda instrumentos e faz as medições de precisão."),
    ("S12", "Ferramentaria, químicos, quarentena e emergência", "apoio", [(75, 4600, 1950, 3925)],
     "Fora do curso do pórtico, junto à porta de fundos e ao lado da lavagem."),
    ("S13", "Faixas de rolamento do pórtico (acesso restrito)", "exclusao",
     [(PERNAS_X[0], PERNAS_X[1], G.YI0, FAIXA_UTIL_Y[0]), (PERNAS_X[0], PERNAS_X[1], FAIXA_UTIL_Y[1], G.YI1)],
     "Faixa varrida pelas pernas + 500 mm de folga: nada fixo e ninguém parado com o pórtico em movimento."),
]

A1_EQUIP = [   # tag, retângulo, rótulo curto
    ("E02", (18000, 19200, 1700, 3100), "Recebimento"),
    ("E03", (17350, 17950, 1500, 2100), "Fluidos"),
    ("E29", (19425, 19925, 1420, 1900), ""),
    ("E04", (15000, 16600, 2000, 3000), "Cavalete desm."),
    ("E05", (13600, 14600, 1550, 2150), "Carro sujo"),
    ("E05", (13600, 14600, 2550, 3150), "Carro sujo"),
    ("E06", (11100, 13100, 1500, 2400), "Bancada transm."),
    ("E08", (10000, 11000, 1500, 2200), "Peças peq."),
    ("E07", (5800, 8300, 1450, 3650), "Lavadora"),
    ("E10", (5900, 7900, 5550, 6350), "Escorrimento/sopro"),
    ("E12", (8150, 9950, 5600, 6350), "Inspeção dim."),
    ("E15", (10200, 12200, 5450, 6350), "Bancada montagem"),
    ("E16", (12500, 13200, 5550, 6350), "Prensa"),
    ("E14", (13550, 14550, 4700, 5300), "Kit OS"),
    ("E14", (13550, 14550, 5650, 6250), "Kit OS"),
    ("E17", (14800, 16400, 4950, 5950), "Cavalete mont."),
    ("E31", (17350, 17950, 5700, 6300), "Testes"),
    ("E18", (18000, 19200, 4750, 6150), "Expedição"),
    ("E23", (75, 325, 2100, 2900), "QDO"),
    ("E22", (75, 675, 3100, 3850), "Lava-olhos"),
    ("E19", (900, 1900, 3325, 3925), "Ferram."),
    ("E19", (2000, 3000, 3325, 3925), "Ferram."),
    ("E20", (3050, 4000, 3325, 3925), "Quarentena"),
    ("E21", (4000, 4600, 2000, 3200), "Químicos"),
    ("E24", (1500, 2500, 4150, 4780), "Desempeno"),
]
A1_CARRO_LAVADORA = (8300, 9900, 1900, 3200)     # carro extraível, posição de carga
A1_PORTA_LAVADORA = 8300

# corredores e rotas (polilinhas e larguras), para cotas, fluxo e infraestrutura
A1_CORREDORES = [
    ("Corredor central", [(11000, 19925, 3300, 4550), (5700, 11000, 3700, 4900)]),
    ("Corredor transversal", [(4600, 5700, FAIXA_UTIL_Y[0], FAIXA_UTIL_Y[1])]),
    ("Rota até a porta de fundos", [(75, 4600, 700, 1950)]),
]

# ---------------------------------------------------------------- Alternativa 2 (ponte rolante)
A2_SETORES = [
    ("S1", "Recebimento, identificação e drenagem", "suja", [(17000, 19925, 75, 2950)], ""),
    ("S2", "Desmontagem de motores", "suja", [(12600, 17000, 75, 2950)], ""),
    ("S3", "Desmontagem de transmissões e redutores", "suja", [(10000, 12600, 75, 2950)], ""),
    ("S4", "Lavagem de peças pequenas e espera de sujos", "lavagem", [(9000, 10000, 75, 2950)], ""),
    ("S5", "Lavagem industrial (lavadora e zona de carga)", "lavagem", [(4600, 9000, 75, 2950)], ""),
    ("S6", "Escorrimento, secagem e inspeção visual", "lavagem", [(7200, 9600, 4550, 7775)], ""),
    ("S7", "Inspeção dimensional e estabilização", "limpa", [(4600, 7200, 4550, 7775)], ""),
    ("S8", "Montagem de transmissões e redutores", "limpa", [(9600, 12600, 4550, 7775)], ""),
    ("S9", "Montagem de motores (2 posições)", "limpa", [(12600, 17000, 4550, 7775)], ""),
    ("S10", "Inspeção final (CQ) e expedição", "qualidade", [(17000, 19925, 4550, 7775)], ""),
    ("S11", "Escritório técnico e metrologia fina (sala existente)", "escritorio", [G.ESCRITORIO], ""),
    ("S12", "Ferramentaria, químicos, quarentena e emergência", "apoio", [(75, 4600, 1950, 3925)], ""),
]
A2_EQUIP = [   # peças pesadas no centro (alcance do gancho); bancadas de peças leves junto às paredes
    ("E02", (17900, 19100, 1200, 2600), "Recebimento"),
    ("E03", (17200, 17800, 150, 750), "Fluidos"),
    ("E29", (19000, 19500, 400, 950), ""),
    ("E04", (14700, 16300, 1300, 2300), "Cavalete desm."),
    ("E05", (13200, 14200, 1100, 1700), "Carro sujo"),
    ("E05", (13200, 14200, 2000, 2600), "Carro sujo"),
    ("E06", (10300, 12300, 950, 1850), "Bancada transm."),
    ("E08", (9100, 9900, 400, 1100), "Peças peq."),
    ("E07", (4900, 7400, 700, 2900), "Lavadora"),
    ("E10", (7400, 9400, 5900, 6700), "Escorrimento/sopro"),
    ("E12", (5400, 7200, 6000, 6750), "Inspeção dim."),
    ("E15", (9800, 11800, 6000, 6900), "Bancada montagem"),
    ("E16", (11900, 12500, 6950, 7700), "Prensa"),
    ("E17", (13000, 14600, 5300, 6300), "Cavalete mont. 1"),
    ("E17", (15200, 16800, 5300, 6300), "Cavalete mont. 2"),
    ("E14", (13000, 14000, 7100, 7700), "Kit OS"),
    ("E14", (15400, 16400, 7100, 7700), "Kit OS"),
    ("E31", (17200, 17800, 7100, 7700), "Testes"),
    ("E18", (17900, 19100, 5300, 6700), "Expedição"),
    ("E23", (75, 325, 2100, 2900), "QDO"),
    ("E22", (75, 675, 3100, 3850), "Lava-olhos"),
    ("E19", (900, 1900, 3325, 3925), "Ferram."),
    ("E19", (2000, 3000, 3325, 3925), "Ferram."),
    ("E20", (3050, 4000, 3325, 3925), "Quarentena"),
    ("E21", (4000, 4600, 2000, 3200), "Químicos"),
    ("E24", (1500, 2500, 4150, 4780), "Desempeno"),
]
A2_CARRO_LAVADORA = (7400, 9000, 1150, 2450)
A2_CORREDORES = [
    ("Corredor central", [(4600, 19925, 2950, 4550)]),
    ("Rota até a porta de fundos", [(75, 4600, 700, 1950)]),
]

# ---------------------------------------------------------------- fluxos (Alternativa 1)
PORTAO_FORA = 21200
A1_FLUXOS = [
    # id, rótulo, cor, estilo, pontos
    ("F1", "Motor/transmissão sujo", "#b8652a", "-", [(PORTAO_FORA, 2400), (18600, 2400), (15800, 2500),
                                                     (12100, 2700), (9100, 2700), (8300, 2550)]),
    ("F2", "Peças limpas → montagem → saída", "#2f7d3c", "-", [(9100, 2550), (9100, 4250), (6900, 4250),
                                                              (6900, 5350), (9050, 5350), (9050, 4620),
                                                              (18600, 4620), (18600, 5450), (PORTAO_FORA, 5450)]),
    ("F3", "Peças pequenas", "#2c6ea8", "--", [(13600, 3200), (10500, 2350), (9300, 3000)]),
    ("F4", "Reprovadas → quarentena / relavagem", "#b3302d", ":", [(8150, 5150), (5150, 5150), (5150, 3280),
                                                                  (3520, 3280)]),
    ("F5", "Pessoas (entrada e rota de fuga)", "#444444", "-.", [(-900, 1350), (4400, 1350), (5150, 2000),
                                                                (5150, 4150), (PORTAO_FORA, 4150)]),
    ("F6", "Resíduos (fluidos, estopas)", "#7a5a00", ":", [(17650, 2100), (17650, 3500), (PORTAO_FORA, 3500)]),
]
A1_ETAPAS = [   # número, texto, ponto do rótulo
    (1, "Recebimento, OS, fotos", (18600, 2400)),
    (2, "Drenagem de fluidos", (17650, 1800)),
    (3, "Desmontagem", (15800, 2500)),
    (4, "Peças pequenas", (10500, 1850)),
    (5, "Lavadora", (7050, 2550)),
    (6, "Escorrimento e secagem", (6900, 5950)),
    (7, "Inspeção dimensional", (9050, 5975)),
    (8, "Montagem", (15600, 5450)),
    (9, "CQ e testes funcionais", (17650, 6000)),
    (10, "Expedição", (18600, 5450)),
]

# ---------------------------------------------------------------- infraestrutura (Alternativa 1)
CAIXAS_PISO = [   # tomadas 220/380 V em caixas de piso de embutir (tampa de ferro fundido)
    ("CP-01", 14400, 3150), ("CP-02", 12200, 3150), ("CP-03", 10500, 3150),
    ("CP-04", 11200, 4700), ("CP-05", 14400, 4700), ("CP-06", 9000, 5050), ("CP-07", 17600, 3350),
]
ENROLADORES = [("R1", 19800, 950), ("R2", 19800, 6900), ("R3", 4150, 4400), ("R4", 4600, 1950)]
LUMINARIAS = [(x, y) for x in (2400, 5700, 9000, 12300, 15600, 18900) for y in (2350, 5500) if (x, y) != (2400, 5500)]
EXTINTORES = [("PQS", 19700, 450), ("PQS", 4450, 1650), ("CO2", 900, 1950), ("PQS", 4400, 5200)]
REDE_ELETRICA = [  # traçados indicativos
    ("eletroduto enterrado sob o corredor", [(325, 2500), (1000, 2500), (5150, 3000), (5150, 4000),
                                           (17600, 4000), (17600, 3350)]),
    ("derivações às caixas de piso", [(14400, 4000), (14400, 3150)]),
    ("", [(12200, 4000), (12200, 3150)]), ("", [(10500, 4000), (10500, 3150)]),
    ("", [(11200, 4000), (11200, 4700)]), ("", [(14400, 4000), (14400, 4700)]), ("", [(9000, 4000), (9000, 5050)]),
    ("alimentação da lavadora (aérea, h ≈ 3,0 m)", [(325, 2650), (5800, 2650)]),
    ("alimentação do pórtico (festão ao longo do trilho B)", [(325, 2800), (-350, 2800), (-350, 7192),
                                                             (TRILHOS_PROPOSTOS_X[0], 7192), (PERNAS_X[1], 7192)]),
]
DUTO_EXAUSTAO = [(5900, 2600), (-700, 2600)]
AGUA = [(-1300, 3500), (375, 3500), (375, 3150), (5800, 3150)]
REDE_AR = [
    ("rede de ar interna (h ≈ 2,5 m)", [(-450, 4600), (-250, 4600), (-250, 3850), (4150, 3850), (4150, 4400)]),
    ("", [(4150, 3850), (4600, 3850), (4600, 1950)]),
    ("rede de ar externa pela fachada A (protegida)", [(-450, 4400), (-1720, 4400), (-1720, -500), (21000, -500),
                                                     (21000, 950), (19925, 950)]),
    ("", [(21000, 950), (21000, 6900), (19925, 6900)]),
]
CANALETA = (9850, 10000, 1500, 3650)                  # CN-1, grelha 150 mm
DRENO = [(9925, 1550), (5700, 1550), (75, 1550), (-1050, 1550), (-1050, 2200)]
CAIXA_CONTENCAO = (-1650, -450, 2200, 3200)           # CX-1, externa, sem saída (ou para tratamento licenciado)
CAIXA_INSPECAO = (5100, 1550)
COMPRESSOR = (-1750, -450, 4300, 5500)                # abrigo externo (a definir)
