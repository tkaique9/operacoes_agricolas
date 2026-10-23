"""Textos e tabelas do memorial e da planilha (premissas, pendências, riscos, normas, plano, orçamento).

Mesma identificação das pranchas: setores S1–S13, equipamentos Exx, interferências Ixx, divergências Dxx.
"""

PREMISSAS = [
    ("PR01", "Dimensões do galpão conforme o desenho CRV(MG)20102026 rev. 00: cotas escritas e, onde não há cota, "
             "medição das linhas do PDF (escala 1:50, precisão ≈ ±5 mm)."),
    ("PR02", "Volume médio: até 1 motor e 1 transmissão em desmontagem e 1 motor e 1 transmissão em montagem ao mesmo tempo."),
    ("PR03", "Conjunto mais pesado + dispositivo + cintas ≤ 2.000 kg (capacidade do pórtico e da talha do projeto). "
             "A confirmar com os pesos reais (P01)."),
    ("PR04", "Pórtico: truque das pernas 1.000 x 520 mm; gancho chega a 750 mm do eixo do trilho; folga de segurança de "
             "500 mm entre a perna em movimento e qualquer objeto fixo ou pessoa."),
    ("PR05", "Envelopes de reserva: lavadora 2.500 x 2.200 x 2.600 mm + carro extraível de 1.600 mm; cavalete com motor "
             "1.600 x 1.000 mm; bancadas 2.000 x 900 mm."),
    ("PR06", "Retífica (bloco, cabeçote, virabrequim, bielas), bomba e bicos injetores, turbocompressor e dinamômetro "
             "são serviços externos."),
    ("PR07", "Pré-lavagem do conjunto (lama, palha, óleo externo) feita fora do galpão, em lavador com separação de óleo."),
    ("PR08", "Nenhum teste com motor em funcionamento dentro do galpão; só testes funcionais (pré-lubrificação, "
             "estanqueidade, giro manual, torque)."),
    ("PR09", "Lavagem por processo aquoso; sem diesel, querosene ou solvente inflamável."),
    ("PR10", "Curso do pórtico limitado por batentes: pernas entre X 5.300 e X 19.250; trilhos de X 4.800 a X 19.600."),
    ("PR11", "Escritório existente mantido, usado como escritório técnico e sala de metrologia fina (com ar-condicionado)."),
    ("PR12", "Forro a ≈ 5,38 m do piso acabado (medido); luminárias e instalações aéreas entre ~4,6 e ~5,3 m."),
]

PENDENCIAS = [   # id, informação que falta, por que importa, quem levanta, bloqueia
    ("P01", "Pesos e dimensões máximos dos motores, transmissões e redutores atendidos (por frota e modelo).",
     "Capacidade do pórtico, dos cavaletes, da lavadora e das bancadas.", "Manutenção / engenharia", "Sim"),
    ("P02", "Dimensões do truque das pernas, aproximação e altura útil do gancho, curso e batentes do pórtico.",
     "Largura útil (Y 1.418–6.432) e alcance do gancho.", "Fornecedor do pórtico", "Sim"),
    ("P03", "Situação da obra: o galpão e a laje já foram executados?",
     "Define se trilhos, caixas de piso, dreno e contenção ainda podem ser embutidos; e se a Alt. 2 é viável.",
     "Engenharia / obras", "Sim"),
    ("P04", "O pórtico e a talha já existem na usina ou serão comprados? Documentação e inspeções.",
     "Capacidade real, adequação à NR-12 e prazo.", "Manutenção", "Sim"),
    ("P05", "D07 recolhimento do portão; D08 vão de 2.250 na lateral A; D09 vãos da vista posterior; "
            "D12 fechamento dos vãos superiores das laterais.", "Acessos, rotas de fuga, ventilação e poeira.",
     "Projetista", "Não"),
    ("P06", "Altura e forro do escritório; possibilidade de ar-condicionado (metrologia).", "Uso como sala de metrologia.",
     "Projetista", "Não"),
    ("P07", "Ponto de entrega de energia: tensão, potência disponível, distância.", "Alimentação da lavadora e do pórtico.",
     "Elétrica da usina", "Sim"),
    ("P08", "Ar comprimido: rede existente ou compressor novo; local do abrigo.", "Rede de ar e enroladores R1–R4.",
     "Manutenção / utilidades", "Não"),
    ("P09", "Destino do efluente da CX-1 (caminhão de sucção ou estação de tratamento da usina) e se a licença "
            "ambiental da usina contempla a atividade.", "Drenagem e conformidade ambiental.", "Meio ambiente", "Sim"),
    ("P10", "Existe lavador externo com caixa separadora de água e óleo para a pré-lavagem (PR07)?",
     "Evita levar lama e óleo para dentro do galpão.", "Manutenção / meio ambiente", "Não"),
    ("P11", "Detergente da lavadora (fornecedor, FDS, compatibilidade com alumínio, temperatura).",
     "Seleção da lavadora, contenção e armazenamento.", "Manutenção / SESMT", "Não"),
    ("P12", "Volume de serviço (motores e transmissões por mês) e tempo médio por OS.",
     "Confirma o número de posições de desmontagem e montagem.", "Manutenção / PCM", "Não"),
    ("P13", "Escopo interno x externo (retífica, injeção, turbos, ensaios de trinca).", "Ferramental e área de inspeção.",
     "Manutenção", "Não"),
    ("P14", "Vestiário, sanitários e lavatórios próximos (NR-24).", "Condições sanitárias para trabalho com óleo.",
     "SESMT", "Não"),
    ("P15", "Implantação: norte, acesso e área de manobra de caminhão/empilhadeira em frente ao portão.",
     "Recebimento e expedição.", "Projetista", "Não"),
    ("P16", "Divergências estruturais D01–D05 (blocos, volumes de concreto).", "Fundação dos pilares e do pórtico.",
     "Projetista estrutural", "Sim"),
    ("P17", "Enquadramento da oficina (NR-31 x NRs gerais) e inclusão no PGR da usina.", "Requisitos legais aplicáveis.",
     "SESMT", "Não"),
    ("P18", "Exigências do CBMMG para a edificação (saídas, extintores, iluminação de emergência, sinalização).",
     "Projeto de prevenção contra incêndio.", "Projetista de PCI", "Sim"),
]

ATIVIDADES = [   # atividade, onde, interno/externo
    ("Recebimento, abertura de OS, fotos e drenagem de fluidos", "S1", "Interno"),
    ("Desmontagem de motores e de transmissões/redutores", "S2, S3", "Interno"),
    ("Lavagem de peças (lavadora e tanque de peças pequenas)", "S4, S5", "Interno"),
    ("Inspeção visual e dimensional (folgas, desgastes, empenos com instrumentos)", "S6, S7, S11", "Interno"),
    ("Montagem, substituição de rolamentos, retentores, juntas, camisas*", "S8, S9", "Interno"),
    ("Testes funcionais: pré-lubrificação, estanqueidade, giro manual, torque", "S9, S10", "Interno"),
    ("Retífica de bloco, cabeçote, virabrequim e bielas; sedes e guias de válvula", "—", "Externo (premissa PR06)"),
    ("Bomba injetora, bicos/unidades injetoras, turbocompressor", "—", "Externo (premissa PR06)"),
    ("Ensaios de trinca (líquido penetrante/partículas magnéticas)", "—", "Externo ou interno a definir (P13)"),
    ("Teste com motor em funcionamento (dinamômetro)", "—", "Externo ou no equipamento (PR08)"),
]

NORMAS = [   # referência, tema, como se aplica aqui
    ("NR-01", "Gerenciamento de riscos ocupacionais (GRO/PGR)", "Incluir a oficina no inventário de riscos e plano de ação do PGR da usina."),
    ("NR-06", "EPI", "Luvas nitrílicas e térmicas, protetor facial, avental impermeável, óculos, protetor auricular, calçado."),
    ("NR-09", "Avaliação das exposições a agentes físicos, químicos e biológicos", "Ruído, névoa de óleo, produtos alcalinos, calor da lavagem."),
    ("NR-10", "Segurança em eletricidade", "QDO, área úmida, caixas de piso, bloqueio e etiquetagem; pessoal autorizado."),
    ("NR-11", "Transporte e movimentação de materiais", "Pórtico, talha, empilhadeira, acessórios de içamento; operadores capacitados."),
    ("NR-12", "Máquinas e equipamentos", "Apreciação de riscos e proteções da lavadora, prensa, pórtico/talha e demais máquinas."),
    ("NR-15", "Atividades insalubres", "Caracterização por exposição (óleos minerais, ruído) — avaliar."),
    ("NR-17", "Ergonomia", "Alturas de bancada, cavaletes rotativos, levantamento de peças; AEP/AET."),
    ("NR-20", "Inflamáveis e combustíveis", "Aplicável se houver armazenamento/manuseio acima dos limites da norma; diesel residual dos motores — verificar."),
    ("NR-23", "Proteção contra incêndios", "Junto com o regulamento estadual e as Instruções Técnicas do CBMMG."),
    ("NR-24", "Condições sanitárias", "Vestiário, lavatórios e sanitários para quem trabalha com óleo (P14)."),
    ("NR-25", "Resíduos industriais", "Óleo usado, lodo, banho esgotado, filtros e estopas contaminadas."),
    ("NR-26", "Sinalização e identificação de segurança", "Rotulagem e FDS dos produtos (GHS), sinalização de riscos."),
    ("NR-31", "Trabalho na agricultura (manutenção de máquinas agrícolas)", "Verificar enquadramento com o SESMT (P17)."),
    ("NR-33", "Espaços confinados", "Só se a CX-1 ou caixas exigirem entrada; projetar para limpeza por sucção, sem entrada."),
    ("NR-35", "Trabalho em altura", "Manutenção do pórtico (viga a 4,46 m), luminárias (~5 m) e forro."),
    ("ABNT NBR 5410", "Instalações elétricas de baixa tensão", "Projeto elétrico da oficina."),
    ("ABNT NBR 5419", "Proteção contra descargas atmosféricas", "Análise de risco e SPDA do galpão."),
    ("ABNT NBR ISO/CIE 8995-1", "Iluminação de ambientes de trabalho", "Níveis por tarefa (bancadas, metrologia) no cálculo luminotécnico."),
    ("ABNT NBR 8400", "Equipamentos de levantamento e movimentação de cargas", "Dimensionamento/verificação do pórtico ou ponte rolante."),
    ("ABNT NBR 10004", "Classificação de resíduos sólidos", "Classificar lodo, banho esgotado, estopas, filtros."),
    ("ABNT NBR 12235", "Armazenamento de resíduos perigosos", "Abrigo externo de resíduos."),
    ("ABNT NBR 17505", "Armazenamento de líquidos inflamáveis e combustíveis", "Somente se houver estoque desses líquidos."),
    ("ABNT NBR 14725", "Produtos químicos — FDS e rotulagem (GHS)", "Detergente, protetivo e demais produtos."),
    ("ABNT NBR 7195 / 13434", "Cores e sinalização de segurança / contra incêndio", "Demarcação de piso, faixas e placas."),
    ("CONAMA 430/2011", "Condições e padrões de lançamento de efluentes", "Nenhum lançamento sem tratamento e autorização (CX-1 estanque)."),
    ("CONAMA 362/2005", "Óleo lubrificante usado ou contaminado (OLUC)", "Coleta por coletor autorizado, com certificado."),
    ("Lei 12.305/2010 e MTR", "Política Nacional de Resíduos Sólidos; manifesto de transporte", "Destinação com MTR (sistema estadual/SINIR)."),
    ("Legislação de MG", "Licenciamento ambiental (SEMAD/FEAM, COPAM) e segurança contra incêndio (CBMMG)",
     "Verificar se a licença da usina cobre a lavagem; projeto de PCI aprovado no CBMMG."),
]

RISCOS = [   # id, perigo, onde, P, S, nível, controles, NR
    ("R01", "Queda de carga suspensa", "E01, todo o curso", 2, 3, "Alto",
     "Plano de içamento; acessórios inspecionados e identificados; capacidade marcada; ninguém sob carga; operador "
     "capacitado; pontos de içamento do fabricante.", "NR-11, NR-12"),
    ("R02", "Esmagamento entre a perna do pórtico e objetos/paredes", "Faixas S13", 2, 3, "Alto",
     "Faixa S13 zebrada e livre; batentes; alarme sonoro/luminoso na translação; comando com ação continuada; velocidade baixa.",
     "NR-12"),
    ("R03", "Tombamento de motor no cavalete ou queda de peças", "S2, S9, S3, S8", 1, 3, "Médio",
     "Cavalete com capacidade compatível e trava de giro; adaptadores corretos; calços; bancadas fixadas.", "NR-12"),
    ("R04", "Prensamento e projeção de peças na prensa", "E16", 2, 3, "Alto",
     "Proteção frontal; dispositivos de apoio corretos; procedimento; NR-12.", "NR-12"),
    ("R05", "Liberação de energia armazenada (molas, acumuladores)", "S3", 2, 3, "Alto",
     "Procedimento e ferramentas especiais do fabricante para compressão de molas e alívio de pressão.", "NR-12"),
    ("R06", "Contato com solução alcalina quente e vapor", "S5", 2, 2, "Médio",
     "Porta intertravada; espera de resfriamento; luvas, avental e protetor facial; lava-olhos E22; FDS no local.",
     "NR-06, NR-26"),
    ("R07", "Queimadura com peças quentes após a lavagem", "S5, S6", 2, 2, "Médio",
     "Luvas térmicas; tempo de resfriamento em S6; manuseio pelo pórtico.", "NR-06"),
    ("R08", "Escorregamento em piso molhado ou com óleo", "S1, S5, S6", 3, 2, "Alto",
     "Contenção e caimento em S5; revestimento antiderrapante; bandejas de gotejamento; kit de absorção; limpeza imediata.",
     "NR-01"),
    ("R09", "Choque elétrico", "S5, caixas de piso, QDO", 1, 3, "Médio",
     "DR; grau de proteção adequado; sem caixa de piso em S5; aterramento; bloqueio na manutenção; QDO restrito.", "NR-10"),
    ("R10", "Incêndio (óleo, estopas, combustível residual, produtos)", "Todo o galpão", 1, 3, "Médio",
     "Sem solvente inflamável; estopas em recipiente fechado; drenar combustível antes de entrar; extintores; PCI.",
     "NR-23, NR-20"),
    ("R11", "Ruído (impacto, ar comprimido, lavadora)", "S2, S3, S6", 2, 2, "Médio",
     "Avaliação quantitativa; bicos de sopro silenciosos; protetor auricular conforme avaliação.", "NR-09"),
    ("R12", "Névoa de óleo e partículas no sopro", "S6", 2, 2, "Médio",
     "Anteparo com coletor de névoa; bico de segurança com pressão limitada; óculos e protetor facial.", "NR-09, NR-06"),
    ("R13", "Sobrecarga postural e levantamento manual", "Bancadas e cavaletes", 3, 2, "Alto",
     "Cavaletes rotativos; bancadas na altura da tarefa; peças pesadas com auxílio mecânico; AEP.", "NR-17"),
    ("R14", "Atropelamento por empilhadeira", "Portão, S1, S10", 2, 3, "Alto",
     "Empilhadeira só na zona do portão; área de manobra demarcada; sinalização e sinaleiro; pedestre fora da rota.",
     "NR-11"),
    ("R15", "Contaminação de solo e água (efluentes, OLUC, lodo)", "S5, CX-1, resíduos", 2, 3, "Alto",
     "Contenção; CX-1 estanque; nenhuma ligação com pluvial/sanitário; MTR; destinadores licenciados; inspeção da CX-1.",
     "NR-25, CONAMA"),
    ("R16", "Queda de altura na manutenção do pórtico e das luminárias", "Pórtico, cobertura", 1, 3, "Médio",
     "Plataforma elevatória ou andaime; ancoragem; permissão de trabalho.", "NR-35"),
    ("R17", "Atmosfera perigosa na CX-1 ou caixas", "CX-1, CI", 1, 3, "Médio",
     "Projetar sem necessidade de entrada (limpeza por sucção externa); se houver entrada, aplicar NR-33.", "NR-33"),
    ("R18", "Aprisionamento na porta ou no carro da lavadora", "E07", 1, 2, "Baixo",
     "Intertravamento; trava do carro; apreciação de riscos do fornecedor.", "NR-12"),
]
ESCALA_RISCO = "P (probabilidade) e S (severidade) de 1 a 3, avaliação preliminar para validar no PGR; nível = P x S (≤2 baixo, 3–4 médio, ≥6 alto)."

MATRIZ_ICAMENTO = [   # critério, ponte rolante, pórtico, monovia, guindaste/pórtico com rodízios
    ("Cobertura", "Quase todo o piso (aproximação do gancho a confirmar)",
     "Faixa entre as pernas (Y 1.408–6.442) no curso dos batentes (X 5.800–18.750)", "Só a linha da viga",
     "Pontual; reposicionado à mão"),
    ("Ocupação de piso", "Nenhuma (vigas nas laterais)", "Trilhos + faixas das pernas (37,5 m² neste galpão)",
     "Nenhuma", "Pernas/base no ponto de uso"),
    ("Exigência estrutural", "Viga de rolamento, pilares e fundações capazes das reações das rodas; os pilares atuais "
     "foram projetados para a cobertura", "Autoportante; trilhos ancorados e nivelados no piso; laje e base dos "
     "trilhos a verificar", "Viga presa à estrutura ou a pórticos próprios — verificar", "Nenhuma"),
    ("Altura útil", "Viga + ponte + talha abaixo do forro (~5,38 m)", "Viga a 4,15–4,46 m (projeto)",
     "Semelhante à ponte", "Limitada pelo equipamento"),
    ("Flexibilidade do layout", "Alta", "Média (tudo na faixa útil)", "Baixa (estações em linha)",
     "Baixa para conjuntos pesados"),
    ("Riscos principais", "Carga suspensa sobre todo o vão", "Esmagamento nas pernas; trilho no piso",
     "Carga sobre o corredor", "Tombamento; esforço manual"),
    ("Manutenção", "Inspeção de vias, rodas, freios, fim de curso", "Inspeção de trilhos, rodas, batentes, viga",
     "Inspeção da viga e do carro", "Simples"),
    ("Situação neste galpão", "Exige revisão do projeto estrutural (Alt. 2)", "Já previsto no desenho (Alt. 1)",
     "Não recomendada (fluxo em U)", "Apoio pontual, complementar"),
]

PLANO = [   # etapa, entregáveis, responsáveis, dependências, riscos, critério de aprovação
    ("1 — Levantamento e validação",
     "Conferência de cotas e interferências em campo; respostas às pendências P01–P18; lista de equipamentos existentes; "
     "pesos e dimensões dos conjuntos.",
     "Manutenção (líder), engenharia, SESMT, meio ambiente", "Acesso ao local e aos manuais",
     "Dados incompletos de pesos e dimensões", "Pendências bloqueantes (P01–P04, P07, P09, P16, P18) respondidas"),
    ("2 — Estudo de layout",
     "Escolha entre Alt. 1 e Alt. 2; layout revisado com os dados reais; validação com quem opera.",
     "Engenharia de manutenção + operação da oficina", "Etapa 1",
     "Escolha da Alt. 2 sem estudo estrutural", "Layout assinado por manutenção, SESMT e engenharia"),
    ("3 — Especificação técnica",
     "Especificação da lavadora, cavaletes, bancadas, prensa e acessórios de içamento; dimensionamento preliminar de "
     "energia, ar, água, drenagem e exaustão; avaliação estrutural do içamento.",
     "Engenharia mecânica, elétrica e civil", "Etapa 2", "Lavadora maior que o envelope reservado",
     "Especificações compatíveis com os envelopes do layout (ou layout revisado)"),
    ("4 — Projetos complementares",
     "Projetos civil e de piso (contenção, caimento, canaleta, base dos trilhos), elétrico, hidráulico/drenagem, "
     "exaustão, PCI e demais exigidos; ARTs.",
     "Profissionais habilitados (CREA/CAU)", "Etapa 3",
     "Concretagem do piso antes dos embutidos", "Projetos aprovados internamente e, quando exigido, pelo CBMMG e órgão ambiental"),
    ("5 — Orçamento e contratação",
     "Quantitativos; cotações com escopo igual; orçamento com data-base, impostos, frete, instalação e contingência; "
     "priorização.", "Suprimentos + engenharia", "Etapa 4", "Cotações não comparáveis",
     "Orçamento aprovado pela gestão"),
    ("6 — Implantação e comissionamento",
     "Adequação do galpão; instalação; testes (estanqueidade de S5 e CX-1, elétrica, iluminância, pórtico com carga "
     "de teste); treinamento; procedimentos; liberação.",
     "Obras, fornecedores, manutenção, SESMT", "Etapa 5", "Liberação sem inspeção do pórtico e da lavadora",
     "Inspeções e testes aprovados; treinamentos registrados; PGR atualizado"),
]

ORCAMENTO = [   # categoria, itens (com quantidade quando sai do layout), base
    ("Obras civis e adequação do piso",
     "Contenção de S5 (≈ 13 m de lombada ou rebaixo de 9,8 m²); revestimento uretano-cimento em S5 (≈ 9,8 m² + rodapé); "
     "canaleta CN-1 (≈ 2,2 m); dreno enterrado (≈ 12 m) com caixa de inspeção; CX-1; base dos trilhos (2 x 14,8 m)."),
    ("Equipamentos de lavagem", "E07 lavadora (1); E08 tanque de peças pequenas (1); E10 mesa de escorrimento com "
     "anteparo e coletor de névoa (1)."),
    ("Bancadas e dispositivos", "E04/E17 cavaletes (2); E06/E15 bancadas (2); E12 bancada de inspeção (1); E16 prensa (1); "
     "E05 (2) e E14 (3) carrinhos; E19 (2), E20 (1), E21 (1) armários e racks; E24 desempeno e armário (1); E31 carro de "
     "testes (1); instrumentos de medição (lista a definir)."),
    ("Equipamentos de movimentação", "E01 pórtico e talha (já previstos — confirmar se existentes); acessórios de içamento "
     "(cintas, balancim, olhais); batentes e alarme de translação."),
    ("Instalações elétricas", "QDO (1); caixas de piso CP-01 a CP-07 (7); eletrodutos enterrados e derivações (≈ 24 m); "
     "alimentação aérea da lavadora (≈ 5,5 m); alimentação do pórtico até o trilho (≈ 10 m) + festão (≈ 14,5 m); "
     "luminárias LED (11, a confirmar no cálculo); iluminação de emergência (4); aterramento."),
    ("Instalações hidráulicas e drenagem", "Água fria até E22 e E07 (≈ 7,5 m); lava-olhos/chuveiro E22 (1)."),
    ("Ar comprimido", "Compressor ou derivação da rede existente (P08); rede interna (≈ 8 m); rede externa pela fachada A "
     "(≈ 38 m); enroladores R1–R4 (4); tratamento do ar (filtro/secador)."),
    ("Tratamento e destinação de efluentes", "CX-1 estanque; contrato de coleta (sucção) e destinação de banho, lodo e "
     "OLUC — ou ligação à estação da usina, se licenciada."),
    ("Ventilação e exaustão", "Duto EX-1 (≈ 6,6 m) com terminal na parede de fundos; ar-condicionado do escritório/metrologia."),
    ("Segurança e prevenção contra incêndio", "Extintores (4 indicativos); sinalização de emergência; iluminação de "
     "emergência; kits de derramamento (2)."),
    ("Sinalização e organização", "Demarcação de piso (corredores, setores, faixas S13 zebradas ≈ 37,5 m²); placas; painéis "
     "de ferramentas; etiquetas de OS e de status."),
    ("Projetos, montagem e comissionamento", "Projetos complementares e ARTs; montagem; testes e comissionamento; treinamento."),
]
ORCAMENTO_BASE = ("Sem cotação nesta etapa: nenhum valor foi estimado. Para cada categoria, cotar com o mesmo escopo e "
                  "registrar moeda (R$), data-base, impostos, frete, instalação e contingência. Quantidades marcadas com ≈ "
                  "saem do layout conceitual e devem ser refeitas nos projetos.")

CHECKLIST = [
    "Pesos e dimensões dos conjuntos conferidos contra a capacidade do pórtico, da talha, dos cavaletes e da lavadora.",
    "Dimensões do truque das pernas e aproximação do gancho confirmadas; faixa útil e curso recalculados.",
    "Cotas conferidas em campo (ou no projeto revisado), inclusive D01–D13.",
    "Decisão Alt. 1 x Alt. 2 tomada antes da concretagem do piso.",
    "Embutidos definidos: trilhos, eletrodutos, caixas de piso, dreno, CN-1, contenção, caimento, juntas seladas.",
    "Lavadora escolhida cabe no envelope (ou layout revisado) e tem utilidades confirmadas.",
    "Destino dos efluentes definido e compatível com a licença ambiental.",
    "Projetos complementares com ART; PCI conforme CBMMG; porta de fundos com abertura no sentido da fuga, se exigido.",
    "Apreciação de riscos NR-12 da lavadora, prensa e pórtico; PGR atualizado; treinamentos (NR-10, NR-11, NR-12).",
    "Teste de estanqueidade de S5 e da CX-1; teste de carga e inspeção do pórtico; medições de iluminância e elétricas.",
    "Sinalização de piso, identificação visual por categoria e procedimentos operacionais publicados.",
]
