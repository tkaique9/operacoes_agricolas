"""Gera o relatório HTML dinâmico (simulador) a partir da análise.

python gerar_html.py  ->  Relatorio_Escalas_EC141_OUT2026.html   (abre direto no navegador, sem internet)
                          build/artifact_escalas.html             (mesma página, sem o esqueleto <html>, para publicar)
Só entram números agregados (função x área x horário). Nomes e matrículas ficam na planilha.
"""
import json
from pathlib import Path

import analise
from formatos import cargo, horario

AQUI = Path(__file__).parent
TEMPLATE = AQUI / "relatorio_template.html"
LOCAL = AQUI / "Relatorio_Escalas_EC141_OUT2026.html"
ARTIFACT = AQUI / "build" / "artifact_escalas.html"

GRUPOS = [
    ("5x1 turnos 24 h", "bad", "Muda: 5x1 → escala nova dos turnos 24 h"),
    ("6x1 operacional (dia)", "bad", "Muda: 6x1 → 5x2"),
    ("ADM com sábado (6x1)", "bad", "Muda: sai o sábado (5x2, 40 h)"),
    ("ADM sábado compensado (5x2)", "warn", "Ajusta horas: 44 h → 40 h"),
    ("12x36 (códigos Portaria)", "ok", "Mantém, com norma coletiva"),
    ("Afastado (INSS)", "off", "Fora da conta; entra ao voltar"),
    ("Sem registro no revezamento", "off", "Verificar a escala real"),
    ("Férias/atestado sem escala de origem", "off", "Verificar a escala de origem"),
]

ROTULOS = {
    "5x1": ([f"Turma {x}" for x in "ABCDEF"], "turmas por horário hoje", "5 de 6 turmas trabalham por dia"),
    "5x2": ([f"Turma {x}" for x in "ABCDEFG"], "turmas por horário (A–F de hoje + G nova)", "5 de 7 turmas trabalham por dia"),
    "4x2": ([f"Turma {x}" for x in "ABC"], "turmas em cada horário de 8 h", "2 de 3 turmas trabalham por dia em cada horário"),
    "12x36": (["A · dia", "B · dia", "C · noite", "D · noite"], "turmas: 2 de dia e 2 de noite", "1 turma de dia e 1 de noite por dia"),
    "3x3": (["A · dia", "B · dia", "C · noite", "D · noite"], "turmas: 2 de dia e 2 de noite", "1 turma de dia e 1 de noite por dia"),
    "4x2_12": (["A · dia", "B · dia", "C · dia", "D · noite", "E · noite", "F · noite"], "turmas: 3 de dia e 3 de noite",
               "2 turmas de dia e 2 de noite por dia"),
}

# o que a fiscalização confere: [tom, texto] por sugestão
IDS = ["5x2", "4x2", "12x36", "3x3", "4x2_12", "4x2_12r"]


def linha(item, *vals):
    return {"item": item, "v": dict(zip(IDS, vals))}


OK, W, B = "ok", "warn", "bad"
CHECAGENS = [
    linha("2 folgas por semana (EC 141)", [OK, "2"], [OK, "2 a 3"], [OK, "3 a 4"], [OK, "3 a 4"], [OK, "2 a 3"], [OK, "2 a 3"]),
    linha("Até 40 h normais por semana", [OK, "39h10"], [OK, "36h33"], [W, "38h30 na média, semanas de 44 h"],
          [W, "38h30 na média, semanas de 44 h"], [B, "37h20 + 9h20 extras"], [B, "37h20 + 9h20 extras"]),
    linha("Duração do dia", [OK, "7h50"], [OK, "7h50"], [W, "11 h, art. 59-A"], [W, "11 h, só com compensação coletiva"],
          [B, "10 h: teto de 8 h + 2 h extras"], [B, "10 h: teto de 8 h + 2 h extras"]),
    linha("Hora extra fixa na escala", [OK, "não"], [OK, "não"], [OK, "não"], [OK, "não"], [B, "2 h em todo plantão"],
          [B, "2 h em todo plantão"]),
    linha("Atividade insalubre (CLT art. 60)", [OK, "não se aplica"], [OK, "não se aplica"], [OK, "dispensa a licença"],
          [W, "licença do MTE ou norma coletiva"], [B, "licença do MTE ou norma coletiva"], [B, "licença do MTE ou norma coletiva"]),
    linha("Norma coletiva", [OK, "não precisa"], [OK, "não precisa"], [W, "obrigatória"], [W, "obrigatória"],
          [W, "acordo de prorrogação; na prática, coletivo"], [W, "acordo de prorrogação; na prática, coletivo"]),
    linha("Margem para atraso na troca de turno", [OK, "até 2 h extras"], [OK, "até 2 h extras"], [W, "pouca"], [W, "pouca"],
          [B, "nenhuma"], [B, "nenhuma"]),
    linha("Noite de trabalhador urbano (hora reduzida)", [W, "conferir na norma coletiva"], [W, "conferir na norma coletiva"],
          [W, "conferir na norma coletiva"], [W, "conferir na norma coletiva"], [B, "passa de 10 h contadas"], [B, "passa de 10 h contadas"]),
    linha("Fadiga (NR-1, NR-12, NR-31)", [OK, "8 h"], [OK, "8 h"], [W, "12 h, dia sim, dia não"], [W, "3 plantões de 12 h seguidos"],
          [B, "4 plantões de 12 h seguidos"], [B, "4 plantões de 12 h seguidos"]),
]

DET_8H = {
    "olha": ["Ponto com as 2 folgas de cada semana registradas.",
             "Trabalho aos domingos: atividade contínua e escala de revezamento com domingo de folga periódico (para mulheres, a cada 15 dias, CLT art. 386).",
             "Hora noturna e intervalo de 30 min conforme a norma coletiva da categoria."],
    "fazer": ["Manter as turmas fixas por horário. Alternar dia e noite transforma a escala em turno ininterrupto de revezamento, que tem limite de 6 h (CF art. 7º, XIV).",
              "Publicar a escala com antecedência e igualar o tamanho das turmas na passagem.",
              "Hora extra só eventual, dentro de 2 h por dia."],
}
DET_412 = {
    "veredito": "Não é proibida, mas é a sugestão com mais chance de auto de infração e de passivo trabalhista.",
    "olha": [
        "10 h de trabalho é o teto da CLT (8 h + 2 h extras, art. 59). Atraso na troca de turno, no abastecimento ou na espera do ônibus passa do limite: infração por empregado e por dia.",
        "Hora extra fixa: 46h40 por semana em média, contra 40 h da EC 141. Escala que já nasce com hora extra pode ser vista como forma de contornar o limite (CLT art. 9º), com risco de ação do Ministério Público do Trabalho.",
        "Atividade insalubre: prorrogar a jornada todo dia exige licença prévia do MTE (CLT art. 60) ou cláusula em norma coletiva (art. 611-A, XIII). O TST já invalidou escala de 12 h em atividade insalubre por falta de licença.",
        "Intervalo de 2 h: um só, registrado e livre. Se a pessoa fica na frente esperando, é tempo à disposição (CLT art. 4º) e o plantão vira 12 h de trabalho.",
        "Turno da noite de trabalhador urbano: com a hora noturna reduzida, 19:00–07:00 conta cerca de 11 h.",
        "4 plantões de 12 h seguidos operando máquina ou dirigindo: a fadiga precisa estar no PGR (NR-1) e na análise de risco (NR-31 e NR-12).",
        "Hora extra habitual entra em 13º, férias, FGTS e DSR. Se for cortada depois de 1 ano, gera indenização (Súmula 291 do TST).",
    ],
    "fazer": [
        "Parecer jurídico antes de implantar.",
        "Só nas frentes agrícolas 24 h (trabalhador rural, hora noturna de 60 min). Na indústria, não.",
        "Acordo coletivo específico com a prorrogação de 2 h, a insalubridade e o intervalo de 2 h.",
        "Ponto eletrônico com alerta antes de 10 h e troca de turno no campo, sem espera.",
        "Piloto de 2 a 3 meses em uma frente, medindo horas, acidentes e produção.",
        "Com rendição a máquina roda 24 h, mas o quadro volta a 3,6 por frota e a folha passa a da 5x2.",
    ],
}
RISCO_DET = {
    "5x1": {"veredito": "Proibida a partir de 15/12/2026: dá 1,17 folga por semana e, no turno de 7h50, 45h42 de média.",
            "olha": ["Menos de 2 folgas por semana.", "Média acima de 42 h (2026) e de 40 h (2027)."],
            "fazer": ["Trocar até 15/12/2026 por uma das sugestões."]},
    "5x2": {"veredito": "Escala que a EC 141 toma como regra. Não depende de norma coletiva para existir.", **DET_8H},
    "4x2": {"veredito": "Cumpre a lei com folga: 2 a 3 folgas por semana e 36h33 de média. Custa mais gente.", **DET_8H},
    "12x36": {"veredito": "Permitida pela CLT (art. 59-A) como regime diferenciado, desde que haja norma coletiva escrita nas regras novas.",
              "olha": ["Acordo ou convenção coletiva válidos depois de 15/12/2026; as cláusulas antigas incompatíveis perdem efeito.",
                       "Intervalo de 1 h gozado. Sem ele são 12 h por plantão e 42 h por semana, acima das 40 h de dez/2027.",
                       "Plantão que passa de 12 h por atraso na troca descaracteriza a escala.",
                       "Fadiga de 12 h operando máquina ou dirigindo à noite (PGR, NR-1)."],
              "fazer": ["Usar nas frentes agrícolas 24 h, com troca de turno no campo.", "Turmas fixas de dia e de noite.",
                        "Ponto com alerta perto de 12 h."]},
    "3x3": {"veredito": "Mesmo quadro da 12x36, com 3 plantões seguidos e 3 dias de folga. Não tem a base legal própria da 12x36: depende de compensação por norma coletiva.",
            "olha": ["Semanas com 44 h só valem com compensação em acordo coletivo (média de 38h30).",
                     "Em atividade insalubre, licença prévia do MTE (CLT art. 60) ou cláusula coletiva (art. 611-A, XIII). A dispensa do art. 60 vale só para a 12x36.",
                     "3 noites de 12 h seguidas: fadiga no PGR (NR-1)."],
            "fazer": ["Preferir a 12x36, que tem base legal própria.", "Se usar, acordo coletivo específico e turmas fixas de dia e de noite."]},
    "4x2_12": DET_412,
    "4x2_12r": {**DET_412, "veredito": "Mesmo risco da 4x2 de 12 h: a rendição resolve a máquina parada, não a jornada."},
}

LEI_TAB = [
    ["Jornada semanal máxima", "44 h", "42 h", "40 h"],
    ["Jornada diária normal", "8 h", "8 h (8h24 em 5 dias com norma coletiva)", "8 h"],
    ["Descanso semanal remunerado", "1 dia", "2 dias, um de preferência no domingo", "2 dias"],
    ["Escala diferente da 5x2", "acordo ou convenção", "só com norma coletiva: média de 2 folgas por semana no mês e 1 folga por semana", "idem"],
    ["Salário e piso", "—", "sem redução", "sem redução"],
    ["Cláusulas coletivas incompatíveis", "valem", "perdem efeito", "—"],
]

FONTES = [
    ("Planalto: EC 141/2026", "https://www.planalto.gov.br/ccivil_03/constituicao/emendas/emc/emc141.htm"),
    ("Agência Brasil: promulgação", "https://agenciabrasil.ebc.com.br/politica/noticia/2026-10/congresso-promulga-emenda-do-fim-da-6x1-escala-5x2-comeca-em-60-dias"),
    ("Agência Câmara: texto aprovado", "https://www.camara.leg.br/noticias/1277141-camara-aprova-em-dois-turnos-fim-da-escala-6x1-com-jornada-maxima-de-40-horas-semanais"),
    ("Contábeis: calendário da EC 141", "https://www.contabeis.com.br/noticias/80151/emenda-141-muda-jornada-e-garante-mais-descanso/"),
    ("Henrique Foreis: regime compensatório", "https://www.henriqueforeis.com.br/blog/fim-da-escala-6x1-emenda-constitucional-141.html"),
    ("TRT-4: prorrogação de 2 h (art. 59)", "https://periodicos.trt4.jus.br/revista/article/download/905/809/1450"),
    ("TST: 12 h sem licença em atividade insalubre", "https://www.lex.com.br/tecnica-de-enfermagem-recebera-horas-extras-por-jornada-12x36/"),
]


def dados(A):
    T = A["turno24"]
    G = A["grupos"]
    lei = A["lei"]
    cen = [{k: c[k] for k in ("id", "nome", "curto", "turnos", "trab", "ciclo", "h_trab", "he_dia", "rend", "turmas",
                              "parada", "risco", "horario", "obs", "h_sem", "he_sem", "folgas_sem")} for c in A["cenarios"]]
    grades = {"5x1": [["T" if (d + t) % 6 < 5 else "F" for d in range(42)] for t in range(6)],
              "5x2": A["modelo_5x2"]["grade"], "4x2": A["modelo_4x2"]["grade"], "12x36": A["modelo_12x36"]["grade"],
              "3x3": A["modelo_3x3"]["grade"], "4x2_12": A["modelo_4x2_12"]["grade"]}
    grades = {k: {"linhas": v, "rotulos": ROTULOS[k][0], "nota": ROTULOS[k][1], "leg": ROTULOS[k][2]} for k, v in grades.items()}
    op6 = []
    for x in A["op6"]["por_codigo"]:
        opa = ("Seg–sex 06:00–14:00 (1 h de intervalo): 7 h/dia, 35 h. Sem contratar e sem perder horas" if x["codigo"] == 300
               else f"Seg–sex no mesmo horário: {analise_hm(x['h_dia'] * 5)}. Sem contratar, mas sem sábado")
        op6.append({"codigo": x["codigo"], "horario": horario(x["descricao"]), "n": x["n"], "h_sem": x["h_sem"],
                    "cargos": ", ".join(f"{cargo(f.split(' (')[0])} {q}" for f, q in x["funcoes"][:3]), "opA": opa,
                    "contratar_seg_sab": x["contratar_seg_sab"]})
    adm_tot = A["adm6"]["n"] + A["adm5"]["n"]
    fam_adm = ", ".join(f"{f.lower()} {q}" for f, q in A["adm6"]["familias"][:5])
    cards = [
        {"tom": "warn", "titulo": f"Horário ADM: {adm_tot} pessoas vão para segunda a sexta, 8 h por dia", "itens": [
            f"{A['adm6']['n']} com sábado (códigos 807, 814, 815 e 801): sai o sábado. Fica 07:00–11:00 / 12:00–16:00, 40 h.",
            f"{A['adm5']['n']} com sábado compensado (910, 76 e 921): hoje fazem 8h48 por dia (44 h). Passam para 8 h por dia (40 h).",
            f"Sem contratação. A capacidade cai {f'{adm_tot * 4:,}'.replace(',', '.')} h por semana (4 h por pessoa), o equivalente a {adm_tot * 4 / 40:.0f} pessoas de 40 h. Funções operacionais nesse horário: {fam_adm}.",
        ]},
        {"tom": "ok", "titulo": f"12x36: {A['d1236']['n']} pessoas (códigos Portaria) continuam na escala", "itens": [
            ", ".join(f"{f} {q}" for f, q in A["d1236"]["familias"][:4]) + ".",
            "A 12x36 só continua com norma coletiva escrita nas regras novas. As cláusulas antigas incompatíveis perdem efeito em 15/12/2026.",
            "Com 1 h de intervalo gozado são 11 h por plantão e 38h30 por semana. Sem o intervalo seriam 42 h, acima das 40 h de dez/2027.",
        ]},
    ]
    atual = [
        ["5x1 turnos 24 h (turno de 7h50)", sum(1 for p in A["pessoas"] if p["grupo"] == "5x1 turnos 24 h" and round(p["h_dia"], 2) == 7.83),
         "45h42", "1,17", ["bad", "Não cumpre"], ["bad", "Não cumpre"]],
        ["5x1 turnos 24 h (turno de 7h20)", sum(1 for p in A["pessoas"] if p["grupo"] == "5x1 turnos 24 h" and round(p["h_dia"], 2) == 7.33),
         "42h47", "1,17", ["bad", "Não cumpre"], ["bad", "Não cumpre"]],
        ["6x1 operacional (dia)", A["op6"]["n"], "34h30 a 47h00", "1", ["bad", "Não cumpre"], ["bad", "Não cumpre"]],
        ["ADM com sábado (6x1)", A["adm6"]["n"], "44h00", "1", ["bad", "Não cumpre"], ["bad", "Não cumpre"]],
        ["ADM sábado compensado (8h48/dia)", A["adm5"]["n"], "44h00", "2", ["bad", "Não cumpre"], ["bad", "Não cumpre"]],
        ["12x36 (códigos Portaria)", A["d1236"]["n"], "38h30", "3,5", ["warn", "Com norma coletiva"], ["warn", "Com norma coletiva"]],
    ]
    c5 = T["contratar"]["5x2"]
    plano = [
        ["Até 31/10/2026", "Diretoria escolhe a escala dos turnos 24 h e aprova as contratações. Jurídico confere a EC 141 e as normas coletivas (rural e industrial)."],
        ["Novembro", f"Recrutamento e treinamento conforme a escala escolhida (5x2: {c5} pessoas para as turmas G). Prioridade: operadores de máquinas e motoristas."],
        ["Novembro", "Negociação com os sindicatos: 12x36, domingos, hora noturna, banco de horas e, se for o caminho, a prorrogação da 4x2 de 12 h."],
        ["Até 01/12/2026", "Cadastrar no ponto os códigos novos, montar as turmas e avisar cada pessoa da escala nova."],
        [lei["fase1"], "Começa a escala nova: 2 folgas por semana e no máximo 42 h."],
        [lei["fase2"], "Limite de 40 h. 5x2, 4x2 de 8 h e 12x36 com intervalo já cumprem."],
    ]
    atencao = [
        "Hora extra não substitui contratação: as 2 folgas não podem ser trabalhadas como rotina. Folga trabalhada é paga em dobro.",
        f"Entressafra: se ela começar antes de {lei['fase1']}, a troca acontece com menos frentes rodando e parte das contratações pode ficar para o início da safra 2027.",
        "Hora noturna na indústria: para o trabalhador urbano a hora noturna é reduzida. Conferir na norma coletiva como o turno da noite é contado.",
        "Mulheres em escala com domingo: domingo de folga a cada 15 dias (CLT art. 386).",
        f"Fora da conta: {G.get('Afastado (INSS)', 0)} afastados pelo INSS (entram na escala nova ao voltar), {G.get('Sem registro no revezamento', 0)} sem registro no revezamento, a maioria lideranças (as 2 folgas valem para eles também), e {G.get('Férias/atestado sem escala de origem', 0)} em férias ou atestado sem escala de origem.",
    ]
    premissas = [
        "5x1 = horários com 6 códigos de escala, um por dia de folga do ciclo de 6 dias.",
        "Mantém a cobertura de hoje: o mesmo número de pessoas trabalhando por dia em cada turno. Férias e faltas continuam na mesma proporção.",
        f"Frotas 24 h = {round(T['frotas24'])} postos que rodam nos 3 turnos; os {round(T['parciais'])} postos-turno parciais ficam em 8 h (5x2) nas sugestões de 12 h.",
        "Turno de 8 h com 7h50 de trabalho; 12x36 e 3x3 com 11 h; 4x2 de 12 h com 10 h (8 h + 2 h extras).",
        "Hora extra: horas por mês ÷ 200 (salário de 40 h) × adicional × reflexos (13º e férias). DSR sobre a hora extra habitual pode somar mais.",
        "Rendição: 1 rendicionista para cada 5 frotas por turno, também em 4x2 de 12 h.",
    ]
    return {
        "base": "20/10/2026", "lei": lei, "total": len(A["pessoas"]), "situacao": A["situacao"],
        "grupos": [{"nome": g, "n": G.get(g, 0), "tom": t, "acao": a} for g, t, a in GRUPOS],
        "celulas": [{k: c[k] for k in ("familia", "area", "triade", "n", "manha", "tarde", "noite")} for c in T["celulas"]],
        "frotas24": round(T["frotas24"]), "parciais": round(T["parciais"]),
        "cenarios": cen, "par": {"he_divisor": A["parametros"]["he_divisor"], "dias_mes": A["parametros"]["dias_mes"]},
        "grades": grades, "op6": {"n": A["op6"]["n"], "contratar_seg_sab": A["op6"]["contratar_seg_sab"], "linhas": op6},
        "adm6": {"n": A["adm6"]["n"]}, "adm5": {"n": A["adm5"]["n"]}, "cards_outros": cards,
        "checagens": CHECAGENS, "risco_det": RISCO_DET, "lei_tab": LEI_TAB, "atual_tab": atual,
        "plano": plano, "atencao": atencao, "premissas": premissas, "fontes": FONTES,
    }


def analise_hm(h):
    m = round(h * 60)
    return f"{m // 60}h{m % 60:02d}"


def gerar(A=None):
    A = A or analise.main()
    js = json.dumps(dados(A), ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    pagina = TEMPLATE.read_text(encoding="utf-8").replace("/*DADOS*/null", js)
    ARTIFACT.parent.mkdir(exist_ok=True)
    ARTIFACT.write_text(pagina, encoding="utf-8")
    corte = pagina.index("</style>") + len("</style>")
    LOCAL.write_text("<!doctype html>\n<html lang=\"pt-BR\"><head><meta charset=\"utf-8\">\n"
                     "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1, viewport-fit=cover\">\n"
                     + pagina[:corte] + "\n</head><body>\n" + pagina[corte:] + "\n</body></html>\n", encoding="utf-8")
    return LOCAL


if __name__ == "__main__":
    print(gerar())
