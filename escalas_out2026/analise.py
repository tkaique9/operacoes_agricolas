"""Lê o relatório de revezamento e o cadastro de ativos, classifica a escala de cada colaborador
e simula as escalas que cumprem a EC 141/2026 (fim da 6x1, 2 folgas por semana, 42 h e depois 40 h).

python analise.py  ->  analise.json
Planilhas em dados/: escalas.xlsx (Relatório de Revezamento) e ativoss.xlsx (Funcionários).
"""
import collections
import json
import math
import re
from pathlib import Path

import openpyxl

AQUI = Path(__file__).parent
DADOS = AQUI / "dados"

# ---------------------------------------------------------------- regras da EC 141/2026
LEI = {
    "promulgacao": "16/10/2026",
    "fase1": "15/12/2026",   # 60 dias: 2 folgas remuneradas por semana e 42 h semanais
    "fase2": "dez/2027",     # 12 meses depois: 40 h semanais
    "h_dia": 8.0,
    "h_sem_fase1": 42.0,
    "h_sem_fase2": 40.0,
    "folgas_sem": 2.0,
}

# ---------------------------------------------------------------- códigos de escala
# Cada horário de turno tem 6 códigos (ex.: 566 a 571 = 07:00 às 15:20), um por dia de folga do
# ciclo de 6 dias: é a 5x1. Os que têm "6x1" no nome, o 300 (código único, folga fixa) e os ADM
# com sábado trabalhado são 6x1.
AUSENCIA = {888: "Afastado (INSS)", 999: "Férias", 107: "Atestado", 500: "Atestado", 997: "Sem revezamento"}
SEIS_X_UM_OP = {919, 924, 552, 300}
ADM_6X1 = {807, 814, 815, 801}
ADM_5X2 = {910, 76, 921}
DOZE_36 = {68, 69, 70, 71}
SEM_INTERVALO_NA_DESCRICAO = {924}   # "6X1 - 06:00 AS 14:20": assumido 30 min como nos demais
H_SEM_ADM = {807: 44.0, 814: 44.0, 815: 44.0, 801: 44.0, 910: 44.0, 76: 44.0, 921: 44.0}


def horas_do_texto(txt):
    """Horas trabalhadas por dia e hora de entrada a partir da descrição do código."""
    t = txt.upper().replace(" ", "")
    hm = [int(a) * 60 + int(b) for a, b in re.findall(r"(\d{2}):(\d{2})", t)]
    if not hm:
        return None, None
    ini = txt.strip()[:5] if re.match(r"\d{2}:\d{2}", txt.strip()) else None
    if ini is None:
        m = re.search(r"(\d{2}:\d{2})", txt)
        ini = m.group(1)
    if "INT" in t and "INTER" not in t and len(hm) == 4:      # 06:00 12:00 - INT 09:00 09:15
        dur = (hm[1] - hm[0]) - (hm[3] - hm[2])
    elif len(hm) == 4:                                        # 07:00 - 12:00 e 13:00 - 15:20
        dur = (hm[1] - hm[0]) % 1440 + (hm[3] - hm[2]) % 1440
    elif len(hm) == 2:                                        # 07:00 AS 15:20 (INTER 30 MIN)
        iv = re.search(r"INTER\.?(\d+)", t)
        dur = (hm[1] - hm[0]) % 1440 - (int(iv.group(1)) if iv else 0)
    else:
        return None, ini
    return dur / 60, ini


def classificar_codigo(cod, desc):
    h, ini = horas_do_texto(desc)
    d = {"codigo": cod, "descricao": desc, "h_dia": h, "inicio": ini}
    if cod in SEM_INTERVALO_NA_DESCRICAO:
        h -= 0.5
        d["h_dia"] = h
    if cod in AUSENCIA:
        d.update(regime="Ausência", grupo=AUSENCIA[cod], h_sem=None)
    elif cod in DOZE_36:
        d.update(regime="12x36", grupo="12x36 (códigos Portaria)", h_sem=38.5, h_dia=11.0)
    elif cod in ADM_6X1:
        d.update(regime="6x1", grupo="ADM com sábado (6x1)", h_sem=H_SEM_ADM[cod], h_dia=8.0)
    elif cod in ADM_5X2:
        d.update(regime="5x2", grupo="ADM sábado compensado (5x2)", h_sem=H_SEM_ADM[cod], h_dia=8.8)
    elif cod in SEIS_X_UM_OP:
        d.update(regime="6x1", grupo="6x1 operacional (dia)", h_sem=round(h * 6, 2))
    elif 524 <= cod <= 585 or 720 <= cod <= 757:
        d.update(regime="5x1", grupo="5x1 turnos 24 h", h_sem=round(h * 35 / 6, 2))
    else:
        d.update(regime="?", grupo="Não classificado", h_sem=None)
    if d["grupo"] == "5x1 turnos 24 h":
        hh = int(ini[:2])
        d["triade"] = {5: "05-13-21", 13: "05-13-21", 21: "05-13-21", 6: "06-14-22", 14: "06-14-22",
                       22: "06-14-22", 7: "07-15-23", 15: "07-15-23", 23: "07-15-23"}[hh]
        d["turno"] = {5: "Manhã", 6: "Manhã", 7: "Manhã", 13: "Tarde", 14: "Tarde", 15: "Tarde"}.get(hh, "Noite")
    return d


# ---------------------------------------------------------------- famílias de função
FAMILIAS = [
    ("Lideranças", r"^(LIDER|LÍDER|ENCARREGADO|SUPERVISOR|COORDENADOR|GERENTE|FISCAL|MONITOR|ESPECIALISTA)"),
    ("Op. de máquinas agrícolas", r"^OP\. DE MAQUINAS AGRICOLAS"),
    ("Motoristas", r"^MOTORISTA|ENTREGADOR"),
    ("Auxiliares agrícolas e rurais", r"^(AUXILIAR AGRICOLA|TRABALHADOR RURAL)"),
    ("Operação industrial", r"^(OPERADOR (INDUSTRIAL|DE COZIMENTO|DE CENTRIFUGA|DE CALDEIRA|TRATAMENTO|DE EVAPORACAO|"
                            r"DE PAINEL|TURBO|DE ETA|DE E\.T\.A|DE CALEADEIRA|DE SONDA)|OP\. DE CENTRIFUGA|DESTILADOR|"
                            r"FERMENTADOR|AUXILIAR DE PRODUCAO)"),
    ("Manutenção", r"^(MECANICO|ELETRICISTA|SOLDADOR|CALDEIREIRO|LUBRIFICADOR|LAVADOR|BORRACHEIRO|INSTRUMENTISTA|"
                   r"TORNEIRO|AUXILIAR DE (MECANICO|ELETRICISTA|INSTRUMENTACAO)|PINTOR|PEDREIRO|SERVENTE|DESENHISTA)"),
    ("Logística e abastecimento", r"^(OPERADOR DE EMPILHADEIRA|OPERADOR GUINDASTE|FRENTISTA|FATURISTA)"),
    ("Laboratório e qualidade", r"^(ANALISTA (DE )?LABORAT|ANALISTA DE MICRO|ANALISTA DE QUALIDADE|QUIMICO|BIOLOGA)"),
    ("Segurança patrimonial", r"^(VIGILANTE|INSPETOR DE SEGURANCA)"),
    ("Administrativo e apoio", r"."),
]


def familia(funcao):
    for nome, rx in FAMILIAS:
        if re.search(rx, funcao):
            return nome
    return "Administrativo e apoio"


def area(cc):
    return "Industrial" if cc.startswith("2") else "Agrícola e automotiva"


# ---------------------------------------------------------------- cenários para os turnos 24 h
# Frota 24 h = posto que roda nos 3 turnos de hoje (menor nº de pessoas trabalhando por dia entre manhã,
# tarde e noite, por função, área e grupo de horários). Os outros postos (postos parciais, quase todos de
# manhã) continuam em turno de 8 h: nos cenários de 12 h eles vão para a 5x2.
# Pessoas por frota = turnos x dias do ciclo / dias trabalhados x (1 + rendicionistas por frota).
HE_DIVISOR = 200          # horas do salário mensal com 40 h semanais
DIAS_MES = 30.4375
HE_ADICIONAL = 0.50       # mínimo legal
HE_REFLEXOS = 0.20        # 13º e férias + 1/3 sobre a hora extra habitual (DSR à parte)
OITO = 7 + 50 / 60        # turno de 8h20 com 30 min de intervalo
CENARIOS = [
    {"id": "5x1", "nome": "5x1 (hoje)", "curto": "5x1 hoje", "turnos": 3, "trab": 5, "ciclo": 6, "h_trab": OITO,
     "turmas": 6, "parada": 1.5, "risco": "Proibida",
     "horario": "07:00–15:20 · 15:00–23:20 · 23:00–07:20 (30 min de intervalo)",
     "obs": "Não cumpre: 1,17 folga por semana e 45h41 de média"},
    {"id": "5x2", "nome": "5x2 rotativa · 3 turnos de 8 h", "curto": "5x2 rotativa", "turnos": 3, "trab": 5, "ciclo": 7,
     "h_trab": OITO, "turmas": 7, "parada": 1.5, "risco": "Baixo",
     "horario": "Mesmos horários de hoje: 07:00–15:20 · 15:00–23:20 · 23:00–07:20",
     "obs": "Mesmos 3 turnos de 8 h; 7 turmas por turno; folgas giram em 7 semanas"},
    {"id": "4x2", "nome": "4x2 · 3 turnos de 8 h", "curto": "4x2 de 8 h", "turnos": 3, "trab": 4, "ciclo": 6,
     "h_trab": OITO, "turmas": 3, "parada": 1.5, "risco": "Baixo",
     "horario": "Mesmos horários de hoje: 07:00–15:20 · 15:00–23:20 · 23:00–07:20",
     "obs": "Mesmos 3 turnos; 3 turmas por turno; mais folga e mais gente"},
    {"id": "12x36", "nome": "12x36 · 2 turnos de 12 h", "curto": "12x36", "turnos": 2, "trab": 1, "ciclo": 2,
     "h_trab": 11.0, "turmas": 2, "parada": 2.0, "risco": "Médio",
     "horario": "07:00–19:00 e 19:00–07:00, 1 h de intervalo (11 h de trabalho)",
     "obs": "Trabalha um dia sim, um dia não; CLT art. 59-A com norma coletiva"},
    {"id": "3x3", "nome": "3x3 · 2 turnos de 12 h", "curto": "3x3 de 12 h", "turnos": 2, "trab": 3, "ciclo": 6,
     "h_trab": 11.0, "turmas": 2, "parada": 2.0, "risco": "Médio",
     "horario": "07:00–19:00 e 19:00–07:00, 1 h de intervalo (11 h de trabalho)",
     "obs": "3 plantões seguidos e 3 dias de folga; compensação por norma coletiva"},
    {"id": "4x2_12", "nome": "4x2 · 12 h com 2 h de intervalo", "curto": "4x2 de 12 h", "turnos": 2, "trab": 4,
     "ciclo": 6, "h_trab": 10.0, "he_dia": 2.0, "turmas": 3, "parada": 4.0, "risco": "Alto",
     "horario": "07:00–19:00 (intervalo 12:00–14:00) e 19:00–07:00 (intervalo 00:00–02:00): 10 h de trabalho",
     "obs": "8 h normais + 2 h extras em todo plantão; máquina para no intervalo"},
    {"id": "4x2_12r", "nome": "4x2 · 12 h com 2 h de intervalo + rendição", "curto": "4x2 12 h + rendição",
     "turnos": 2, "trab": 4, "ciclo": 6, "h_trab": 10.0, "he_dia": 2.0, "rend": 0.2, "turmas": 3, "parada": 0.0,
     "risco": "Alto",
     "horario": "Igual à 4x2 de 12 h, com 1 rendicionista para cada 5 frotas cobrindo os intervalos",
     "obs": "Máquina roda 24 h; rendicionista também em 4x2 de 12 h"},
]
for c in CENARIOS:
    c.setdefault("he_dia", 0.0)
    c.setdefault("rend", 0.0)
    c["fator"] = c["ciclo"] / c["trab"]
    c["pessoas_posto"] = c["turnos"] * c["fator"] * (1 + c["rend"])
    c["h_sem"] = c["h_trab"] * 7 * c["trab"] / c["ciclo"]
    c["he_sem"] = c["he_dia"] * 7 * c["trab"] / c["ciclo"]
    c["he_mes"] = c["he_dia"] * DIAS_MES * c["trab"] / c["ciclo"]
    c["he_pct"] = c["he_mes"] / HE_DIVISOR * (1 + HE_ADICIONAL) * (1 + HE_REFLEXOS)
    c["folgas_sem"] = 7 * (c["ciclo"] - c["trab"]) / c["ciclo"]
BASE = CENARIOS[0]["pessoas_posto"]
CEN = {c["id"]: c for c in CENARIOS}


def necessidade(c, p24, parc):
    """Pessoas para cobrir p24 frotas 24 h e parc postos-turno parciais (por dia)."""
    if c["turnos"] == 3:
        return (3 * p24 + parc) * c["fator"]
    return p24 * c["pessoas_posto"] + parc * CEN["5x2"]["fator"]


def folha(c, p24, parc):
    """Folha em salários médios: quem faz hora extra custa 1 + he_pct."""
    if c["turnos"] == 3:
        return necessidade(c, p24, parc)
    return p24 * c["pessoas_posto"] * (1 + c["he_pct"]) + parc * CEN["5x2"]["fator"]


def ler_planilha(nome):
    return list(openpyxl.load_workbook(DADOS / nome, read_only=True, data_only=True).worksheets[0]
                .iter_rows(values_only=True))


def modelo_5x2():
    """Ciclo de 7 semanas: a dupla de folga da turma recua 1 dia por semana (sáb-dom, sex-sáb, ...)."""
    duplas = [(5, 6), (4, 5), (3, 4), (2, 3), (1, 2), (0, 1), (0, 6)]   # 0 = segunda
    grade = []
    for t in range(7):
        dias = []
        for w in range(7):
            fol = duplas[(w + t) % 7]
            dias += ["F" if d in fol else "T" for d in range(7)]
        grade.append(dias)
    return grade


def modelo_4x2():
    return [["T" if (d + 2 * t) % 6 < 4 else "F" for d in range(42)] for t in range(3)]


def modelo_12x36():
    """4 turmas: A e B de dia, C e D de noite, cada uma trabalha um dia sim, um dia não."""
    return [["T" if (d + t) % 2 == 0 else "F" for d in range(28)] for t in range(4)]


def modelo_3x3():
    """A e B de dia, C e D de noite: 3 plantões e 3 folgas."""
    return [["T" if (d + 3 * (t % 2)) % 6 < 3 else "F" for d in range(42)] for t in range(4)]


def modelo_4x2_12():
    """A, B e C de dia; D, E e F de noite: 4 plantões e 2 folgas, 2 de 3 turmas por dia."""
    return [["T" if (d + 2 * (t % 3)) % 6 < 4 else "F" for d in range(42)] for t in range(6)]


def rateio(total, partes):
    """Divide um total inteiro entre as partes pelo maior resto (a soma fecha com o total)."""
    base = {k: math.floor(v) for k, v in partes.items()}
    falta = total - sum(base.values())
    for k in sorted(partes, key=lambda k: -(partes[k] - base[k]))[:max(falta, 0)]:
        base[k] += 1
    return base


def checar_grade(grade):
    """Folgas por semana, maior sequência de trabalho e domingos de folga de cada turma."""
    res = []
    for linha in grade:
        sem = [linha[i:i + 7].count("F") for i in range(0, len(linha), 7)]
        dobro = linha + linha
        seq = mx = 0
        for v in dobro:
            seq = seq + 1 if v == "T" else 0
            mx = max(mx, seq)
        dom = sum(1 for i in range(6, len(linha), 7) if linha[i] == "F")
        res.append({"folgas_semana_min": min(sem), "folgas_semana_max": max(sem), "max_seguidos": mx,
                    "domingos_folga": dom, "semanas": len(linha) // 7})
    por_dia = [sum(1 for linha in grade if linha[d] == "T") for d in range(len(grade[0]))]
    return res, por_dia


def main():
    esc = [r for r in ler_planilha("escalas.xlsx") if isinstance(r[0], int) and isinstance(r[3], int)]
    ati = [r for r in ler_planilha("ativoss.xlsx") if isinstance(r[1], int) and r[4]]

    codigos = {}
    for r in esc:
        codigos.setdefault(r[3], classificar_codigo(r[3], r[4].strip()))
    hist = collections.defaultdict(list)
    for r in sorted(esc, key=lambda r: r[2], reverse=True):
        hist[r[0]].append(r)

    pessoas = []
    for r in ati:
        mat, nome, cc, sit, func, adm = r[1], r[2].strip(), r[0].strip(), r[4], r[6].strip(), r[7]
        h = hist.get(mat, [])
        atual = h[0] if h else None
        origem = next((x for x in h if x[3] not in AUSENCIA), None)
        p = {"matricula": mat, "nome": nome, "cc": cc, "area": area(cc), "situacao": sit, "funcao": func,
             "familia": familia(func), "admissao": adm.date().isoformat() if adm else None,
             "cod_atual": atual[3] if atual else None, "desde": atual[2].date().isoformat() if atual else None}
        if atual is None:
            p["grupo"] = "Sem registro no revezamento"
        elif sit == "AF" or atual[3] == 888:
            p["grupo"] = "Afastado (INSS)"
        elif atual[3] in AUSENCIA:
            if origem:
                p["cod_origem"] = origem[3]
                p["grupo"] = codigos[origem[3]]["grupo"]
                p["obs"] = f"{AUSENCIA[atual[3]]}; escala de origem pelo histórico ({origem[3]})"
            else:
                p["grupo"] = "Férias/atestado sem escala de origem"
        else:
            p["grupo"] = codigos[atual[3]]["grupo"]
        cod = p.get("cod_origem") or p["cod_atual"]
        if cod in codigos and codigos[cod]["regime"] not in ("Ausência",):
            c = codigos[cod]
            p.update(cod_escala=cod, escala=c["descricao"], regime=c["regime"], h_dia=c["h_dia"], h_sem=c["h_sem"],
                     inicio=c["inicio"], triade=c.get("triade"), turno=c.get("turno"))
        pessoas.append(p)

    # turma atual = posição do código dentro do horário (6 códigos por horário na 5x1)
    por_inicio = collections.defaultdict(set)
    for c in codigos.values():
        if c["grupo"] == "5x1 turnos 24 h":
            por_inicio[(c["inicio"], round(c["h_dia"], 2))].add(c["codigo"])
    letra = {}
    for _, cods in por_inicio.items():
        for i, cod in enumerate(sorted(cods)):
            letra[cod] = "ABCDEF"[i] if i < 6 else "F"
    for p in pessoas:
        if p["grupo"] == "5x1 turnos 24 h":
            p["turma_atual"] = letra[p["cod_escala"]]

    grupos = collections.Counter(p["grupo"] for p in pessoas)
    sit = collections.Counter(p["situacao"] for p in pessoas)
    turno24 = [p for p in pessoas if p["grupo"] == "5x1 turnos 24 h"]
    op6 = [p for p in pessoas if p["grupo"] == "6x1 operacional (dia)"]

    # células função x área x grupo de horários: frotas 24 h e postos parciais (pessoas trabalhando por dia)
    cel = collections.defaultdict(collections.Counter)
    for p in turno24:
        cel[(p["familia"], p["area"], p["triade"])][p["turno"]] += 1
    celulas = []
    for (fam, ar, tri), cnt in sorted(cel.items()):
        n = sum(cnt.values())
        p24 = min(cnt["Manhã"], cnt["Tarde"], cnt["Noite"]) * 5 / 6
        celulas.append({"familia": fam, "area": ar, "triade": tri, "n": n, "manha": cnt["Manhã"],
                        "tarde": cnt["Tarde"], "noite": cnt["Noite"], "p24": p24, "parc": n * 5 / 6 - 3 * p24})

    def agrega(campo):
        out = collections.defaultdict(lambda: {"n": 0, "p24": 0.0, "parc": 0.0})
        for c in celulas:
            k = out[c[campo]] if campo else out["Total"]
            k["n"] += c["n"]
            k["p24"] += c["p24"]
            k["parc"] += c["parc"]
        return out

    TOT = agrega(None)["Total"]
    tot_frac = {c["id"]: necessidade(c, TOT["p24"], TOT["parc"]) - TOT["n"] for c in CENARIOS[1:]}
    tot = {k: math.ceil(v - 1e-9) for k, v in tot_frac.items()}
    folha_tot = {c["id"]: folha(c, TOT["p24"], TOT["parc"]) for c in CENARIOS}

    def simular(campo):
        ag = agrega(campo)
        res = {k: {"chave": k, **v} for k, v in ag.items()}
        for c in CENARIOS[1:]:
            fr = {k: necessidade(c, v["p24"], v["parc"]) - v["n"] for k, v in ag.items()}
            for k, v in rateio(tot[c["id"]], fr).items():
                res[k][c["id"]] = v
                res[k][c["id"] + "_frac"] = fr[k]
        return sorted(res.values(), key=lambda x: -x["n"])

    sim_familia = simular("familia")
    sim_area = simular("area")
    # por cargo só os cenários de 8 h (os de 12 h dependem da divisão em frotas 24 h, feita por família)
    cont_f = collections.Counter(p["funcao"] for p in turno24)
    sim_funcao = []
    for k, n in cont_f.items():
        linha = {"chave": k, "n": n}
        for cid in ("5x2", "4x2"):
            linha[cid + "_frac"] = n * (CEN[cid]["pessoas_posto"] / BASE - 1)
        sim_funcao.append(linha)
    for cid in ("5x2", "4x2"):
        for k, v in rateio(tot[cid], {x["chave"]: x[cid + "_frac"] for x in sim_funcao}).items():
            next(x for x in sim_funcao if x["chave"] == k)[cid] = v
    sim_funcao.sort(key=lambda x: -x["n"])

    # turnos: pessoas por horário e turma (folga) atual
    tab_turno = collections.defaultdict(collections.Counter)
    for p in turno24:
        tab_turno[(p["triade"], p["inicio"], round(p["h_dia"], 2))][p["turma_atual"]] += 1
    turnos = []
    for (tri, ini, hd), cnt in sorted(tab_turno.items(), key=lambda x: (x[0][0], x[0][1][:2] in ("21", "22", "23"),
                                                                         x[0][1])):
        n = sum(cnt.values())
        turnos.append({"triade": tri, "inicio": ini, "h_dia": hd, "n": n, "turmas": dict(sorted(cnt.items())),
                       "dia_min": n - max(cnt.values()), "dia_med": n * 5 / 6})
    g = rateio(tot["5x2"], {i: t["n"] / 6 for i, t in enumerate(turnos)})
    for i, t in enumerate(turnos):
        t["nova_turma_G"] = g[i]

    # 6x1 operacional: manter seg-sáb (folga escalonada) ou concentrar em seg-sex
    op6_cod = collections.Counter(p["cod_escala"] for p in op6)
    op6_tab = []
    seg_sab = rateio(math.ceil(len(op6) / 5 - 1e-9), {k: n / 5 for k, n in op6_cod.items()})
    for cod, n in op6_cod.most_common():
        c = codigos[cod]
        op6_tab.append({"codigo": cod, "descricao": c["descricao"], "n": n, "h_dia": c["h_dia"], "h_sem": c["h_sem"],
                        "contratar_seg_sab": seg_sab[cod],
                        "funcoes": collections.Counter(p["funcao"] for p in op6 if p["cod_escala"] == cod)
                        .most_common(4)})

    adm6 = [p for p in pessoas if p["grupo"] == "ADM com sábado (6x1)" and p["situacao"] != "AF"]
    adm5 = [p for p in pessoas if p["grupo"] == "ADM sábado compensado (5x2)" and p["situacao"] != "AF"]
    d1236 = [p for p in pessoas if p["grupo"] == "12x36 (códigos Portaria)"]

    # proposta por pessoa (cenário recomendado)
    for p in pessoas:
        g = p["grupo"]
        if g == "5x1 turnos 24 h":
            p["proposta"] = f"5x2 rotativa · {p['escala'][:5]} (mesmo horário) · Turma {p['turma_atual']}"
            p["h_sem_nova"] = round(p["h_dia"] * 5, 2)
        elif g == "6x1 operacional (dia)":
            p["proposta"] = f"5x2 · {p['escala'][:5] if p['escala'][:2].isdigit() else p['inicio']} · folga escalonada seg–sáb"
            p["h_sem_nova"] = round(p["h_dia"] * 5, 2)
        elif g in ("ADM com sábado (6x1)", "ADM sábado compensado (5x2)"):
            p["proposta"] = "5x2 ADM · seg–sex 07:00–11:00 / 12:00–16:00"
            p["h_sem_nova"] = 40.0
        elif g == "12x36 (códigos Portaria)":
            p["proposta"] = "12x36 mantida · 1 h de intervalo gozado (11 h por plantão)"
            p["h_sem_nova"] = 38.5
        elif g == "Afastado (INSS)":
            p["proposta"] = "Entra na nova escala do setor ao retornar"
        elif g == "Sem registro no revezamento":
            p["proposta"] = "Verificar escala real (2 folgas por semana valem para todos)"
        else:
            p["proposta"] = "Definir escala de origem e aplicar a regra do grupo"

    g5, g4 = modelo_5x2(), modelo_4x2()
    chk5, dia5 = checar_grade(g5)
    chk4, dia4 = checar_grade(g4)
    g12 = modelo_12x36()
    chk12, dia12 = checar_grade(g12)
    g33, g412 = modelo_3x3(), modelo_4x2_12()
    chk33, dia33 = checar_grade(g33)
    chk412, dia412 = checar_grade(g412)

    out = {
        "lei": LEI, "cenarios": CENARIOS, "codigos": sorted(codigos.values(), key=lambda c: c["codigo"]),
        "pessoas": pessoas, "grupos": dict(grupos), "situacao": dict(sit),
        "turno24": {"n": len(turno24), "por_familia": sim_familia, "por_funcao": sim_funcao, "por_area": sim_area,
                    "contratar": tot, "contratar_frac": tot_frac, "turnos": turnos, "folha": folha_tot,
                    "celulas": celulas, "frotas24": TOT["p24"], "parciais": TOT["parc"],
                    "por_turno": dict(collections.Counter(p["turno"] for p in turno24)),
                    "por_hdia": dict(collections.Counter(str(round(p["h_dia"], 2)) for p in turno24))},
        "op6": {"n": len(op6), "por_codigo": op6_tab, "contratar_seg_sab": sum(x["contratar_seg_sab"] for x in op6_tab)},
        "adm6": {"n": len(adm6), "por_codigo": dict(collections.Counter(p["cod_escala"] for p in adm6)),
                 "familias": collections.Counter(p["familia"] for p in adm6).most_common()},
        "adm5": {"n": len(adm5), "por_codigo": dict(collections.Counter(p["cod_escala"] for p in adm5))},
        "d1236": {"n": len(d1236), "familias": collections.Counter(p["familia"] for p in d1236).most_common()},
        "modelo_5x2": {"grade": g5, "checagem": chk5, "trabalhando_por_dia": dia5},
        "modelo_4x2": {"grade": g4, "checagem": chk4, "trabalhando_por_dia": dia4},
        "modelo_12x36": {"grade": g12, "checagem": chk12, "trabalhando_por_dia": dia12},
        "modelo_3x3": {"grade": g33, "checagem": chk33, "trabalhando_por_dia": dia33},
        "modelo_4x2_12": {"grade": g412, "checagem": chk412, "trabalhando_por_dia": dia412},
        "parametros": {"he_divisor": HE_DIVISOR, "dias_mes": DIAS_MES, "he_adicional": HE_ADICIONAL,
                       "he_reflexos": HE_REFLEXOS},
    }
    (AQUI / "analise.json").write_text(json.dumps(out, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    return out


if __name__ == "__main__":
    o = main()
    print("situação", o["situacao"])
    for k, v in sorted(o["grupos"].items(), key=lambda x: -x[1]):
        print(f"{v:5d}  {k}")
    t = o["turno24"]
    print("turno 24h", t["n"], "frotas 24h", round(t["frotas24"], 1), "parciais", round(t["parciais"], 1))
    print("contratar", t["contratar"], {k: round(v, 1) for k, v in t["contratar_frac"].items()})
    print("folha", {k: round(v, 1) for k, v in t["folha"].items()})
    print("por turno", t["por_turno"], "h/dia", t["por_hdia"])
    for x in t["turnos"]:
        print(x)
    for x in t["por_familia"]:
        print(x)
    for x in t["por_area"]:
        print(x)
    print("op6", o["op6"]["n"], o["op6"]["contratar_seg_sab"])
    for x in o["op6"]["por_codigo"]:
        print(x)
    print("adm6", o["adm6"]), print("adm5", o["adm5"]), print("12x36", o["d1236"])
    print("5x2", o["modelo_5x2"]["checagem"], o["modelo_5x2"]["trabalhando_por_dia"][:14])
    print("4x2", o["modelo_4x2"]["checagem"], o["modelo_4x2"]["trabalhando_por_dia"][:14])
    for c in o["cenarios"]:
        print(c["id"], round(c["pessoas_posto"], 2), round(c["h_sem"], 2), round(c["he_sem"], 2),
              round(c["folgas_sem"], 2), round(c["he_pct"], 3))
    for k in ("modelo_3x3", "modelo_4x2_12"):
        print(k, o[k]["checagem"][0], o[k]["trabalhando_por_dia"][:12])
    nc = [c for c in o["codigos"] if c["regime"] == "?" or (c["h_dia"] is None and c["regime"] != "Ausência")]
    print("sem classificação/horas", nc)
