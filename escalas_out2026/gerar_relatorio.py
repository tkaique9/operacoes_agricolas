"""Gera o relatório da diretoria (PDF, A4 paisagem) e a planilha de simulação (xlsx).

python gerar_relatorio.py  ->  Relatorio_Escalas_EC141_OUT2026.pdf
                               Simulacao_Escalas_EC141_OUT2026.xlsx
Gráficos em matplotlib, páginas em HTML, PDF impresso pelo Edge em modo headless.
Depois rode recalcular_planilha.ps1 para o Excel gravar os resultados das fórmulas.
"""
import base64
import collections
import re
import subprocess
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import openpyxl  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

import analise  # noqa: E402
import gerar_html  # noqa: E402
from formatos import cargo, horario  # noqa: E402

AQUI = Path(__file__).parent
BUILD = AQUI / "build"
BUILD.mkdir(exist_ok=True)
PDF = AQUI / "Relatorio_Escalas_EC141_OUT2026.pdf"
XLSX = AQUI / "Simulacao_Escalas_EC141_OUT2026.xlsx"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

NAVY, RED, ORANGE, GREEN, GRAY = "#1f3d63", "#b3302d", "#d9902f", "#1e6b4a", "#7d8ea3"
LIGHT = "#c9d3e0"

plt.rcParams.update({"font.size": 8, "axes.spines.top": False, "axes.spines.right": False,
                     "axes.edgecolor": "#c9d0d8", "axes.grid": True, "grid.color": "#e8ecf0",
                     "axes.axisbelow": True, "legend.frameon": False})


# ---------------------------------------------------------------- formatação
def n(x, d=0):
    s = f"{x:,.{d}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def pct(a, b, d=1):
    return n(a / b * 100, d) + "%"


def hm(h):
    m = round(h * 60)
    return f"{m // 60}h{m % 60:02d}"


def img(fig, nome):
    caminho = BUILD / f"{nome}.png"
    fig.savefig(caminho, dpi=200, bbox_inches="tight")
    plt.close(fig)
    return "data:image/png;base64," + base64.b64encode(caminho.read_bytes()).decode()


# ---------------------------------------------------------------- dados
A = analise.main()
LEI = A["lei"]
CEN = {c["id"]: c for c in A["cenarios"]}
P = A["pessoas"]
G = A["grupos"]
T24 = A["turno24"]
N24 = T24["n"]
CONTR = T24["contratar"]
OP6 = A["op6"]
ADM6, ADM5, D1236 = A["adm6"], A["adm5"], A["d1236"]
TOTAL = len(P)
POSTOS = N24 / CEN["5x1"]["pessoas_posto"]
H733 = sum(1 for p in P if p["grupo"] == "5x1 turnos 24 h" and round(p["h_dia"], 2) == 7.33)
H783 = N24 - H733
SEM_ORIGEM = G.get("Férias/atestado sem escala de origem", 0)
INSS = G.get("Afastado (INSS)", 0)
SEMREG = G.get("Sem registro no revezamento", 0)
ADM_TOT = ADM6["n"] + ADM5["n"]
OP6_SEGSAB = OP6["contratar_seg_sab"]
OP6_MAQ = [x for x in OP6["por_codigo"] if x["codigo"] != 300]
OP6_300 = next((x for x in OP6["por_codigo"] if x["codigo"] == 300), None)
IDS = ["5x1", "5x2", "4x2", "12x36", "3x3", "4x2_12", "4x2_12r"]
GRADES = {"5x1": [["T" if (d + t) % 6 < 5 else "F" for d in range(42)] for t in range(6)],
          "5x2": A["modelo_5x2"]["grade"], "4x2": A["modelo_4x2"]["grade"], "12x36": A["modelo_12x36"]["grade"],
          "3x3": A["modelo_3x3"]["grade"], "4x2_12": A["modelo_4x2_12"]["grade"]}
GRADES["4x2_12r"] = GRADES["4x2_12"]
CHK = {}
for _i, _g in GRADES.items():
    _r, _ = analise.checar_grade(_g)
    CHK[_i] = {"dom": max(x["domingos_folga"] for x in _r), "sem": _r[0]["semanas"], "seq": max(x["max_seguidos"] for x in _r)}
FOLHA = T24["folha"]
COR_RISCO = {"Proibida": GRAY, "Baixo": GREEN, "Médio": ORANGE, "Alto": RED}

ORDEM_GRUPOS = [
    ("5x1 turnos 24 h", RED, "Muda: 5x1 → 5x2 rotativa"),
    ("6x1 operacional (dia)", RED, "Muda: 6x1 → 5x2"),
    ("ADM com sábado (6x1)", RED, "Muda: sai o sábado (5x2, 40 h)"),
    ("ADM sábado compensado (5x2)", ORANGE, "Ajusta horas: 44 h → 40 h"),
    ("12x36 (códigos Portaria)", GREEN, "Mantém (com norma coletiva)"),
    ("Afastado (INSS)", GRAY, "Fora da conta (entra ao voltar)"),
    ("Sem registro no revezamento", GRAY, "Verificar escala real"),
    ("Férias/atestado sem escala de origem", GRAY, "Verificar escala de origem"),
]


# ---------------------------------------------------------------- gráficos
def g_quadro():
    fig, ax = plt.subplots(figsize=(6.2, 3.1))
    nomes = [g for g, _, _ in ORDEM_GRUPOS][::-1]
    vals = [G.get(g, 0) for g in nomes]
    cores = [c for _, c, _ in ORDEM_GRUPOS][::-1]
    ax.barh(nomes, vals, color=cores, height=0.62)
    for i, v in enumerate(vals):
        ax.text(v + 12, i, n(v), va="center", fontsize=8, fontweight="bold", color="#1d2733")
    ax.set_xlim(0, max(vals) * 1.13)
    ax.grid(axis="y", visible=False)
    ax.tick_params(axis="y", length=0, labelsize=8)
    ax.set_xlabel("colaboradores")
    fig.tight_layout()
    return img(fig, "01_quadro")


def g_cenarios():
    rot = ["5x1\nhoje", "5x2\nrotativa", "4x2\n8 h", "12x36", "3x3\n12 h", "4x2\n12 h", "4x2 12 h\n+ rendição"]
    cores = [COR_RISCO[CEN[i]["risco"]] for i in IDS]
    fig, axs = plt.subplots(1, 2, figsize=(8.4, 2.75))
    pp = [CEN[i]["pessoas_posto"] for i in IDS]
    axs[0].bar(rot, pp, color=cores, width=0.62)
    for i, v in enumerate(pp):
        axs[0].text(i, v + 0.07, n(v, 1), ha="center", fontsize=8, fontweight="bold")
    axs[0].set_ylim(0, 5.3)
    axs[0].set_title("Pessoas por frota rodando 24 h × 7", fontsize=8.5, loc="left")
    tot = [N24 + CONTR.get(i, 0) for i in IDS]
    axs[1].bar(rot, tot, color=cores, width=0.62)
    for i, v in enumerate(tot):
        extra = "" if i == 0 else "\n" + ("+" if CONTR[IDS[i]] > 0 else "−") + n(abs(CONTR[IDS[i]]))
        axs[1].text(i, v + 25, n(v) + extra, ha="center", fontsize=7.4, fontweight="bold")
    axs[1].set_ylim(0, max(tot) * 1.25)
    axs[1].set_title(f"Pessoas nos turnos 24 h (hoje {n(N24)})", fontsize=8.5, loc="left")
    for ax in axs:
        ax.grid(axis="x", visible=False)
        ax.tick_params(axis="x", length=0, labelsize=6.6)
    fig.tight_layout(w_pad=2.5)
    return img(fig, "02_cenarios")


def g_turnos():
    tur = T24["turnos"]
    fig, ax = plt.subplots(figsize=(7.4, 2.5))
    rot = [f"{t['inicio']}\n{hm(t['h_dia'])}" for t in tur]
    hoje = [t["n"] for t in tur]
    novo = [t["nova_turma_G"] for t in tur]
    x = range(len(tur))
    ax.bar(x, hoje, color=NAVY, width=0.62, label="Pessoas hoje (turmas A–F)")
    ax.bar(x, novo, bottom=hoje, color=GREEN, width=0.62, label="Turma G a contratar")
    for i in x:
        ax.text(i, hoje[i] + novo[i] + 4, f"+{novo[i]}", ha="center", fontsize=7, color=GREEN, fontweight="bold")
    ax.set_xticks(list(x), rot, fontsize=6.6)
    ax.set_ylim(0, max(h + v for h, v in zip(hoje, novo)) * 1.15)
    ax.grid(axis="x", visible=False)
    ax.tick_params(axis="x", length=0)
    ax.legend(loc="upper right", fontsize=7.5)
    ax.set_ylabel("pessoas")
    fig.tight_layout()
    return img(fig, "03_turnos")


# ---------------------------------------------------------------- html
CSS = """
@page { size: A4 landscape; margin: 0; }
* { box-sizing: border-box; }
body { margin: 0; font-family: Arial, Helvetica, sans-serif; color: #1d2733; }
.page { width: 297mm; height: 210mm; position: relative; page-break-after: always; overflow: hidden; }
.page:last-child { page-break-after: auto; }
.hdr { background: #142a47; color: #fff; height: 21mm; padding: 5.5mm 11mm 0 11mm; position: relative; }
.hdr h1 { margin: 0; font-size: 17.5pt; letter-spacing: .2px; }
.hdr .sub { font-size: 10pt; color: #c9d3e0; margin-top: 2.3mm; }
.hdr .r1 { position: absolute; right: 11mm; top: 7mm; font-size: 9pt; color: #c9d3e0; }
.hdr .r2 { position: absolute; right: 11mm; top: 14.6mm; font-size: 9pt; color: #c9d3e0; }
.body { padding: 5mm 11mm 0 11mm; }
.ftr { position: absolute; bottom: 5mm; left: 11mm; right: 11mm; border-top: 1px solid #dde2e8;
       padding-top: 1.5mm; font-size: 7.4pt; color: #6b7785; }
.ftr span { float: right; }
.kpis { display: flex; gap: 4mm; margin-bottom: 4.5mm; }
.kpi { flex: 1; border: 1px solid #d5dbe2; border-top: 5px solid #142a47; border-radius: 6px;
       padding: 3.5mm 4.5mm 3mm 4.5mm; background: #fff; }
.kpi.red { border-color: #b3302d; border-top-color: #b3302d; background: #fbe9e8; }
.kpi.green { border-color: #1e6b4a; border-top-color: #1e6b4a; background: #e6f2ec; }
.kpi.orange { border-color: #d9902f; border-top-color: #d9902f; background: #fbf1e3; }
.kpi .t { font-size: 8pt; font-weight: bold; color: #6b7785; text-transform: uppercase; }
.kpi .v { font-size: 21pt; font-weight: bold; margin: 2mm 0 1.2mm 0; color: #142a47; }
.kpi.red .v { color: #b3302d; } .kpi.green .v { color: #1e6b4a; } .kpi.orange .v { color: #b8741c; }
.kpi .s { font-size: 8pt; color: #6b7785; display: flex; justify-content: space-between; align-items: center; gap: 2mm; }
.badge { white-space: nowrap; background: #142a47; color: #fff; font-weight: bold; border-radius: 4px; padding: 1.2mm 2.6mm; font-size: 8.4pt; }
.red .badge { background: #b3302d; } .green .badge { background: #1e6b4a; } .orange .badge { background: #d9902f; }
h2 { font-size: 12.5pt; margin: 0 0 2.5mm 0; border-left: 4.5mm solid #142a47; padding-left: 3.5mm; line-height: 6.5mm; }
h2.o { border-left-color: #d9902f; } h2.g { border-left-color: #1e6b4a; } h2.r { border-left-color: #b3302d; }
.row { display: flex; gap: 6mm; }
.col { flex: 1; min-width: 0; }
img.ch { width: 100%; display: block; }
table { border-collapse: collapse; width: 100%; font-size: 8.4pt; }
th { background: #142a47; color: #fff; padding: 1.6mm 2mm; text-align: right; font-weight: bold; border: 1px solid #fff; }
th:first-child, td:first-child { text-align: left; }
td { padding: 1.35mm 2mm; text-align: right; border: 1px solid #e1e6ec; }
tr:nth-child(even) td { background: #f5f7f9; }
tr.tot td { background: #e9edf1; font-weight: bold; }
tr.rec td { background: #e6f2ec !important; font-weight: bold; }
td.l, th.l { text-align: left; } td.c, th.c { text-align: center; }
.small td, .small th { font-size: 7.6pt; padding: 1.05mm 1.6mm; }
.box { background: #eef1f4; border: 1px solid #d5dbe2; border-radius: 6px; padding: 3mm 4.5mm; font-size: 8.8pt; line-height: 1.45; }
.box.o { background: #fbf1e3; border-color: #d9902f; }
.box.r { background: #fbe9e8; border-color: #b3302d; }
.box.g { background: #e6f2ec; border-color: #1e6b4a; }
.box h3 { margin: 0 0 1.5mm 0; font-size: 10.3pt; color: #142a47; }
.box.o h3 { color: #c47a1b; } .box.r h3 { color: #b3302d; } .box.g h3 { color: #1e6b4a; }
.box p { margin: 0 0 1mm 0; } .box ol, .box ul { margin: 0; padding-left: 5mm; } .box li { margin-bottom: .6mm; }
.note { font-size: 7.3pt; color: #6b7785; margin-top: 1.5mm; line-height: 1.35; }
.dot { display: inline-block; width: 2.6mm; height: 2.6mm; border-radius: 50%; margin-right: 1.5mm; vertical-align: middle; }
b.r { color: #b3302d; } b.g { color: #1e6b4a; } b.o { color: #b8741c; }
.rk { font-weight: bold; border-radius: 3mm; padding: .3mm 2mm; white-space: nowrap; }
.rk.Baixo { background: #e6f2ec; color: #1e6b4a; } .rk.Médio { background: #fbf1e3; color: #b8741c; }
.rk.Alto, .rk.Proibida { background: #fbe9e8; color: #b3302d; }
.ok { color: #1e6b4a; font-weight: bold; } .no { color: #b3302d; font-weight: bold; } .meh { color: #b8741c; font-weight: bold; }
.tl { display: flex; margin: 1mm 0 4mm 0; }
.tl .st { flex: 1; position: relative; padding: 0 3mm 0 0; }
.tl .st .bar { height: 3mm; background: #142a47; margin-bottom: 2mm; }
.tl .st.a .bar { background: #7d8ea3; } .tl .st.b .bar { background: #b3302d; } .tl .st.c .bar { background: #1e6b4a; }
.tl .st .d { font-size: 11pt; font-weight: bold; color: #142a47; }
.tl .st.b .d { color: #b3302d; } .tl .st.c .d { color: #1e6b4a; }
.tl .st .x { font-size: 8.3pt; color: #3a4654; line-height: 1.4; margin-top: .8mm; }
.grid { border-collapse: collapse; width: auto; font-size: 6.2pt; }
.grid td, .grid th { width: 5.2mm; height: 5.6mm; padding: 0; text-align: center; border: 1px solid #fff; }
.grid th { background: #142a47; color: #fff; font-weight: bold; }
.grid th.dom { background: #b3302d; }
.grid tr td.T { background: #dfe5ec; color: #6b7785; }
.grid tr td.F { background: #1e6b4a; color: #fff; font-weight: bold; }
.grid tr td.Fd { background: #b3302d; color: #fff; font-weight: bold; }
.grid tr td.lab { width: 17mm; text-align: left; padding-left: 1.5mm; font-weight: bold; background: #eef1f4; color: #142a47; font-size: 7.4pt; }
.grid tr td.cnt { background: #fff; color: #142a47; font-weight: bold; }
.grid td.wk { border-left: 2px solid #142a47; }
.mini .grid td, .mini .grid th { height: 4.4mm; }
.grid th.wk { border-left: 2px solid #142a47; }
"""

TITULO = "ESCALAS DE TRABALHO — FIM DA 6x1 (EC 141/2026)"
FONTE = ("Fonte: Relatório de Revezamento e Relação de Funcionários (sistema de RH, 20/10/2026) · "
         "Emenda Constitucional 141/2026, promulgada em 16/10/2026.")
PAGINAS = []


def pagina(sub, aba, corpo):
    PAGINAS.append((sub, aba, corpo))


def montar():
    tot = len(PAGINAS)
    out = []
    for i, (sub, aba, corpo) in enumerate(PAGINAS, 1):
        out.append(f"""<div class="page"><div class="hdr"><h1>{TITULO}</h1><div class="sub">{sub}</div>
<div class="r1">CRV Industrial Ltda · base de 20/10/2026</div><div class="r2">{aba}</div></div>
<div class="body">{corpo}</div><div class="ftr">{FONTE}<span>Página {i} de {tot}</span></div></div>""")
    return f"<!doctype html><html><head><meta charset='utf-8'><style>{CSS}</style></head><body>{''.join(out)}</body></html>"


def kpi(t, v, s, b, cls=""):
    return (f'<div class="kpi {cls}"><div class="t">{t}</div><div class="v">{v}</div>'
            f'<div class="s"><span>{s}</span><span class="badge">{b}</span></div></div>')


def tabela(cab, linhas, cls="", tot=None, classes=None, alin=None):
    """alin: uma letra por coluna (l = esquerda, c = centro, r = direita)."""
    a = [f' class="{x}"' if x in "lc" else "" for x in (alin or "l" + "r" * (len(cab) - 1))]
    h = "".join(f"<th{a[j]}>{c}</th>" for j, c in enumerate(cab))
    corpo = ""
    for i, ln in enumerate(linhas):
        c = f' class="{classes[i]}"' if classes and classes[i] else ""
        corpo += f"<tr{c}>" + "".join(f"<td{a[j]}>{x}</td>" for j, x in enumerate(ln)) + "</tr>"
    if tot:
        corpo += '<tr class="tot">' + "".join(f"<td{a[j]}>{c}</td>" for j, c in enumerate(tot)) + "</tr>"
    return f'<table class="{cls}"><tr>{h}</tr>{corpo}</table>'


OK, NO = '<span class="ok">Cumpre</span>', '<span class="no">Não cumpre</span>'


def linha_do_tempo():
    return f"""<div class="tl">
<div class="st a"><div class="bar"></div><div class="d">16/10/2026</div><div class="x">EC 141/2026 promulgada e publicada. Começa a contar o prazo.</div></div>
<div class="st b"><div class="bar"></div><div class="d">{LEI['fase1']}</div><div class="x"><b>2 folgas remuneradas por semana</b> e no máximo <b>42 h</b> semanais. Acaba a 6x1 e a 5x1.</div></div>
<div class="st c"><div class="bar"></div><div class="d">{LEI['fase2']}</div><div class="x">Limite cai para <b>40 h</b> semanais (8 h por dia). Salário não pode ser reduzido.</div></div>
</div>"""


# ---------------------------------------------------------------- página 1: resumo
def p1():
    c5 = CEN["5x2"]
    corpo = f"""<div class="kpis">
{kpi("Turnos 24 h em 5x1", n(N24), f"≈ {n(POSTOS)} postos 24 h equivalentes", "muda", "red")}
{kpi("6x1 operacional (dia)", n(OP6['n']), "folga fixa no domingo", "muda", "red")}
{kpi("Horário ADM", n(ADM_TOT), f"{ADM6['n']} com sábado + {ADM5['n']} compensado", "5x2 · 40 h", "orange")}
{kpi("Contratar (5x2 rotativa)", f"+{CONTR['5x2']}", f"+{pct(CONTR['5x2'], N24)} nos turnos", f"+0 a {OP6_SEGSAB} na 6x1", "green")}
</div>
<h2>O que a lei manda e quando</h2>{linha_do_tempo()}
<div class="row"><div class="col">
<div class="box g"><h3>Recomendação para os turnos 24 h: 5x2 rotativa</h3><ul>
<li>Ficam os <b>mesmos 3 turnos de 8 h</b> e os mesmos horários de hoje (ex.: 07:00, 15:00 e 23:00). Ninguém troca de turno.</li>
<li>Cada horário passa de 6 para <b>7 turmas</b>. As 6 turmas de folga de hoje continuam; a <b>turma G</b> é a contratação.</li>
<li>Cada pessoa trabalha 5 dias e folga 2 em <b>toda</b> semana: {hm(c5['h_sem'])} no turno de 7h50 e 36h40 no de 7h20. Nunca mais de 5 dias seguidos e 2 domingos de folga a cada 7 semanas.</li>
<li>Cumpre as duas fases (42 h em dez/2026 e 40 h em dez/2027). Não precisa refazer a escala em 2027.</li>
<li>Custo: <b>+{CONTR['5x2']} pessoas</b> (1 a cada 6 da 5x1). O equipamento continua com o mesmo número de pessoas por turno.</li>
</ul></div></div>
<div class="col"><div class="box o"><h3>O que a diretoria precisa decidir até o fim de outubro</h3><ol>
<li>Escala dos turnos 24 h: <b>5x2 rotativa</b> (+{CONTR['5x2']}), 12x36 (+{CONTR['12x36']}, com acordo coletivo), 4x2 de 8 h (+{CONTR['4x2']}) ou 4x2 de 12 h com 2 h de intervalo ({CONTR['4x2_12']:+d}, risco alto no MTE; página 6).</li>
<li>Aprovar as contratações e o início do recrutamento: faltam <b>8 semanas</b> para {LEI['fase1']}.</li>
<li>6x1 operacional ({OP6['n']} pessoas): manter o sábado com folga escalonada (+{OP6_SEGSAB}) ou trabalhar de segunda a sexta (sem contratar).</li>
<li>Horário ADM ({ADM_TOT} pessoas): segunda a sexta, 8 h por dia (40 h). Sem contratação.</li>
<li>Abrir a negociação com os sindicatos: 12x36, domingos, hora noturna e banco de horas.</li>
</ol></div></div></div>
<p class="note">Conta feita mantendo a cobertura de hoje: o mesmo número de pessoas trabalhando por dia em cada turno. Afastados (INSS), lideranças sem registro no revezamento e férias sem escala de origem ficam fora da conta (página 8).</p>"""
    pagina("Resumo para a diretoria — quem muda de escala, escala recomendada e contratações", "Resumo", corpo)


# ---------------------------------------------------------------- página 2: lei x escalas atuais
def p2():
    regras = tabela(["Regra", "Até 14/12/2026", f"A partir de {LEI['fase1']}", f"A partir de {LEI['fase2']}"], [
        ["Jornada semanal máxima", "44 h", "42 h", "40 h"],
        ["Jornada diária normal", "8 h", "8 h (até 8h24 em 5 dias, com norma coletiva)", "8 h"],
        ["Descanso semanal remunerado", "1 dia", "2 dias (um de preferência no domingo)", "2 dias"],
        ["Escala diferente da 5x2 (ex.: 12x36)", "acordo/convenção", "só com norma coletiva: média de 2 folgas por semana no mês e ao menos 1 folga por semana", "idem"],
        ["Salário e piso", "—", "sem redução nominal ou proporcional", "sem redução"],
        ["Cláusulas coletivas incompatíveis", "valem", "perdem efeito", "—"],
    ], "small", alin="lccc")
    c = CEN

    def h_lin(nome, qtd, hd, hs, fol, f1, f2, cls=""):
        return [nome, n(qtd), hd, hs, fol, f1, f2], cls

    adm5h = "8h48"
    linhas = [
        h_lin("5x1 turnos 24 h (turno de 7h50)", H783, "7h50", hm((7 + 50 / 60) * 35 / 6), n(c["5x1"]["folgas_sem"], 2), NO, NO),
        h_lin("5x1 turnos 24 h (turno de 7h20)", H733, "7h20", hm((7 + 20 / 60) * 35 / 6), n(c["5x1"]["folgas_sem"], 2), NO, NO),
        h_lin("6x1 operacional (dia)", OP6["n"], "5h45 a 7h50", "34h30 a 47h00", "1", NO, NO),
        h_lin("ADM com sábado (6x1)", ADM6["n"], "8h + sábado", "44h00", "1", NO, NO),
        h_lin("ADM sábado compensado (5x2)", ADM5["n"], adm5h, "44h00", "2", NO, NO),
        h_lin("12x36 (códigos Portaria)", D1236["n"], "11h (+1 h intervalo)", "38h30", "3,5", '<span class="meh">Com norma coletiva</span>', '<span class="meh">Com norma coletiva</span>'),
    ]
    tb = tabela(["Escala atual", "Pessoas", "Horas por dia", "Horas por semana (média)", "Folgas por semana",
                 f"{LEI['fase1']}", f"{LEI['fase2']}"], [x for x, _ in linhas], "small", alin="lrcrrcc")
    corpo = f"""<div class="row"><div class="col" style="flex:1.05"><h2>Regras da EC 141/2026</h2>{regras}
<p class="note">Fontes: Agência Brasil, Agência Câmara e análises jurídicas publicadas de 14 a 19/10/2026 sobre a EC 141/2026 (PEC 221/2019). Turnos ininterruptos de revezamento continuam com 6 h por dia, salvo negociação coletiva (art. 7º, XIV). Diplomados com salário de 2,5 vezes o teto do INSS ou mais podem ficar fora do controle de jornada, mas mantêm as 2 folgas. Conferir a redação final com o jurídico.</p>
</div><div class="col"><h2 class="r">Como ficam as escalas de hoje</h2>{tb}
<div class="box r" style="margin-top:3mm"><h3>Leitura</h3>
<p>Nenhuma escala de turno de hoje passa em {LEI['fase1']}: a 5x1 dá 1,17 folga por semana e, no turno de 7h50, {hm((7 + 50 / 60) * 35 / 6)} de média, acima das 42 h e das 40 h.</p>
<p>A 6x1 tem 1 folga. O horário ADM com sábado chega a 44 h e o sábado compensado tem 8h48 por dia, acima das 8 h. Só a 12x36 fica, desde que a norma coletiva seja renovada nas regras novas e o intervalo de 1 h seja de fato gozado (sem ele, o plantão de 12 h dá 42 h por semana, acima das 40 h de 2027).</p></div>
<p class="note">Horas de 5x1 = horas do turno × 5 dias a cada 6 (média de 5,83 dias por semana). 6x1 do código 924 sem intervalo na descrição: assumido 30 min. Horário ADM: 807 "7_11_12_16", 814/815 "07:20 - DOM" e 801 com sábado trabalhado; 910, 76 e 921 com sábado compensado (44 h de segunda a sexta).</p>
</div></div>"""
    pagina("O que a lei exige e onde cada escala de hoje deixa de cumprir", "A Lei", corpo)


# ---------------------------------------------------------------- página 3: quadro atual
def p3():
    leg = tabela(["Grupo", "Pessoas", "%", "O que acontece"],
                 [[f'<span class="dot" style="background:{c}"></span>{g}', n(G.get(g, 0)), pct(G.get(g, 0), TOTAL), o]
                  for g, c, o in ORDEM_GRUPOS], "small", tot=["Total do cadastro", n(TOTAL), "100%", ""], alin="lrrl")
    tur = collections.Counter((p["turno"], p["area"]) for p in P if p["grupo"] == "5x1 turnos 24 h")
    areas = ["Agrícola e automotiva", "Industrial"]
    lt = [[t] + [n(tur[(t, a)]) for a in areas] + [n(sum(tur[(t, a)] for a in areas))] for t in ("Manhã", "Tarde", "Noite")]
    tb_t = tabela(["Turno (5x1)"] + ["Agrícola e autom.", "Industrial"] + ["Total"], lt, "small",
                  tot=["Total"] + [n(sum(tur[(t, a)] for t in ("Manhã", "Tarde", "Noite"))) for a in areas] + [n(N24)])
    fam = T24["por_familia"]
    tb_f = tabela(["Função (5x1)", "Pessoas", "%"], [[f["chave"], n(f["n"]), pct(f["n"], N24)] for f in fam], "small",
                  tot=["Total", n(N24), "100%"])
    corpo = f"""<div class="row"><div class="col" style="flex:1.2"><h2>Os {n(TOTAL)} colaboradores do cadastro, pela escala de hoje</h2>
<img class="ch" src="{g_quadro()}">{leg}</div>
<div class="col"><h2 class="r">Quem está na 5x1 dos turnos 24 h</h2>{tb_t}<div style="height:3mm"></div>{tb_f}
<p class="note">Cadastro: {n(A['situacao'].get('AT', 0))} ativos, {n(A['situacao'].get('FE', 0))} em férias e {n(A['situacao'].get('AF', 0))} afastados. Escala atual = último registro do Relatório de Revezamento; para quem está de férias ou atestado vale o registro anterior, quando existe. Manhã inclui os turnos de 05:00, 06:00 e 07:00; tarde, 13:00 a 15:00; noite, 21:00 a 23:00. Agrícola e automotiva = centros de custo 1xx; industrial = 2xx.</p>
</div></div>"""
    pagina("Quadro atual — escala de cada colaborador e o que muda", "Quadro Atual", corpo)


# ---------------------------------------------------------------- página 4: cenários
def p4():
    cct = {"5x1": "—", "5x2": "Não", "4x2": "Não", "12x36": "Sim", "3x3": "Sim", "4x2_12": "Sim, na prática",
           "4x2_12r": "Sim, na prática"}
    linhas, cls = [], []
    for i in IDS:
        c, k = CEN[i], CHK[i]
        ext = CONTR.get(i, 0)
        horas = hm(c["h_sem"] - c["he_sem"]) + (f" + {hm(c['he_sem'])} extras" if c["he_sem"] else "")
        linhas.append([c["nome"], f'{c["turnos"]} × {"12 h" if c["turnos"] == 2 else "8 h"}', n(c["pessoas_posto"], 1),
                       horas, n(c["folgas_sem"], 2), f'{k["dom"]} a cada {k["sem"]} sem.', k["seq"], n(N24 + ext),
                       "—" if i == "5x1" else ("+" if ext > 0 else "−") + n(abs(ext)),
                       "—" if i == "5x1" else ("+" if FOLHA[i] > FOLHA["5x1"] else "") + pct(FOLHA[i] - FOLHA["5x1"], FOLHA["5x1"]),
                       f'{n(c["parada"], 1)} h' if c["parada"] else "0 (rendição)", cct[i],
                       f'<span class="rk {c["risco"]}">{c["risco"]}</span>'])
        cls.append("rec" if i == "5x2" else "")
    tb = tabela(["Sugestão", "Turnos por dia", "Pessoas por frota 24 h", "Horas por semana", "Folgas por semana",
                 "Domingos de folga", "Máx. dias seguidos", "Pessoas", "Contratar", "Folha x hoje",
                 "Máquina parada no intervalo (por dia)", "Precisa norma coletiva", "Risco no MTE"], linhas, "small",
                classes=cls, alin="lcrrrcrrrrrcc")
    corpo = f"""<h2>Sugestões de escala para manter cada frota rodando 24 h por dia, 7 dias por semana</h2>{tb}
<div class="row" style="margin-top:3mm"><div class="col" style="flex:1.3"><img class="ch" src="{g_cenarios()}"></div>
<div class="col"><div class="box g"><h3>Por que a 5x2 rotativa</h3><ul>
<li>É a escala que a lei toma como regra: não depende de acordo coletivo para existir.</li>
<li>Muda só a folga. Horários, turnos, ônibus e troca de turno continuam como hoje.</li>
<li>Sem hora extra fixa e com folga para atrasos na troca de turno: risco baixo no MTE.</li>
</ul></div>
<div class="box o" style="margin-top:2.2mm"><h3>As outras sugestões</h3><ul>
<li><b>12x36 e 3x3</b>: {CONTR['5x2'] - CONTR['12x36']} contratações a menos que a 5x2, mas 12 h operando ou dirigindo à noite e norma coletiva obrigatória.</li>
<li><b>4x2 de 12 h com 2 h de intervalo</b>: 3 pessoas por frota, {n(abs(CONTR['4x2_12']))} a menos que hoje. A folha sobe quase como na 5x2 por causa das 2 h extras de todo plantão, e a máquina para 4 h por dia. Risco alto (página 6).</li>
<li><b>4x2 de 8 h</b>: mais descanso, +{CONTR['4x2']} pessoas.</li>
</ul></div></div></div>
<p class="note">Pessoas por frota = turnos por dia × dias do ciclo ÷ dias trabalhados. Nas sugestões de 12 h, as {n(T24['frotas24'])} frotas que rodam nos 3 turnos vão para 12 h e os {n(T24['parciais'])} postos-turno parciais (quase todos de manhã) ficam em 5x2 de 8 h. Folha em salários médios, com hora extra a 50% e 20% de reflexos (13º e férias). Rendição: 1 rendicionista para cada 5 frotas por turno. 5x1, 6x1 e 6x2 ficam abaixo de 2 folgas por semana e não podem ser usadas.</p>"""
    pagina("Sugestões de escala para os turnos 24 h — pessoas, horas, folgas e risco no Ministério do Trabalho", "Sugestões", corpo)


# ---------------------------------------------------------------- página 5: modelo 5x2
def grade_html(grade, rotulos, semanas, gr_dom=True):
    dias = "STQQSSD"
    cab = '<tr><th style="width:16mm;background:#fff"></th>' + "".join(
        f'<th class="{"dom " if d % 7 == 6 else ""}{"wk" if d % 7 == 0 and d else ""}">{dias[d % 7]}</th>'
        for d in range(semanas * 7)) + "</tr>"
    sem = '<tr><th style="width:16mm;background:#fff"></th>' + "".join(
        f'<th colspan="7" class="{"wk" if w else ""}" style="background:#eef1f4;color:#142a47">Semana {w + 1}</th>'
        for w in range(semanas)) + "</tr>"
    corpo = ""
    for r, linha in zip(rotulos, grade):
        corpo += f'<tr><td class="lab">{r}</td>' + "".join(
            f'<td class="{"F" + ("d" if d % 7 == 6 else "") if v == "F" else "T"}{" wk" if d % 7 == 0 and d else ""}">'
            f'{"F" if v == "F" else ""}</td>' for d, v in enumerate(linha)) + "</tr>"
    cnt = [sum(1 for linha in grade if linha[d] == "T") for d in range(semanas * 7)]
    corpo += '<tr><td class="lab">Trabalhando</td>' + "".join(
        f'<td class="cnt{" wk" if d % 7 == 0 and d else ""}">{c}</td>' for d, c in enumerate(cnt)) + "</tr>"
    return f'<table class="grid">{sem}{cab}{corpo}</table>'


def p5():
    g = A["modelo_5x2"]["grade"]
    rot = [f"Turma {x}" for x in "ABCDEFG"]
    exemplo = next(t for t in T24["turnos"] if t["inicio"] == "07:00" and t["h_dia"] == 7.33)
    corpo = f"""<h2 class="g">Escala 5x2 rotativa — ciclo de 7 semanas, 7 turmas por horário (F = folga; vermelho = domingo de folga)</h2>
{grade_html(g, rot, 7)}
<div class="row" style="margin-top:3.5mm"><div class="col"><div class="box g"><h3>Como funciona</h3><ul>
<li>A dupla de folga de cada turma recua um dia por semana: sáb-dom, sex-sáb, qui-sex, qua-qui, ter-qua, seg-ter, seg e dom. Depois recomeça.</li>
<li>Toda semana (segunda a domingo) cada pessoa tem <b>2 folgas e 5 dias de trabalho</b>; nunca mais de 5 dias seguidos.</li>
<li>Em todos os dias trabalham <b>5 das 7 turmas</b>. Cada turma tem 1/5 dos postos do horário.</li>
<li>Domingo de folga 2 vezes a cada 7 semanas, para todos.</li>
</ul></div></div>
<div class="col"><div class="box"><h3>Passagem da 5x1 para a 5x2 sem trocar ninguém de horário</h3>
<p>Hoje cada horário tem 6 grupos de folga (6 códigos; ex.: 720 a 725 = 07:00–15:20 com 1 h de almoço). Cada grupo vira uma turma (A a F) com as mesmas pessoas. A <b>turma G</b> é montada com as contratações, do tamanho de um grupo de hoje.</p>
<p>Exemplo: 07:00–15:20 (1 h de almoço) tem {exemplo['n']} pessoas em 6 grupos ({', '.join(f'{k} {v}' for k, v in exemplo['turmas'].items())}); trabalham em média {n(exemplo['dia_med'])} por dia. Na 5x2 a turma G tem {exemplo['nova_turma_G']} pessoas e o horário fica com {exemplo['n'] + exemplo['nova_turma_G']}, mantendo os mesmos {n(exemplo['dia_med'])} por dia.</p>
<p>Os grupos de hoje têm tamanhos diferentes (no 07:00 com 1 h de almoço vão de 28 a 45 pessoas). A passagem é a hora de igualar as turmas.</p></div></div></div>"""
    pagina("Escala recomendada para os turnos 24 h: 5x2 rotativa com 7 turmas", "Escala 5x2", corpo)


# ---------------------------------------------------------------- página 6: 4x2 de 12 h
def p5b():
    c = CEN["4x2_12"]
    det = gerar_html.RISCO_DET["4x2_12"]
    rot = ["A · dia", "B · dia", "C · dia", "D · noite", "E · noite", "F · noite"]
    corpo = f"""<div class="kpis">
{kpi("Pessoas por frota 24 h", n(c['pessoas_posto'], 1), f"12x36: {n(CEN['12x36']['pessoas_posto'], 1)} · 5x2: {n(CEN['5x2']['pessoas_posto'], 1)}", "1 a menos", "green")}
{kpi("Pessoas nos turnos 24 h", n(N24 + CONTR['4x2_12']), f"hoje {n(N24)}", "−" + n(abs(CONTR['4x2_12'])), "green")}
{kpi("Folha mensal x hoje", "+" + pct(FOLHA['4x2_12'] - FOLHA['5x1'], FOLHA['5x1']), f"5x2: +{pct(FOLHA['5x2'] - FOLHA['5x1'], FOLHA['5x1'])}", f"+{pct(c['he_pct'], 1, 0)} por pessoa", "orange")}
{kpi("Risco no MTE", "Alto", "10 h é o teto: 8 h + 2 h extras", "jurídico", "red")}
</div>
<h2 class="o">Escala: 07:00–19:00 (intervalo 12:00–14:00) e 19:00–07:00 (intervalo 00:00–02:00) · 3 turmas de dia e 3 de noite</h2>
<div class="mini">{grade_html(A['modelo_4x2_12']['grade'], rot, 6)}</div>
<div class="row" style="margin-top:2.5mm"><div class="col" style="flex:1.3"><div class="box r" style="font-size:7.6pt;line-height:1.33;padding:2.5mm 4mm"><h3 style="font-size:9.5pt">É prejudicial com o Ministério do Trabalho? {det['veredito']}</h3><ul>
{''.join(f'<li>{x}</li>' for x in det['olha'])}
</ul></div></div>
<div class="col"><div class="box" style="font-size:7.6pt;line-height:1.33;padding:2.5mm 4mm"><h3>Se a diretoria quiser seguir com ela</h3><ul>
{''.join(f'<li>{x}</li>' for x in det['fazer'])}
</ul></div></div></div>"""
    pagina("Sugestão: 4x2 de 12 h com 2 h de intervalo — quadro menor e o risco com o Ministério do Trabalho", "4x2 de 12 h", corpo)


# ---------------------------------------------------------------- página 7: por turno e função
def p6():
    tur = T24["turnos"]
    lt = [[t["triade"].replace("-", " · "), f"{t['inicio']}", hm(t["h_dia"]), n(t["n"]), n(t["dia_med"], 1),
           f"+{t['nova_turma_G']}", n(t["n"] + t["nova_turma_G"])] for t in tur]
    tb_t = tabela(["Grupo de horários", "Entrada", "Horas/dia", "Pessoas hoje", "Trabalhando por dia",
                   "Turma G (contratar)", "Pessoas na 5x2"], lt, "small", alin="lccrrrr",
                  tot=["Total", "", "", n(N24), n(N24 * 5 / 6, 1), f"+{CONTR['5x2']}", n(N24 + CONTR['5x2'])])
    fam = T24["por_familia"]
    sg = lambda v: ("+" if v > 0 else "−" if v < 0 else "") + n(abs(v))  # noqa: E731
    lf = [[f["chave"], n(f["n"]), sg(f["5x2"]), sg(f["12x36"]), sg(f["4x2"]), sg(f["4x2_12"])] for f in fam]
    tb_f = tabela(["Função (5x1 hoje)", "Pessoas", "5x2 rotativa", "12x36", "4x2 de 8 h", "4x2 de 12 h"], lf, "small",
                  tot=["Total", n(N24), sg(CONTR['5x2']), sg(CONTR['12x36']), sg(CONTR['4x2']), sg(CONTR['4x2_12'])])
    top = sorted(T24["por_funcao"], key=lambda x: -x["n"])[:8]
    lfun = [[cargo(f["chave"]), n(f["n"]), f"+{f['5x2']}"] for f in top]
    tb_fun = tabela(["Cargo (maiores)", "Pessoas", "5x2"], lfun, "small")
    corpo = f"""<div class="row"><div class="col" style="flex:1.15"><h2>Turma G por horário — 5x2 rotativa</h2>
<img class="ch" src="{g_turnos()}">{tb_t}</div>
<div class="col"><h2>Contratações por função em cada sugestão</h2>{tb_f}<div style="height:3mm"></div>{tb_fun}
<p class="note">Turma G = pessoas do horário ÷ 6. Os totais fecham com o arredondamento do total geral (maior resto). Linhas de 06:00, 14:00 e 22:00 aparecem duas vezes porque há turnos de 7h50 (30 min de intervalo) e de 7h20 (1 h). A lista pessoa a pessoa, com a turma de cada um, está na planilha (aba Colaboradores).</p></div></div>"""
    pagina("Onde entram as contratações — por horário de turno e por função", "Por Turno e Função", corpo)


# ---------------------------------------------------------------- página 7: 6x1, ADM, 12x36
def p7():
    lo = []
    for x in OP6["por_codigo"]:
        fun = ", ".join(f"{cargo(f.split(' (')[0])} {q}" for f, q in x["funcoes"][:3])
        if x["codigo"] == 300:
            a = "Seg–sex 06:00–14:00 (1 h de intervalo): 7 h/dia, 35 h. Sem contratar e sem perder horas"
        else:
            a = f"Seg–sex no mesmo horário: {hm(x['h_dia'] * 5)}. Sem contratar, mas sem sábado"
        lo.append([f"{x['codigo']}", horario(x["descricao"]), n(x["n"]), fun, hm(x["h_sem"]), a, f"+{x['contratar_seg_sab']}"])
    tb_o = tabela(["Código", "Horário", "Pessoas", "Principais cargos", "h/sem hoje", "Opção A — segunda a sexta",
                   "Opção B — seg–sáb com folga escalonada"], lo, "small",
                  tot=["Total", "", n(OP6["n"]), "", "", "0", f"+{OP6_SEGSAB}"], alin="llrlrlr")
    perda = ADM_TOT * 4
    fam_adm = ", ".join(f"{f.lower()} {q}" for f, q in ADM6["familias"][:5])
    corpo = f"""<h2 class="r">6x1 operacional de dia — {OP6['n']} pessoas com folga fixa no domingo</h2>{tb_o}
<p class="note">Opção B: a operação continua de segunda a sábado e cada pessoa folga 2 dias, em dias diferentes (5 de cada 6 trabalham por dia): pessoas ÷ 5 a contratar. Para o 300 a opção A estica o dia (de 5h45 para 7 h) e mantém as horas da semana; mudar o horário diário pede concordância ou norma coletiva.</p>
<div class="row" style="margin-top:3mm"><div class="col"><div class="box o"><h3>Horário ADM — {n(ADM_TOT)} pessoas vão para segunda a sexta, 8 h por dia</h3><ul>
<li><b>{ADM6['n']} com sábado</b> (códigos 807, 814, 815 e 801): sai o sábado. Fica 07:00–11:00 / 12:00–16:00, 40 h.</li>
<li><b>{ADM5['n']} com sábado compensado</b> (910, 76 e 921): hoje fazem 8h48 por dia (44 h). Passam para 8 h por dia (40 h). Dá para fazer 8h24 até dez/2027 com norma coletiva, mas mudar uma vez só é mais simples.</li>
<li>Sem contratação. A capacidade cai {n(perda)} h por semana (4 h por pessoa), o equivalente a {n(perda / 40)} pessoas de 40 h. Vale olhar as funções operacionais que estão nesse horário ({fam_adm}).</li>
</ul></div></div>
<div class="col"><div class="box g"><h3>12x36 — {D1236['n']} pessoas (códigos Portaria) continuam na escala</h3><ul>
<li>{', '.join(f'{f.lower()} {q}' for f, q in D1236['familias'][:4])}.</li>
<li>A 12x36 só continua com norma coletiva escrita nas regras novas (média de 2 folgas por semana no mês). As cláusulas antigas incompatíveis perdem efeito em {LEI['fase1']}.</li>
<li>Com 1 h de intervalo gozado são 11 h por plantão e 38h30 por semana. Sem o intervalo seriam 42 h, acima das 40 h a partir de {LEI['fase2']}.</li>
</ul></div></div></div>"""
    pagina("6x1 operacional, horário ADM e 12x36 — o que muda em cada um", "6x1, ADM e 12x36", corpo)


# ---------------------------------------------------------------- página 8: implantação
def p8():
    passos = tabela(["Quando", "O quê"], [
        ["Até 31/10/2026", "Diretoria decide a escala e aprova as contratações. Jurídico confere a redação da EC 141 e as normas coletivas (rural e industrial)."],
        ["Novembro", f"Recrutamento, exames e treinamento de {CONTR['5x2']} pessoas para as turmas G (+{OP6_SEGSAB} se a 6x1 mantiver o sábado). Prioridade para operadores de máquinas e motoristas: {T24['por_familia'][0]['5x2'] + T24['por_familia'][1]['5x2']} das {CONTR['5x2']} vagas."],
        ["Novembro", "Negociação com os sindicatos: 12x36, domingos, hora noturna, banco de horas e horário de 8h24 até dez/2027, se for usado."],
        ["Até 01/12/2026", "Cadastrar no ponto os códigos novos (5x2 turmas A–G por horário, ADM 40 h), montar as turmas e avisar cada pessoa da escala nova."],
        [LEI["fase1"], "Começa a 5x2: 2 folgas por semana e no máximo 42 h. A 5x2 de 7h50 já fica em 39h10."],
        [LEI["fase2"], "Limite de 40 h. A escala recomendada já cumpre, nada a refazer."],
    ], "small", alin="ll")
    corpo = f"""<div class="row"><div class="col" style="flex:1.1"><h2>Plano de implantação</h2>{passos}
<div class="box" style="margin-top:3mm"><h3>Custo</h3><p>O salário de quem fica não muda. O custo novo vem das contratações: a cada R$ 1.000 de custo médio mensal por pessoa (salário, encargos e benefícios), as {CONTR['5x2']} vagas da 5x2 somam R$ {n(CONTR['5x2'])} mil por mês. A planilha tem um campo para o custo médio de cada função e calcula o total de cada cenário.</p></div>
</div><div class="col"><h2 class="o">Pontos de atenção</h2><div class="box o"><ul>
<li><b>Hora extra não resolve:</b> as 2 folgas não podem ser trabalhadas como rotina. Folga trabalhada é paga em dobro e a lei passa a ser descumprida.</li>
<li><b>Entressafra:</b> se ela começar antes de {LEI['fase1']}, a troca acontece com menos frentes rodando e parte das contratações pode ser feita até o início da safra 2027.</li>
<li><b>Hora noturna na indústria:</b> para o trabalhador urbano a hora noturna é reduzida. Conferir na norma coletiva como o turno de 23:00 é contado, para não passar de 8 h por dia.</li>
<li><b>Mulheres em escala com domingo:</b> a CLT (art. 386) prevê domingo de folga a cada 15 dias. Montar as turmas respeitando isso.</li>
<li><b>Fora da conta:</b> {INSS} afastados pelo INSS (entram na escala nova ao voltar); {SEMREG} sem registro no revezamento, a maioria lideranças (as 2 folgas valem para eles também); {SEM_ORIGEM} em férias ou atestado sem escala de origem no relatório.</li>
</ul></div>
<div class="box" style="margin-top:3mm"><h3>Premissas da simulação</h3><ul>
<li>5x1 = horários com 6 códigos de escala, um por dia de folga do ciclo de 6 dias.</li>
<li>A cobertura de hoje é mantida: mesmo número de pessoas trabalhando por dia em cada turno. Férias, faltas e afastamentos continuam na mesma proporção.</li>
<li>Escalas por código e horas da descrição de cada código (aba Códigos da planilha, onde dá para corrigir).</li>
</ul></div></div></div>"""
    pagina("Plano até 15/12/2026, custo e pontos de atenção", "Implantação", corpo)


# ---------------------------------------------------------------- planilha
AZUL, AMARELO, CINZA = "142A47", "FFF4CC", "EEF1F4"
FINO = Side(style="thin", color="D5DBE2")
BORDA = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)


def cab(ws, linha, valores):
    for j, v in enumerate(valores, 1):
        c = ws.cell(linha, j, v)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=AZUL)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = BORDA


def titulo(ws, txt, sub=None):
    ws["A1"] = txt
    ws["A1"].font = Font(bold=True, size=14, color=AZUL)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(italic=True, color="6B7785")


def larguras(ws, d):
    for col, w in d.items():
        ws.column_dimensions[col].width = w


def entrada(c):
    c.fill = PatternFill("solid", fgColor=AMARELO)
    c.border = BORDA


def base_xlsx():
    wb = openpyxl.Workbook()

    # ---- Premissas (cenários editáveis)
    wp = wb.active
    wp.title = "Premissas"
    titulo(wp, "Premissas das sugestões (células amarelas podem ser alteradas)",
           "Pessoas por frota 24 h = turnos por dia × dias do ciclo ÷ dias trabalhados × (1 + rendicionistas por frota).")
    cab(wp, 4, ["Sugestão", "Turnos por dia", "Dias trabalhados no ciclo", "Dias do ciclo", "Horas trabalhadas por turno",
                "Horas extras por plantão", "Rendicionistas por frota e turno", "Pessoas por frota 24 h", "Horas por semana",
                "Extras por semana", "Folgas por semana", "Hora extra sobre o salário", "Risco no MTE", "Observação"])
    for i, cid in enumerate(IDS):
        c = CEN[cid]
        r = 5 + i
        wp.cell(r, 1, c["nome"])
        for j, v in enumerate([c["turnos"], c["trab"], c["ciclo"], round(c["h_trab"], 4), c["he_dia"], c["rend"]], 2):
            entrada(wp.cell(r, j, v))
        wp.cell(r, 8, f"=B{r}*D{r}/C{r}*(1+G{r})").number_format = "0.00"
        wp.cell(r, 9, f"=E{r}*7*C{r}/D{r}").number_format = "0.00"
        wp.cell(r, 10, f"=F{r}*7*C{r}/D{r}").number_format = "0.00"
        wp.cell(r, 11, f"=7*(D{r}-C{r})/D{r}").number_format = "0.00"
        wp.cell(r, 12, f"=F{r}*$C$16*C{r}/D{r}/$C$17*(1+$C$18)*(1+$C$19)").number_format = "0.0%"
        wp.cell(r, 13, c["risco"])
        wp.cell(r, 14, c["obs"])
    wp["A14"] = "Parâmetros"
    wp["A14"].font = Font(bold=True, color=AZUL)
    for r, (t, v) in enumerate([("Linha da 5x2 (postos parciais vão para ela nas sugestões de 12 h)", 6),
                                ("Dias por mês", analise.DIAS_MES), ("Horas do salário mensal (40 h por semana)", analise.HE_DIVISOR),
                                ("Adicional de hora extra", analise.HE_ADICIONAL), ("Reflexos da hora extra (13º e férias)", analise.HE_REFLEXOS),
                                (f"Horas semanais a partir de {LEI['fase1']}", 42), (f"Horas semanais a partir de {LEI['fase2']}", 40),
                                ("Folgas remuneradas por semana (mínimo)", 2)], 15):
        wp.cell(r, 1, t)
        entrada(wp.cell(r, 3, v))
    wp["C18"].number_format = wp["C19"].number_format = "0%"
    wp["A24"] = ("Nas sugestões de 12 h só as frotas que rodam nos 3 turnos vão para 12 h; os postos parciais ficam em 5x2 de 8 h. "
                 "12x36 e 3x3 com 11 h de trabalho; 4x2 de 12 h com 10 h (8 h + 2 h extras).")
    larguras(wp, {"A": 44, "B": 9, "C": 11, "D": 9, "E": 11, "F": 10, "G": 13, "H": 11, "I": 10, "J": 10, "K": 10,
                  "L": 12, "M": 10, "N": 60})

    # ---- Simulação por função (fórmulas)
    ws = wb.create_sheet("Simulação")
    titulo(ws, "Simulação — turnos 24 h hoje em 5x1",
           "Frotas 24 h e postos parciais em pessoas trabalhando por dia. Preencha o custo médio mensal por pessoa (amarelo) para ver o custo.")
    SUGS = IDS[1:]
    cab_s = ["Função", "Pessoas hoje (5x1)", "Frotas 24 h", "Postos parciais", "Custo médio mensal por pessoa (R$)"]
    for cid in SUGS:
        cab_s += [f"{CEN[cid]['curto']} — necessário", f"{CEN[cid]['curto']} — contratar", f"{CEN[cid]['curto']} — folha (salários)"]
    cab(ws, 4, cab_s)
    fam = T24["por_familia"]
    r0 = 5
    L = get_column_letter
    for i, f in enumerate(fam):
        r = r0 + i
        ws.cell(r, 1, f["chave"])
        ws.cell(r, 2, f["n"])
        ws.cell(r, 3, round(f["p24"], 4)).number_format = "0.0"
        ws.cell(r, 4, round(f["parc"], 4)).number_format = "0.0"
        entrada(ws.cell(r, 5))
        ws.cell(r, 5).number_format = "#,##0.00"
        for k, cid in enumerate(SUGS):
            pr = 5 + IDS.index(cid)
            col = 6 + 3 * k
            oito = f"$D{r}*Premissas!$D$6/Premissas!$C$6"
            nec = (f"=IF(Premissas!$B${pr}=3,(3*$C{r}+$D{r})*Premissas!$D${pr}/Premissas!$C${pr},"
                   f"$C{r}*Premissas!$H${pr}+{oito})")
            ws.cell(r, col, nec).number_format = "0.0"
            ws.cell(r, col + 1, f"={L(col)}{r}-$B{r}").number_format = "0.0"
            ws.cell(r, col + 2, f"=IF(Premissas!$B${pr}=3,{L(col)}{r},$C{r}*Premissas!$H${pr}*(1+Premissas!$L${pr})+{oito})").number_format = "0.0"
    rt = r0 + len(fam)
    ws.cell(rt, 1, "Total (contratar arredondado para cima)").font = Font(bold=True)
    for j, colsum in ((2, "B"), (3, "C"), (4, "D")):
        ws.cell(rt, j, f"=SUM({colsum}{r0}:{colsum}{rt - 1})").font = Font(bold=True)
    for k in range(len(SUGS)):
        col = 6 + 3 * k
        ws.cell(rt, col, f"=SUM({L(col)}{r0}:{L(col)}{rt - 1})").number_format = "0.0"
        ws.cell(rt, col + 1, f"=-INT(-ROUND(SUM({L(col + 1)}{r0}:{L(col + 1)}{rt - 1}),6))")
        ws.cell(rt, col + 2, f"=SUM({L(col + 2)}{r0}:{L(col + 2)}{rt - 1})").number_format = "0.0"
        for j in (col, col + 1, col + 2):
            ws.cell(rt, j).font = Font(bold=True)
    ws.cell(rt + 1, 1, "Custo mensal da folha (R$)").font = Font(bold=True)
    for k in range(len(SUGS)):
        col = 6 + 3 * k + 2
        ws.cell(rt + 1, col, f'=IF(COUNT($E${r0}:$E${rt - 1})=0,"preencher custo",SUMPRODUCT({L(col)}{r0}:{L(col)}{rt - 1},$E${r0}:$E${rt - 1}))').number_format = "#,##0"
    ws.cell(rt + 1, 5, f'=IF(COUNT($E${r0}:$E${rt - 1})=0,"",SUMPRODUCT(B{r0}:B{rt - 1},E{r0}:E{rt - 1}))').number_format = "#,##0"
    ws.cell(rt + 2, 1, "Custo de hoje fica na coluna E da linha acima (pessoas hoje × custo médio).")
    for r in range(r0, rt + 2):
        for j in range(1, len(cab_s) + 1):
            ws.cell(r, j).border = BORDA
    ws.cell(rt + 4, 1, "6x1 operacional mantendo seg–sáb (folga escalonada): contratar = pessoas ÷ 5")
    ws.cell(rt + 5, 1, "Pessoas na 6x1 operacional")
    ws.cell(rt + 5, 2, OP6["n"])
    ws.cell(rt + 6, 1, "Contratar (seg–sáb)")
    ws.cell(rt + 6, 2, f"=-INT(-B{rt + 5}/5)")
    larguras(ws, {"A": 32, "B": 11, "C": 10, "D": 10, "E": 15})
    for j in range(6, len(cab_s) + 1):
        ws.column_dimensions[L(j)].width = 12
    ws.row_dimensions[4].height = 45
    ws.freeze_panes = "B5"

    # ---- Resumo
    wr = wb.create_sheet("Resumo", 0)
    titulo(wr, "Escalas de trabalho — fim da 6x1 (EC 141/2026)", "CRV Industrial Ltda · base de 20/10/2026")
    cab(wr, 4, ["Grupo (escala de hoje)", "Pessoas", "O que acontece"])
    for i, (g, _, o) in enumerate(ORDEM_GRUPOS):
        wr.cell(5 + i, 1, g)
        wr.cell(5 + i, 2, G.get(g, 0))
        wr.cell(5 + i, 3, o)
    rr = 5 + len(ORDEM_GRUPOS)
    wr.cell(rr, 1, "Total do cadastro").font = Font(bold=True)
    wr.cell(rr, 2, f"=SUM(B5:B{rr - 1})").font = Font(bold=True)
    cab(wr, rr + 2, ["Sugestão para os turnos 24 h", "Contratar", "Pessoas nos turnos 24 h", "Pessoas por frota 24 h",
                     "Horas por semana", "Folgas por semana", "Folha (salários médios)", "Risco no MTE"])
    for i, cid in enumerate(SUGS):
        r = rr + 3 + i
        pr = 5 + IDS.index(cid)
        col = 6 + 3 * i
        wr.cell(r, 1, CEN[cid]["nome"])
        wr.cell(r, 2, f"=Simulação!{L(col + 1)}{rt}")
        wr.cell(r, 3, f"=Simulação!B{rt}+B{r}")
        wr.cell(r, 4, f"=Premissas!H{pr}").number_format = "0.00"
        wr.cell(r, 5, f"=Premissas!I{pr}").number_format = "0.00"
        wr.cell(r, 6, f"=Premissas!K{pr}").number_format = "0.00"
        wr.cell(r, 7, f"=Simulação!{L(col + 2)}{rt}").number_format = "#,##0"
        wr.cell(r, 8, CEN[cid]["risco"])
    r = rr + 3 + len(SUGS) + 1
    wr.cell(r, 1, "6x1 operacional: manter seg–sáb")
    wr.cell(r, 2, f"=Simulação!B{rt + 6}")
    wr.cell(r + 1, 1, "Horário ADM → 5x2 seg–sex 8 h")
    wr.cell(r + 1, 2, 0)
    wr.cell(r + 3, 1, f"Prazos: {LEI['fase1']} = 2 folgas por semana e 42 h · {LEI['fase2']} = 40 h. Promulgação: {LEI['promulgacao']}.")
    larguras(wr, {"A": 46, "B": 11, "C": 22, "D": 13, "E": 11, "F": 11, "G": 14, "H": 12})

    # ---- Por turno
    wt = wb.create_sheet("Por turno")
    titulo(wt, "5x2 rotativa — turma G por horário",
           "Turmas A–F = grupos de folga de hoje (códigos do horário em ordem). Turma G = contratação.")
    cab(wt, 4, ["Grupo de horários", "Entrada", "Horas/dia", "A", "B", "C", "D", "E", "F", "Pessoas hoje",
                "Trabalhando por dia (média)", "Turma G (contratar)", "Pessoas na 5x2", "Trabalhando por dia na 5x2"])
    for i, t in enumerate(T24["turnos"]):
        r = 5 + i
        vals = [t["triade"], t["inicio"], round(t["h_dia"], 2)] + [t["turmas"].get(x, 0) for x in "ABCDEF"]
        for j, v in enumerate(vals, 1):
            wt.cell(r, j, v)
        wt.cell(r, 10, f"=SUM(D{r}:I{r})")
        wt.cell(r, 11, f"=J{r}*5/6").number_format = "0.0"
        wt.cell(r, 12, t["nova_turma_G"])
        wt.cell(r, 13, f"=J{r}+L{r}")
        wt.cell(r, 14, f"=M{r}*5/7").number_format = "0.0"
    r = 5 + len(T24["turnos"])
    wt.cell(r, 1, "Total").font = Font(bold=True)
    for j in range(4, 15):
        L = get_column_letter(j)
        wt.cell(r, j, f"=SUM({L}5:{L}{r - 1})").font = Font(bold=True)
        if j in (11, 14):
            wt.cell(r, j).number_format = "0.0"
    larguras(wt, {"A": 16, "B": 9, "C": 9, "J": 11, "K": 14, "L": 12, "M": 12, "N": 14})

    # ---- Por cargo (fracionário)
    wf = wb.create_sheet("Por cargo")
    titulo(wf, "Necessidade por cargo nas sugestões de 8 h (5x1 hoje)",
           "Valores fracionários. As sugestões de 12 h dependem das frotas 24 h e estão por família na aba Simulação.")
    cab(wf, 4, ["Cargo", "Família", "Pessoas hoje", "5x2 rotativa — contratar", "4x2 de 8 h — contratar"])
    fam_de = {p["funcao"]: p["familia"] for p in P}
    for i, f in enumerate(sorted(T24["por_funcao"], key=lambda x: -x["n"])):
        r = 5 + i
        wf.cell(r, 1, f["chave"])
        wf.cell(r, 2, fam_de.get(f["chave"]))
        wf.cell(r, 3, f["n"])
        for k, prem in enumerate([6, 7]):
            wf.cell(r, 4 + k, f"=C{r}*(Premissas!$H${prem}/Premissas!$H$5-1)").number_format = "0.0"
    larguras(wf, {"A": 42, "B": 28, "C": 12, "D": 16, "E": 16})
    wf.auto_filter.ref = f"A4:E{4 + len(T24['por_funcao'])}"
    wf.freeze_panes = "A5"

    # ---- 6x1, ADM, 12x36
    wo = wb.create_sheet("6x1 e ADM")
    titulo(wo, "6x1 operacional e horário ADM")
    cab(wo, 4, ["Código", "Horário", "Pessoas", "Horas/dia", "Horas/semana hoje", "Opção A — seg–sex (contratar)",
                "Horas/semana opção A", "Opção B — seg–sáb folga escalonada (contratar)"])
    for i, x in enumerate(OP6["por_codigo"]):
        r = 5 + i
        for j, v in enumerate([x["codigo"], x["descricao"], x["n"], round(x["h_dia"], 2), round(x["h_sem"], 2), 0,
                               35 if x["codigo"] == 300 else round(x["h_dia"] * 5, 2), f"=ROUNDUP(C{r}/5,0)"], 1):
            wo.cell(r, j, v)
    r = 6 + len(OP6["por_codigo"])
    cab(wo, r, ["Grupo ADM", "Códigos", "Pessoas", "Horas/semana hoje", "Horas/semana na 5x2", "Horas a menos por semana"])
    for k, (nome, d) in enumerate([("ADM com sábado (6x1)", ADM6), ("ADM sábado compensado (5x2)", ADM5)]):
        rr2 = r + 1 + k
        for j, v in enumerate([nome, ", ".join(f"{c} ({q})" for c, q in d["por_codigo"].items()), d["n"], 44, 40,
                               f"=C{rr2}*(D{rr2}-E{rr2})"], 1):
            wo.cell(rr2, j, v)
    larguras(wo, {"A": 30, "B": 38, "C": 10, "D": 12, "E": 14, "F": 18, "G": 14, "H": 22})

    # ---- Colaboradores
    wc = wb.create_sheet("Colaboradores")
    colunas = ["Matrícula", "Nome", "Cargo", "Família", "Centro de custo", "Área", "Situação", "Código atual",
               "Desde", "Código da escala", "Escala", "Grupo", "Regime hoje", "Entrada", "Horas/dia", "Horas/semana hoje",
               "Turma hoje (5x1)", "Proposta", "Horas/semana na proposta", "Observação"]
    cab(wc, 1, colunas)
    ordem = {g: i for i, (g, _, _) in enumerate(ORDEM_GRUPOS)}
    for p in sorted(P, key=lambda p: (ordem.get(p["grupo"], 99), p.get("inicio") or "", p.get("turma_atual") or "", p["nome"])):
        wc.append([p["matricula"], p["nome"], p["funcao"], p["familia"], p["cc"], p["area"], p["situacao"],
                   p.get("cod_atual"), p.get("desde"), p.get("cod_escala"), p.get("escala"), p["grupo"],
                   p.get("regime"), p.get("inicio"), round(p["h_dia"], 2) if p.get("h_dia") else None,
                   round(p["h_sem"], 2) if p.get("h_sem") else None, p.get("turma_atual"), p.get("proposta"),
                   p.get("h_sem_nova"), p.get("obs")])
    larguras(wc, {"A": 10, "B": 36, "C": 34, "D": 24, "E": 14, "F": 18, "G": 8, "H": 9, "I": 11, "J": 9, "K": 34,
                  "L": 26, "M": 9, "N": 8, "O": 8, "P": 10, "Q": 8, "R": 52, "S": 10, "T": 40})
    wc.freeze_panes = "C2"
    wc.auto_filter.ref = f"A1:T{len(P) + 1}"

    # ---- Códigos
    wk = wb.create_sheet("Códigos")
    cab(wk, 1, ["Código", "Descrição", "Grupo atribuído", "Regime", "Entrada", "Horas/dia", "Horas/semana (média)", "Pessoas"])
    cnt = collections.Counter(p.get("cod_atual") for p in P)
    for c in A["codigos"]:
        wk.append([c["codigo"], c["descricao"], c["grupo"], c["regime"], c.get("inicio"),
                   round(c["h_dia"], 2) if c.get("h_dia") else None, round(c["h_sem"], 2) if c.get("h_sem") else None,
                   cnt.get(c["codigo"], 0)])
    larguras(wk, {"A": 9, "B": 44, "C": 30, "D": 10, "E": 9, "F": 9, "G": 12, "H": 9})
    wk.freeze_panes = "A2"
    wk.auto_filter.ref = f"A1:H{len(A['codigos']) + 1}"

    # ---- Modelos de escala
    for nome, chave, rot in [("Modelo 5x2", "modelo_5x2", [f"Turma {x}" for x in "ABCDEFG"]),
                             ("Modelo 4x2", "modelo_4x2", [f"Turma {x}" for x in "ABC"]),
                             ("Modelo 12x36", "modelo_12x36", ["Turma A (dia)", "Turma B (dia)", "Turma C (noite)", "Turma D (noite)"]),
                             ("Modelo 3x3", "modelo_3x3", ["Turma A (dia)", "Turma B (dia)", "Turma C (noite)", "Turma D (noite)"]),
                             ("Modelo 4x2 12h", "modelo_4x2_12", ["Turma A (dia)", "Turma B (dia)", "Turma C (dia)",
                                                                  "Turma D (noite)", "Turma E (noite)", "Turma F (noite)"])]:
        wm = wb.create_sheet(nome)
        g = A[chave]["grade"]
        dias = len(g[0])
        titulo(wm, f"{nome.replace('Modelo ', 'Escala ')} — F = folga, T = trabalho (semana de segunda a domingo)")
        wm.cell(3, 1, "Dia")
        for d in range(dias):
            c = wm.cell(3, 2 + d, "STQQSSD"[d % 7])
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="B3302D" if d % 7 == 6 else AZUL)
            c.alignment = Alignment(horizontal="center")
            wm.column_dimensions[get_column_letter(2 + d)].width = 3.2
        for t, linha in enumerate(g):
            wm.cell(4 + t, 1, rot[t]).font = Font(bold=True)
            for d, v in enumerate(linha):
                c = wm.cell(4 + t, 2 + d, v)
                c.alignment = Alignment(horizontal="center")
                if v == "F":
                    c.fill = PatternFill("solid", fgColor="B3302D" if d % 7 == 6 else "1E6B4A")
                    c.font = Font(bold=True, color="FFFFFF")
                else:
                    c.fill = PatternFill("solid", fgColor="DFE5EC")
        rr3 = 4 + len(g)
        wm.cell(rr3, 1, "Trabalhando").font = Font(bold=True)
        for d in range(dias):
            L = get_column_letter(2 + d)
            wm.cell(rr3, 2 + d, f'=COUNTIF({L}4:{L}{rr3 - 1},"T")').alignment = Alignment(horizontal="center")
        wm.column_dimensions["A"].width = 16

    wb.save(XLSX)


def main():
    for f in (p1, p2, p3, p4, p5, p5b, p6, p7, p8):
        f()
    html = BUILD / "relatorio.html"
    html.write_text(montar(), encoding="utf-8")
    tmp = BUILD / "relatorio.pdf"
    tmp.unlink(missing_ok=True)
    perfil = BUILD / "edge_perfil"  # perfil próprio: não briga com um Edge já aberto
    subprocess.run([EDGE, "--headless=new", "--disable-gpu", "--no-pdf-header-footer", f"--user-data-dir={perfil}",
                    "--virtual-time-budget=5000", f"--print-to-pdf={tmp}", html.resolve().as_uri()], check=True, timeout=180)
    tmp.replace(PDF)  # falha se o PDF estiver aberto em outro programa
    base_xlsx()
    print(PDF, XLSX, gerar_html.gerar(A))


if __name__ == "__main__":
    main()
