"""Pranchas A3 do estudo de layout (PDF vetorial multipágina + SVG e PNG por folha).

python pranchas.py  ->  Pranchas_Sala_Motores_OUT2026.pdf, build/SM-0x.svg, build/SM-0x.png
Plantas em 1:75 e detalhe da lavagem em 1:40 no papel A3, ocupando a largura útil da folha;
legendas na coluna da direita e quadros na faixa de baixo. A escala gráfica vale em qualquer impressão.
"""
import json
import math
import textwrap
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.font_manager as fm  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.backends.backend_pdf import PdfPages  # noqa: E402
from matplotlib.patches import Arc, Circle, FancyArrowPatch, Polygon, Rectangle  # noqa: E402

import galpao as G  # noqa: E402
import layout as LY  # noqa: E402
import verificacao as V  # noqa: E402

AQUI = Path(__file__).parent
BUILD = AQUI / "build"
BUILD.mkdir(exist_ok=True)
PDF = AQUI / "Pranchas_Sala_Motores_OUT2026.pdf"
DATA = "21/10/2026"
REV = "A — estudo preliminar"
NFOLHAS = 7

for f in ("ARIALN.TTF", "ARIALNB.TTF"):
    fm.fontManager.addfont(f"C:/Windows/Fonts/{f}")
plt.rcParams.update({"font.family": "Arial Narrow", "font.size": 6, "svg.fonttype": "none",
                     "pdf.fonttype": 42, "hatch.linewidth": 0.3, "lines.solid_capstyle": "butt"})

PRETO, CINZA, CINZA2 = "#1a1a1a", "#8c8c8c", "#c8c8c8"
CONFIRMAR = "#c2185b"      # tudo que é premissa ou depende de confirmação
W_A3, H_A3 = 420, 297

# moldura útil: x 20–410, y 10–287; carimbo x 245–410, y 10–62
ESC = 75                       # plantas
XL = (-1800, 21800)            # janela das plantas no galpão (mm)
PL_X, PL_TOPO = 22, 281        # planta encostada no topo e na esquerda
COL_X = 342                    # coluna da direita (legendas)
COL_L = 66
BX0, BX1 = 22, 243             # faixa de baixo, à esquerda do carimbo
BD_X = 249                     # faixa de baixo, sobre o carimbo (até 408)

CURTO = {"S1": "Recebimento e drenagem", "S2": "Desmontagem de motores", "S3": "Desmontagem de transmissões",
         "S4": "Lavagem de peças pequenas", "S5": "Lavagem industrial", "S6": "Escorrimento e secagem",
         "S7": "Inspeção dimensional", "S8": "Montagem de transmissões", "S9": "Montagem de motores",
         "S10": "CQ e expedição", "S11": "Escritório e metrologia", "S12": "Ferramentaria e apoio",
         "S13": "Faixas do pórtico (restrito)"}


def br(v, d=0):
    s = f"{v:,.{d}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def chars(larg_mm, fs):
    """Quantos caracteres de Arial Narrow cabem numa largura."""
    return max(10, int(larg_mm / (fs * 0.3528 * 0.44)))


# ---------------------------------------------------------------- papel
def folha(num, titulo, escala):
    fig = plt.figure(figsize=(W_A3 / 25.4, H_A3 / 25.4))
    p = fig.add_axes([0, 0, 1, 1])
    p.set_xlim(0, W_A3)
    p.set_ylim(0, H_A3)
    p.axis("off")
    p.add_patch(Rectangle((20, 10), 390, 277, fill=False, lw=0.8, ec=PRETO))
    carimbo(p, num, titulo, escala)
    return fig, p


def carimbo(p, num, titulo, escala):
    x0, y0, w, h = 245, 10, 165, 52
    p.add_patch(Rectangle((x0, y0), w, h, fill=False, lw=0.8, ec=PRETO))
    for yy in (y0 + 10, y0 + 19, y0 + 31, y0 + 42):
        p.plot([x0, x0 + w], [yy, yy], lw=0.4, c=PRETO)
    p.plot([x0 + 110, x0 + 110], [y0, y0 + 19], lw=0.4, c=PRETO)
    p.plot([x0 + 55, x0 + 55], [y0, y0 + 10], lw=0.4, c=PRETO)
    t = lambda x, y, s, **k: p.text(x, y, s, va="center", **k)  # noqa: E731
    t(x0 + 2, y0 + 47, "Oficina de motores e transmissões com lavagem de componentes", fontsize=8.5, weight="bold")
    t(x0 + 2, y0 + 36.5, titulo, fontsize=9.5, weight="bold")
    t(x0 + 2, y0 + 27.5, "Local: CRV Açúcar & Álcool — Capinópolis/MG.  Base: desenho CRV(MG)20102026 rev. 00 (16/10/2026)",
      fontsize=6)
    t(x0 + 2, y0 + 22.5, "Responsável técnico: a definir (profissional habilitado, com ART/RRT)", fontsize=6)
    t(x0 + 2, y0 + 15.5, "ESTUDO CONCEITUAL — NÃO EXECUTIVO — NÃO LIBERADO PARA OBRA", fontsize=7.5, weight="bold",
      color="#b3302d")
    t(x0 + 2, y0 + 11.8, "Dimensões de equipamentos são envelopes de reserva; itens em magenta dependem de confirmação.",
      fontsize=5.3, color=CONFIRMAR)
    t(x0 + 112, y0 + 15.5, f"Escala: {escala}", fontsize=6.5)
    t(x0 + 112, y0 + 11.8, "Cotas em mm; alturas em mm acima do piso", fontsize=5.5)
    t(x0 + 2, y0 + 5, f"Data: {DATA}", fontsize=6.5)
    t(x0 + 57, y0 + 5, f"Revisão: {REV}", fontsize=6.5)
    t(x0 + 112, y0 + 5, f"Folha {num} de SM-0{NFOLHAS}", fontsize=8, weight="bold")


def ax_escala(fig, x, y, xlim, ylim, escala):
    """Eixo com escala exata no papel (x, y em mm no A3; limites em mm do galpão)."""
    w = (xlim[1] - xlim[0]) / escala
    h = (ylim[1] - ylim[0]) / escala
    ax = fig.add_axes([x / W_A3, y / H_A3, w / W_A3, h / H_A3])
    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)
    ax.set_aspect("equal")
    ax.axis("off")
    ax._escala = escala
    return ax


def planta(fig, yl=(-1800, 9700), xl=XL, esc=ESC):
    """Planta encostada no topo e à esquerda; devolve o eixo e o y da base no papel."""
    h = (yl[1] - yl[0]) / esc
    return ax_escala(fig, PL_X, PL_TOPO - h, xl, yl, esc), PL_TOPO - h


def mm_txt(ax, pt):
    """Converte tamanho em pt do papel para mm do desenho (para afastar textos)."""
    return pt * 0.3528 * ax._escala


def escala_grafica(p, x, y, escala, metros=5, rotulo=None):
    passo = 1000 / escala
    for i in range(metros):
        p.add_patch(Rectangle((x + i * passo, y), passo, 1.4, fc=PRETO if i % 2 == 0 else "white", ec=PRETO, lw=0.4))
        p.text(x + i * passo, y - 1.4, str(i), ha="center", va="top", fontsize=5.2)
    p.text(x + metros * passo, y - 1.4, f"{metros} m", ha="center", va="top", fontsize=5.2)
    p.text(x, y + 2.6, rotulo or f"Escala gráfica 1:{escala} (A3)", fontsize=5.3)


def titulo_vista(p, x, y, txt, escala=None, fs=8.5):
    t = p.text(x, y, txt, fontsize=fs, weight="bold")
    if escala:
        p.text(x, y - 3.2, f"ESCALA 1:{escala}", fontsize=5.3)
    return t


def bloco(p, x, y, titulo, linhas, larg, fs=5.3, passo=2.55, tit_fs=7):
    """Texto com quebra automática; linha iniciada com * sai em magenta (a confirmar). Devolve o y final."""
    if titulo:
        p.text(x, y, titulo, fontsize=tit_fs, weight="bold")
        y -= 4.0
    n = chars(larg, fs)
    for ln in linhas:
        cor = CONFIRMAR if ln.startswith("*") else PRETO
        txt = ln.lstrip("*")
        recuo = len(txt) - len(txt.lstrip(" "))
        partes = textwrap.wrap(txt.strip(), n - recuo) or [""]
        for k, parte in enumerate(partes):
            p.text(x + recuo * 0.8 + (1.8 if k and txt.strip()[:1] in "•–" else 0), y, parte, fontsize=fs,
                   va="top", color=cor)
            y -= passo
    return y


def lista(p, x, y, itens, titulo, larg, fs=5.2, passo=2.5, cor_tag=PRETO):
    """Itens (tag, texto) com quebra de linha; devolve o y final."""
    p.text(x, y, titulo, fontsize=7, weight="bold")
    yy = y - 4.0
    n = chars(larg - 7, fs)
    for tag, txt in itens:
        p.text(x, yy, tag, fontsize=fs, weight="bold", va="top", color=cor_tag)
        for ln in textwrap.wrap(txt, n):
            p.text(x + 7, yy, ln, fontsize=fs, va="top")
            yy -= passo
        yy -= 0.5
    return yy


def tabela(p, x, y, cols, linhas, fs=5.2, passo=2.45, titulo=None, cor_linha=None):
    """Tabela com quebra de texto por célula. cols = larguras em mm. Primeira linha = cabeçalho."""
    if titulo:
        p.text(x, y, titulo, fontsize=7, weight="bold")
        y -= 4.2
    for k, ln in enumerate(linhas):
        partes = [textwrap.wrap(str(c), chars(w - 1.5, fs)) or [""] for c, w in zip(ln, cols)]
        alt = max(len(pp) for pp in partes)
        xx = x
        for c, (pp, w) in enumerate(zip(partes, cols)):
            for i, s in enumerate(pp):
                p.text(xx, y - i * passo, s, fontsize=fs, va="top", weight="bold" if k == 0 or c == 0 and k else "normal",
                       color=CONFIRMAR if s.endswith("*") else (cor_linha(k) if cor_linha else PRETO))
            xx += w
        y -= alt * passo + 0.8
        p.plot([x, x + sum(cols)], [y + 0.4, y + 0.4], lw=0.25 if k else 0.5, c=CINZA2 if k else PRETO)
    return y


# ---------------------------------------------------------------- galpão existente
def rect(ax, r, **k):
    ax.add_patch(Rectangle((r[0], r[2]), r[1] - r[0], r[3] - r[2], **k))


def paredes():
    e = G.PAREDE / 2
    L, B = G.L, G.B
    return [(-e, e, -e, G.PORTA_FUNDOS[0]), (-e, e, G.PORTA_FUNDOS[1], B + e),
            (L - e, L + e, -e, G.PORTAO[0]), (L - e, L + e, G.PORTAO[1], B + e),
            (-e, G.ABERTURA_LATERAL[0], -e, e), (G.ABERTURA_LATERAL[1], L + e, -e, e),
            (-e, L + e, B - e, B + e)]


def paredes_escritorio():
    x0, x1, y0, y1 = G.ESCRITORIO
    p = G.PAREDE
    return [(x0, x1, y0, y0 + p), (x0, x1, y1 - p, y1), (x1 - p, x1, y0, G.ESCRITORIO_PORTA[0]),
            (x1 - p, x1, G.ESCRITORIO_PORTA[1], y1)]


def desenhar_galpao(ax, trilhos="proposto", eixos=True, blocos=True, rotulos=True):
    if eixos:
        for x in G.PILARES_X:
            ax.plot([x, x], [-1300, G.B + 1300], lw=0.25, c=CINZA, ls=(0, (12, 3, 2, 3)), zorder=1)
        for y in (0, G.B):
            ax.plot([-1300, G.L + 1300], [y, y], lw=0.25, c=CINZA, ls=(0, (12, 3, 2, 3)), zorder=1)
        if rotulos:
            r = mm_txt(ax, 6)
            for i, x in enumerate(G.PILARES_X, 1):
                for y in (-1300 - r, G.B + 1300 + r):
                    ax.add_patch(Circle((x, y), r, fill=False, lw=0.4, ec=PRETO, zorder=5))
                    ax.text(x, y, str(i), ha="center", va="center", fontsize=6.5, zorder=6)
            for nm, y in (("A", 0), ("B", G.B)):
                for x in (-1300 - r, G.L + 1300 + r):
                    ax.add_patch(Circle((x, y), r, fill=False, lw=0.4, ec=PRETO, zorder=5))
                    ax.text(x, y, nm, ha="center", va="center", fontsize=6.5, zorder=6)
    for r in paredes() + paredes_escritorio():
        rect(ax, r, fc="#9c9c9c", ec=PRETO, lw=0.35, zorder=4)
    if blocos:
        h = G.BLOCO / 2
        for x in G.PILARES_X:
            for y in (0, G.B):
                rect(ax, (x - h, x + h, y - h, y + h), fc="none", ec=PRETO, lw=0.3, hatch="////", zorder=3)
    for x in G.PILARES_X:
        for y in (0, G.B):
            px, py = G.PILAR
            rect(ax, (x - px / 2, x + px / 2, y - py / 2, y + py / 2), fc=PRETO, ec=PRETO, lw=0.3, zorder=6)
    y0, y1 = G.PORTA_FUNDOS
    ax.add_patch(Arc((0, y0), 2 * (y1 - y0), 2 * (y1 - y0), theta1=0, theta2=90, lw=0.3, ls="--", ec=PRETO, zorder=5))
    ax.plot([0, y1 - y0], [y0, y0], lw=0.7, c=PRETO, zorder=5)
    d0, d1 = G.ESCRITORIO_PORTA
    xh = G.ESCRITORIO[1] - G.PAREDE / 2
    ax.add_patch(Arc((xh, d1), 2 * (d1 - d0), 2 * (d1 - d0), theta1=180, theta2=270, lw=0.3, ls="--", ec=PRETO, zorder=5))
    ax.plot([xh - (d1 - d0), xh], [d1, d1], lw=0.7, c=PRETO, zorder=5)
    for dx in (-35, 30):
        ax.plot([G.L + dx, G.L + dx], list(G.PORTAO), lw=0.4, c=PRETO, zorder=5)
    ax.plot(list(G.ABERTURA_LATERAL), [0, 0], lw=0.4, c=CONFIRMAR, zorder=5)
    xs = G.TRILHOS_X if trilhos == "original" else LY.TRILHOS_PROPOSTOS_X if trilhos == "proposto" else None
    if xs:
        for y in G.TRILHOS_Y:
            for dy in (-25, 25):
                ax.plot(list(xs), [y + dy, y + dy], lw=0.3, c="#5a5a5a", zorder=3)


def rotulos_galpao(ax, completo=False, fs=5.8):
    ax.text(G.L + 230, (G.PORTAO[0] + G.PORTAO[1]) / 2, "PORTÃO DE ARRASTE 4.000 x 5.000", rotation=90,
            ha="left", va="center", fontsize=fs, zorder=7)
    ax.text(-230, (G.PORTA_FUNDOS[0] + G.PORTA_FUNDOS[1]) / 2, "PORTA 1.000", rotation=90, ha="right",
            va="center", fontsize=fs, zorder=7)
    ax.text((G.ABERTURA_LATERAL[0] + G.ABERTURA_LATERAL[1]) / 2, -260, "vão 2.250 a identificar (D08)",
            ha="center", va="top", fontsize=fs, color=CONFIRMAR, zorder=7)
    if completo:
        ax.text(2075, 5400, "ESCRITÓRIO\n(existente)", ha="center", va="center", fontsize=7, zorder=7)


# ---------------------------------------------------------------- camadas do estudo
ROTULO_SETOR = {"S5": "dir", "S12": (800, 2030, "left", "bottom")}   # onde o canto superior esquerdo tem equipamento


def desenhar_setores(ax, setores, rotulo=True, alfa=1.0, fs=7):
    for sid, nome, zona, rets, _ in setores:
        fc, ec, _ = LY.ZONAS[zona]
        for r in rets:
            if zona == "exclusao":
                rect(ax, r, fc=fc, ec=ec, lw=0.3, hatch="\\\\\\\\", alpha=alfa, zorder=2)
            else:
                rect(ax, r, fc=fc, ec=ec, lw=0.5, alpha=alfa, zorder=2)
        if rotulo and zona not in ("exclusao", "escritorio"):
            r = rets[0]
            pos = ROTULO_SETOR.get(sid) if setores is LY.A1_SETORES else None
            x, y, ha, va = (r[0] + 70, r[3] - 70, "left", "top") if pos is None else \
                (r[1] - 70, r[3] - 70, "right", "top") if pos == "dir" else pos
            ax.text(x, y, sid, ha=ha, va=va, fontsize=fs, weight="bold", color=ec, zorder=8)


def desenhar_equip(ax, equip, carro=None, fs=5.6, rot=True, so_tag=False):
    for tag, r, curto in equip:
        rect(ax, r, fc="white", ec=PRETO, lw=0.5, zorder=6)
        if tag in ("E04", "E17"):   # motor no cavalete
            cx, cy = (r[0] + r[1]) / 2, (r[2] + r[3]) / 2
            rect(ax, (cx - 600, cx + 600, cy - 350, cy + 350), fc="none", ec=PRETO, lw=0.3, ls="--", zorder=6)
        if tag == "E07":
            ax.plot([r[0], r[1]], [r[2], r[3]], lw=0.25, c=PRETO, zorder=6)
            ax.plot([r[0], r[1]], [r[3], r[2]], lw=0.25, c=PRETO, zorder=6)
        if rot:
            cx, cy = (r[0] + r[1]) / 2, (r[2] + r[3]) / 2
            vert = (r[3] - r[2]) > 1.6 * (r[1] - r[0])
            txt = tag if (so_tag or not curto) else f"{tag}\n{curto}"
            ax.text(cx, cy, txt, ha="center", va="center", fontsize=fs, rotation=90 if vert else 0, zorder=7,
                    linespacing=0.95, color=PRETO if not so_tag else "#444",
                    bbox=dict(boxstyle="square,pad=0.08", fc="white", ec="none") if tag == "E07" else None)
    if carro:
        rect(ax, carro, fc="none", ec=PRETO, lw=0.4, ls=(0, (4, 2)), zorder=6)
        if rot and not so_tag:
            ax.text((carro[0] + carro[1]) / 2, carro[2] + 100, "carro extraível", ha="center", va="bottom",
                    fontsize=fs - 0.6, zorder=7)


def desenhar_portico(ax, posicao=14700, cobertura=True, fs=5.6):
    gx, gy = LY.GANCHO_X, LY.GANCHO_Y
    if cobertura:
        rect(ax, (gx[0], gx[1], gy[0], gy[1]), fc="none", ec="#1565c0", lw=0.6, ls=(0, (6, 3)), zorder=9)
        ax.text(gx[0] + 80, gy[1] + 50, f"alcance do gancho X {br(gx[0])}–{br(gx[1])}, Y {br(gy[0])}–{br(gy[1])} (premissa)",
                fontsize=fs, color="#1565c0", va="bottom", zorder=9)
    for (y0, y1) in LY.PERNA_FAIXA:
        rect(ax, (LY.PERNAS_X[0], LY.PERNAS_X[1], y0, y1), fc="none", ec="#c9a400", lw=0.4, ls=(0, (2, 2)), zorder=3)
    if posicao:
        h = LY.TRUQUE_X / 2
        for (y0, y1) in LY.PERNA_FAIXA:
            rect(ax, (posicao - h, posicao + h, y0, y1), fc="#fff3b0", ec="#7a6200", lw=0.6, zorder=9)
        ax.plot([posicao, posicao], [LY.PERNA_FAIXA[0][0], LY.PERNA_FAIXA[1][1]], lw=1.0, c="#7a6200",
                ls=(0, (8, 2, 2, 2)), zorder=9)
        ax.text(posicao + 560, LY.PERNA_FAIXA[1][0] - 30, "E01 pórtico\n(posição ilustrativa)", fontsize=fs - 0.3,
                color="#7a6200", va="top", zorder=9)
    for x in LY.PERNAS_X:
        for (y0, y1) in LY.PERNA_FAIXA:
            ax.plot([x, x], [y0 - 80, y1 + 80], lw=1.2, c="#7a6200", zorder=9)
    ax.text(LY.PERNAS_X[0] - 80, LY.PERNA_FAIXA[0][0] - 50, "batente", ha="right", va="top", fontsize=fs - 0.5,
            color=CONFIRMAR, zorder=9)


# ---------------------------------------------------------------- cotas e setas
def cota_h(ax, x0, x1, y, txt=None, ref=None, fs=5.6, cor=PRETO):
    ax.plot([x0, x1], [y, y], lw=0.3, c=cor, zorder=10)
    t = mm_txt(ax, 2.2)
    for x in (x0, x1):
        ax.plot([x - t / 2, x + t / 2], [y - t / 2, y + t / 2], lw=0.6, c=cor, zorder=10)
        if ref is not None:
            ax.plot([x, x], [ref, y + (t if y > ref else -t)], lw=0.2, c=cor, zorder=10)
    ax.text((x0 + x1) / 2, y + mm_txt(ax, 0.8), txt or br(abs(x1 - x0)), ha="center", va="bottom", fontsize=fs,
            color=cor, zorder=10)


def cota_v(ax, y0, y1, x, txt=None, ref=None, fs=5.6, cor=PRETO):
    ax.plot([x, x], [y0, y1], lw=0.3, c=cor, zorder=10)
    t = mm_txt(ax, 2.2)
    for y in (y0, y1):
        ax.plot([x - t / 2, x + t / 2], [y - t / 2, y + t / 2], lw=0.6, c=cor, zorder=10)
        if ref is not None:
            ax.plot([ref, x + (t if x > ref else -t)], [y, y], lw=0.2, c=cor, zorder=10)
    ax.text(x - mm_txt(ax, 0.8), (y0 + y1) / 2, txt or br(abs(y1 - y0)), ha="right", va="center", rotation=90,
            fontsize=fs, color=cor, zorder=10)


def cadeia_h(ax, xs, y, ref, fs=5.6):
    for a, b in zip(xs, xs[1:]):
        cota_h(ax, a, b, y, ref=ref, fs=fs)


def cadeia_v(ax, ys, x, ref, fs=5.6):
    for a, b in zip(ys, ys[1:]):
        cota_v(ax, a, b, x, ref=ref, fs=fs)


def seta(ax, pts, cor, ls="-", lw=1.1, z=11):
    for i, (a, b) in enumerate(zip(pts, pts[1:])):
        ultimo = i == len(pts) - 2
        ax.add_patch(FancyArrowPatch(a, b, arrowstyle="-|>" if ultimo else "-", mutation_scale=7, lw=lw,
                                     color=cor, ls=ls, shrinkA=0, shrinkB=0, zorder=z))
        if not ultimo and math.dist(a, b) > 2500:
            m = ((a[0] + b[0]) / 2, (a[1] + b[1]) / 2)
            d = ((b[0] - a[0]) / math.dist(a, b) * 10, (b[1] - a[1]) / math.dist(a, b) * 10)
            ax.add_patch(FancyArrowPatch(m, (m[0] + d[0], m[1] + d[1]), arrowstyle="-|>", mutation_scale=7,
                                         lw=0, color=cor, zorder=z))


# ---------------------------------------------------------------- legendas
def legenda_setores(p, x, y, setores, n_alt, larg=COL_L, titulo="Setores e áreas"):
    p.text(x, y, titulo, fontsize=7, weight="bold")
    yy = y - 4.2
    q = {i: a for i, _, _, a, _ in V.quadro_areas(n_alt)}
    for sid, nome, zona, rets, _ in setores:
        fc, ec, _ = LY.ZONAS[zona]
        p.add_patch(Rectangle((x, yy - 1.2), 3.6, 2.4, fc=fc, ec=ec, lw=0.4, hatch="\\\\\\" if zona == "exclusao" else None))
        a = q.get(sid, sum(V.area(r) for r in rets))
        rot = CURTO[sid] + (" (2 posições)" if "2 posições" in nome else "")
        p.text(x + 4.8, yy, f"{sid}  {rot}", fontsize=5.4, va="center")
        p.text(x + larg, yy, f"{br(a, 2)} m²", fontsize=5.4, va="center", ha="right")
        yy -= 3.25
    circ = [a for i, nm, z, a, pc in V.quadro_areas(n_alt) if z == "circulacao"][0]
    p.add_patch(Rectangle((x, yy - 1.2), 3.6, 2.4, fc="white", ec=CINZA, lw=0.4))
    p.text(x + 4.8, yy, "Circulação e acessos", fontsize=5.4, va="center")
    p.text(x + larg, yy, f"{br(circ, 2)} m²", fontsize=5.4, va="center", ha="right")
    yy -= 3.25
    p.plot([x, x + larg], [yy + 1.6, yy + 1.6], lw=0.4, c=PRETO)
    p.text(x + 4.8, yy, "Área interna 19,85 x 7,70 m", fontsize=5.4, va="center", weight="bold")
    p.text(x + larg, yy, f"{br(G.AREA_INTERNA, 2)} m²", fontsize=5.4, va="center", ha="right", weight="bold")
    return yy - 4.5


def tabela_equip(p, x, y, equip, colunas=2, larg=110, passo=2.95, titulo="Equipamentos (envelope de reserva, mm)"):
    p.text(x, y, titulo, fontsize=7, weight="bold")
    tags = []
    for t, _, _ in equip:
        if t not in tags:
            tags.append(t)
    qtd = {t: sum(1 for e in equip if e[0] == t) for t in tags}
    por_col = math.ceil(len(tags) / colunas)
    n = chars(larg - 32, 5.3)
    for i, t in enumerate(tags):
        c, l = divmod(i, por_col)
        xx, yy = x + c * larg, y - 4.2 - l * passo
        nome = LY.EQUIP[t][0]
        nome = nome if len(nome) <= n else nome[:n - 1] + "…"
        p.text(xx, yy, t, fontsize=5.4, weight="bold", va="center")
        p.text(xx + 6.5, yy, nome, fontsize=5.3, va="center")
        p.text(xx + larg - 3, yy, f"{qtd[t]}x  {LY.EQUIP[t][3].split(' (')[0][:22]}", fontsize=5.2, va="center",
               ha="right")
    return y - 4.2 - por_col * passo


def legenda_simbolos(p, x, y, itens, titulo="Convenções", fs=5.4, passo=3.25):
    p.text(x, y, titulo, fontsize=7, weight="bold")
    yy = y - 4.2
    for desenho, txt in itens:
        desenho(p, x, yy)
        p.text(x + 8.5, yy, txt, fontsize=fs, va="center")
        yy -= passo
    return yy - 1.5


def sim_linha(cor, ls="-", lw=0.8):
    return lambda p, x, y: p.plot([x, x + 6.5], [y, y], c=cor, ls=ls, lw=lw)


def sim_ret(fc, ec, hatch=None, ls="-"):
    return lambda p, x, y: p.add_patch(Rectangle((x, y - 1.2), 6.5, 2.4, fc=fc, ec=ec, lw=0.5, hatch=hatch, ls=ls))


def pdf_salvar(pdf, fig, nome):
    for ax in fig.axes[1:]:          # textos das vistas não vazam para fora da janela
        for t in ax.texts:
            t.set_clip_on(True)
    pdf.savefig(fig)
    fig.savefig(BUILD / f"{nome}.svg")
    fig.savefig(BUILD / f"{nome}.png", dpi=200)
    plt.close(fig)


def base_alt1(ax, setores=True, equip=True, portico=True):
    if setores:
        desenhar_setores(ax, LY.A1_SETORES)
    desenhar_galpao(ax)
    rotulos_galpao(ax)
    if portico:
        desenhar_portico(ax)
    if equip:
        desenhar_equip(ax, LY.A1_EQUIP, carro=LY.A1_CARRO_LAVADORA)
    ax.text(2075, 5350, "S11\nESCRITÓRIO\nTÉCNICO E\nMETROLOGIA", ha="center", va="center", fontsize=6.3, zorder=8,
            color="#8a7d4f", weight="bold")


# ---------------------------------------------------------------- SM-01 existente
def f0_existente(pdf):
    fig, p = folha("SM-01", "Levantamento do galpão existente e interferências", "1:75 planta; 1:100 corte (A3)")
    ax, yb = planta(fig, yl=(-2700, 9700))
    orig = json.loads((AQUI / "desenho_original.json").read_text(encoding="utf-8"))
    for s in orig["planta_segmentos_mm"]:
        ax.plot([s[0][0], s[1][0]], [s[0][1], s[1][1]], lw=0.4, c="#7fb3e0", zorder=1)
    desenhar_galpao(ax, trilhos="original")
    rotulos_galpao(ax, completo=True)
    for nm, x0, x1 in G.AREAS_DEMARCADAS:
        rect(ax, (x0, x1, G.AREAS_Y[0], G.AREAS_Y[1]), fc="none", ec="#555", lw=0.6, ls=(0, (3, 2)), zorder=4)
        ax.text((x0 + x1) / 2, 2600, f"{nm}\n(demarcação, D11)", ha="center", fontsize=6, zorder=7)
    for y in G.TRILHOS_Y:
        ax.text(9000, y + 110, "trilho do pórtico (projeto atual)", fontsize=5.6, zorder=7)
    cadeia_h(ax, G.PILARES_X, -1850, ref=-450)
    cota_h(ax, 0, G.L, -2450, txt="20.000 (entre eixos)")
    cota_v(ax, 0, G.B, -1000, txt="7.850 (entre eixos)", ref=-450)
    cadeia_v(ax, [0, G.PORTAO[0], G.PORTAO[1], G.B], G.L + 1000, ref=G.L + 200)
    cota_v(ax, *G.TRILHOS_Y, 6500, txt="6.534 (trilhos)")
    cota_h(ax, 75, 4075, 7300, txt="4.000")
    cota_v(ax, 3925, 6925, 4500, txt="3.000")
    cota_h(ax, *G.ABERTURA_LATERAL, -1000, txt="2.250", cor=CONFIRMAR)
    cota_v(ax, 6925, G.B - 75, 2500, txt="850")
    cota_v(ax, 0, G.TRILHOS_Y[0], 3000, txt="658")
    marcas = [("I01", 10000, 520), ("I02", 12000, 658), ("I03", 7000, 7192), ("I05", 1000, 5000),
              ("I08", 600, 1350), ("I09", 19400, 3925), ("I10", 18600, 300), ("D10", 3500, 7300),
              ("D01", 15000, -520), ("D08", 17000, -260)]
    for tg, x, y in marcas:
        cor = CONFIRMAR if tg.startswith("D") else "#0d47a1"
        ax.text(x, y, tg, fontsize=6, weight="bold", color="white", ha="center", va="center", zorder=12,
                bbox=dict(boxstyle="round,pad=0.15", fc=cor, ec="none"))
    titulo_vista(p, PL_X, 283.5, "PLANTA DO EXISTENTE — redesenho deste estudo (cinza/preto) sobre as linhas do PDF original "
                 "(azul-claro)   ·   ESCALA 1:75")
    # coluna da direita: divergências
    yy = lista(p, COL_X, 280, [(d, t) for d, t, _ in G.DIVERGENCIAS], "Divergências e pontos sem informação (D)",
               larg=COL_L, cor_tag=CONFIRMAR, fs=5.1, passo=2.4)
    yy = tabela(p, COL_X, yy - 3, [36, 12, 18], [("Altura acima do piso", "mm", "fonte")] +
                [(k, br(v), "cota" if "cota" in f else "medido*") for k, (v, f) in G.ALTURAS.items()],
                titulo="Alturas", fs=5.1, passo=2.35)
    tabela(p, COL_X, yy - 3, [44, 22], [("Área", "m²"), ("Entre eixos (20,00 x 7,85)", br(G.AREA_EIXOS, 2)),
                                        ("Construída (faces externas)", br(G.AREA_CONSTRUIDA, 2)),
                                        ("Interna (faces internas)", br(G.AREA_INTERNA, 2)),
                                        ("Escritório (com paredes)", br(G.AREA_ESCRITORIO, 2))],
           titulo="Áreas", fs=5.1, passo=2.35)
    # faixa de baixo: corte, dados, interferências
    ac = ax_escala(fig, BX0, 24, (-1500, 9600), (-900, 6900), 100)
    corte_transversal(ac, completo=False)
    titulo_vista(p, BX0, yb - 6, "SECÇÃO S1-S1 (simplificada)", 100)
    bloco(p, 137, yb - 6, "Dados do desenho original", [
        "Confirmadas (cota no desenho, conferidas nas linhas a ±5 mm): 20.000 x 7.850 entre eixos; vãos de 5.000; "
        "paredes de 150; portão 4.000 x 5.000 (1.925 + 4.000 + 1.925); porta 1.000; escritório 4.000 x 3.000; "
        "trilhos com vão de 6.534; viga do pórtico a 4.460; laje de 100.",
        f"Áreas: entre eixos {br(G.AREA_EIXOS, 2)} m²; construída (faces externas) {br(G.AREA_CONSTRUIDA, 2)} m²; "
        f"interna {br(G.AREA_INTERNA, 2)} m²; escritório {br(G.AREA_ESCRITORIO, 2)} m² com paredes.",
        "*Medidas tiradas das linhas (sem cota): forro ≈ 5.380 acima do piso; base da viga ≈ 4.150; perna do pórtico "
        "≈ 520; blocos 700 x 700 com topo 150 acima do piso.",
        "Alturas acima do piso: viga 4,15–4,46 m; forro ≈ 5,38 m; beiral ≈ 5,42 m; cumeeira ≈ 6,27 m.",
        "Piso: laje 100 mm, 25 MPa, tela 15 x 15 fio 3,8, juntas a cada 3 m. Fundação: bloco 600 x 600 x 500 "
        "sobre estaca Ø 250 x 4.000. Pórtico 2.000 kg (viga W 310 x 44,5) com talha elétrica 2.000 kg.",
        "Eixos 1–5 e A–B nomeados neste estudo. Norte não indicado no original.",
    ], larg=BX1 - 137, fs=5.2, passo=2.45)
    lista(p, BD_X, yb - 6, G.INTERFERENCIAS, "Interferências físicas (I)", larg=408 - BD_X, cor_tag="#0d47a1",
          fs=5.1, passo=2.35)
    escala_grafica(p, BX0, 14, ESC, rotulo="Escala gráfica da planta 1:75 (A3)")
    escala_grafica(p, 100, 14, 100, rotulo="Escala gráfica do corte 1:100 (A3)")
    pdf_salvar(pdf, fig, "SM-01")


def corte_transversal(ac, completo=True):
    """Corte transversal esquemático (Y na horizontal, alturas na vertical), olhando para o portão."""
    B, e = G.B, G.PAREDE / 2
    ac.plot([-1200, B + 1200], [0, 0], lw=0.6, c=PRETO)
    ac.add_patch(Rectangle((-1200, -100), B + 2400, 100, fc="#dddddd", ec=PRETO, lw=0.3, hatch="...."))
    for y in (0, B):
        rect(ac, (y - e, y + e, 0, 5420), fc="#9c9c9c", ec=PRETO, lw=0.3)
        rect(ac, (y - 350, y + 350, -350, 150), fc="none", ec=PRETO, lw=0.3, hatch="////")
    ac.plot([-300, B / 2, B + 300], [5420 - 60, 6270, 5420 - 60], lw=0.7, c=PRETO)
    ac.plot([e, B - e], [5380, 5380], lw=0.6, c=CINZA, ls="--")
    ac.text(B / 2, 5440, "forro ≈ 5.380 (medido)", ha="center", va="bottom", fontsize=5.2, color=CONFIRMAR)
    for yc in G.TRILHOS_Y:
        rect(ac, (yc - 260, yc + 260, 0, 4150), fc="#fff3b0", ec="#7a6200", lw=0.5)
    rect(ac, (G.TRILHOS_Y[0] - 260, G.TRILHOS_Y[1] + 260, 4150, 4460), fc="#fff3b0", ec="#7a6200", lw=0.5)
    ac.text(B / 2, 4300, "viga do pórtico W 310 (4.150 a 4.460)", ha="center", va="center", fontsize=5.2)
    rect(ac, (B / 2 - 250, B / 2 + 250, 3650, 4150), fc="white", ec=PRETO, lw=0.4)
    ac.plot([B / 2, B / 2], [3650, 3100], lw=0.4, c=PRETO)
    ac.text(B / 2 + 300, 3500, "talha 2 t", fontsize=5.2)
    if completo:
        rect(ac, (1450, 3650, 0, 2600), fc="#cfe3f5", ec="#2c6ea8", lw=0.6)
        ac.text(2550, 1300, "E07 lavadora\n(envelope 2.600\nde altura*)", ha="center", va="center", fontsize=5.2)
        ac.add_patch(Circle((2600, 3000), 200, fc="#e1bee7", ec="#6a1b9a", lw=0.5))
        ac.text(2850, 3000, "duto de vapor h≈3.000", fontsize=5.2, va="center")
        rect(ac, (5550, 6350, 0, 900), fc="#cfe3f5", ec="#2c6ea8", lw=0.6)
        ac.text(5950, 1050, "E10", ha="center", fontsize=5.2)
        for (y0, y1) in ((e, LY.FAIXA_UTIL_Y[0]), (LY.FAIXA_UTIL_Y[1], B - e)):
            rect(ac, (y0, y1, 0, 120), fc="#fff3b0", ec="#c9a400", lw=0.3, hatch="\\\\\\")
        for lx in (2350, 5500):
            rect(ac, (lx - 250, lx + 250, 4950, 5050), fc="white", ec=PRETO, lw=0.4)
        ac.text(5800, 4850, "luminárias h≈5.000*", fontsize=5.2, va="top", color=CONFIRMAR)
        cadeia = [e, LY.FAIXA_UTIL_Y[0], 3700, 4900, LY.FAIXA_UTIL_Y[1], B - e]
        for a, b in zip(cadeia, cadeia[1:]):
            cota_h(ac, a, b, -650, fs=5.2)
        rot = [("faixa do\npórtico", e, LY.FAIXA_UTIL_Y[0]), ("lado A (sujo /\nlavagem)", LY.FAIXA_UTIL_Y[0], 3700),
               ("corredor", 3700, 4900), ("lado B\n(limpo)", 4900, LY.FAIXA_UTIL_Y[1]),
               ("faixa do\npórtico", LY.FAIXA_UTIL_Y[1], B - e)]
        for t, a, b in rot:
            ac.text((a + b) / 2, -1000, t, ha="center", va="top", fontsize=5.2, linespacing=0.9)
        ac.text(B + 150, -1650, "corte na altura da lavadora, X ≈ 7.000", ha="right", fontsize=5.2)
    else:
        cota_h(ac, 0, B, -700, txt="7.850")
        cota_h(ac, *G.TRILHOS_Y, -400, txt="6.534 (trilhos)")
    for z, t in ((150, "150 topo dos blocos"), (4150, "4.150*"), (4460, "4.460"), (5380, "≈5.380*"), (6270, "≈6.270*")):
        ac.plot([B + 150, B + 650], [z, z], lw=0.25, c=PRETO)
        ac.text(B + 700, z, t, fontsize=5.2, va="center", color=CONFIRMAR if "*" in t else PRETO)


# ---------------------------------------------------------------- SM-02 planta proposta
def f1_planta(pdf):
    fig, p = folha("SM-02", "Desenho 1 — Planta baixa proposta (Alternativa 1, recomendada)", "1:75 (A3)")
    ax, yb = planta(fig)
    base_alt1(ax)
    for nm, rs in LY.A1_CORREDORES:
        r = rs[0]
        ax.text((r[0] + r[1]) / 2 if nm != "Corredor central" else 14500, (r[2] + r[3]) / 2,
                nm.upper(), ha="center", va="center", fontsize=6, color="#555",
                rotation=90 if nm == "Corredor transversal" else 0, zorder=8)
    for y in (800, 7050):
        ax.text(10500, y, "S13 faixa das pernas do pórtico — acesso restrito", ha="center", va="center", fontsize=6,
                color="#8a6d00", zorder=8)
    titulo_vista(p, PL_X, 283.5, "PLANTA BAIXA PROPOSTA — fluxo em U: lado A sujo, lado B limpo, lavagem no fundo   ·   "
                 "ESCALA 1:75")
    yy = legenda_setores(p, COL_X, 280, LY.A1_SETORES, 1)
    yy = legenda_simbolos(p, COL_X, yy, [
        (sim_ret("#fff3b0", "#c9a400", "\\\\\\"), "Faixa das pernas + 500 de folga"),
        (sim_linha("#1565c0", (0, (6, 3)), 0.8), "Alcance do gancho (premissa)"),
        (sim_ret("white", PRETO), "Equipamento (envelope) — tag Exx"),
        (sim_ret("none", PRETO, None, "--"), "Carro extraível / motor no cavalete"),
        (sim_ret("#9c9c9c", PRETO), "Parede existente (150)"),
        (sim_ret("none", PRETO, "////"), "Bloco de fundação (topo +150)"),
    ])
    bloco(p, COL_X, yy, "Notas", [
        f"1. Curso do pórtico limitado por batentes: pernas entre X {br(LY.PERNAS_X[0])} e {br(LY.PERNAS_X[1])} "
        "(folga ≥ 500 do escritório e da parede do portão); trilhos só de X 4.800 a 19.600.",
        "2. Faixa S13: nada fixo e ninguém parado com o pórtico em movimento; zebrado amarelo/preto.",
        "3. Corredor central ≥ 1.200 do portão à lavagem; o transversal (1.100) leva à porta de fundos.",
        "4. Sem teste de motor em funcionamento no galpão: só testes funcionais.",
        "*5. Pré-lavagem externa (lama, palha) fora do galpão, em lavador com separação de óleo.",
    ], larg=COL_L, fs=5.2, passo=2.45)
    yt = tabela_equip(p, BX0, yb - 5, LY.A1_EQUIP, colunas=2, larg=(BX1 - BX0) / 2)
    lista(p, BX0, yt - 3, [(sid, f"{nome}: {just}") for sid, nome, _, _, just in LY.A1_SETORES],
          "Justificativa da posição de cada setor", larg=BX1 - BX0, fs=5.1, passo=2.35)
    quadro_categorias(p, BD_X, yb - 5)
    escala_grafica(p, BX0, 14, ESC)
    p.text(100, 15, "Orientação: portão à direita, como no desenho CRV(MG)20102026 (norte não indicado).", fontsize=5.3)
    pdf_salvar(pdf, fig, "SM-02")


CATEGORIAS = [   # linhas do quadro de áreas pedido no escopo
    ("Recebimento", ["S1"], "Junto ao portão, lado A"),
    ("Desmontagem", ["S2", "S3"], "Motores e transmissões, lado A"),
    ("Lavagem", ["S4", "S5", "S6"], "Área prioritária; S5 com contenção"),
    ("Inspeção e metrologia", ["S7", "S11"], "S7 no lado limpo + sala existente (área interna)"),
    ("Montagem", ["S8", "S9"], "Lado B"),
    ("Testes, qualidade e expedição", ["S10"], "Sem teste com motor funcionando"),
    ("Armazenamento e apoio", ["S12"], "Ferramentas, químicos, quarentena; kits de OS nos carrinhos"),
    ("Faixas do pórtico (restrito)", ["S13"], "Não utilizável para nada fixo"),
]


def linhas_categorias(n_alt):
    q = {i: a for i, _, _, a, _ in V.quadro_areas(n_alt)}
    paredes = [a for i, nm, z, a, _ in V.quadro_areas(n_alt) if nm == "Paredes do escritório"][0]
    circ = [a for i, nm, z, a, _ in V.quadro_areas(n_alt) if z == "circulacao"][0]
    out = []
    for nome, ids, obs in CATEGORIAS:
        a = sum(q.get(i, 0) for i in ids)
        if a:
            out.append((nome, a, obs))
    out.append(("Circulação e acessos", circ, "Corredores, rota de fuga e áreas livres"))
    out.append(("Paredes internas", paredes, "Paredes do escritório existente"))
    return out


def quadro_categorias(p, x, y, n_alt=1, larg=408 - BD_X):
    linhas = [("Setor", "m²", "% útil", "Observações")]
    tot = 0
    for nome, a, obs in linhas_categorias(n_alt):
        linhas.append((nome, br(a, 2), br(a / G.AREA_INTERNA * 100, 1) + "%", obs))
        tot += a
    linhas.append(("Total (área interna)", br(tot, 2), "100%", "Sem sobreposição; confere com 19,85 x 7,70 m"))
    return tabela(p, x, y, [38, 14, 13, larg - 65], linhas, titulo="Quadro de áreas (formato do escopo)", fs=5.2,
                  passo=2.4)


# ---------------------------------------------------------------- SM-03 cotas
def f2_cotas(pdf):
    fig, p = folha("SM-03", "Desenho 2 — Layout cotado e corte transversal (Alternativa 1)", "1:75 planta; 1:100 corte (A3)")
    ax, yb = planta(fig, yl=(-2700, 9700))
    desenhar_setores(ax, LY.A1_SETORES)
    desenhar_galpao(ax, rotulos=False)
    desenhar_portico(ax, posicao=None, cobertura=False)
    desenhar_equip(ax, LY.A1_EQUIP, carro=LY.A1_CARRO_LAVADORA, so_tag=True, fs=5.4)
    cadeia_h(ax, [75, 4600, 5700, 10000, 11000, 13400, 17300, 19925], -1250, ref=LY.FAIXA_UTIL_Y[0])
    cadeia_h(ax, [75, 4075, 4600, 5700, 8100, 10000, 13400, 17300, 19925], 9100, ref=LY.FAIXA_UTIL_Y[1])
    cota_h(ax, LY.PERNAS_X[0], LY.PERNAS_X[1], -2150, txt=f"{br(LY.PERNAS_X[1] - LY.PERNAS_X[0])} curso das pernas*",
           cor=CONFIRMAR)
    cota_h(ax, LY.GANCHO_X[0], LY.GANCHO_X[1], 8400, txt=f"{br(LY.GANCHO_X[1] - LY.GANCHO_X[0])} alcance do gancho*",
           cor=CONFIRMAR)
    cadeia_v(ax, [75, LY.FAIXA_UTIL_Y[0], 3300, 4550, LY.FAIXA_UTIL_Y[1], 7775], 21200, ref=19925)
    cadeia_v(ax, [75, LY.FAIXA_UTIL_Y[0], 3700, 4900, LY.FAIXA_UTIL_Y[1], 7775], -700, ref=75)
    cota_v(ax, 75, 7775, -1450, txt="7.700 (interno)")
    cota_v(ax, 1450, 3650, 5500, txt="2.200")
    cota_h(ax, 5800, 8300, 1100, txt="2.500")
    cota_h(ax, 8300, 9900, 1600, txt="1.600")
    cota_h(ax, 15000, 16600, 1650, txt="1.600")
    cota_v(ax, 2000, 3000, 14850, txt="1.000")
    cota_h(ax, 14800, 16400, 6600, txt="1.600")
    cota_h(ax, 11100, 13100, 1150, txt="2.000")
    cota_h(ax, 10200, 12200, 6600, txt="2.000")
    cota_h(ax, 18000, 19200, 1250, txt="1.200")
    cota_v(ax, 4750, 6150, 19400, txt="1.400")
    cota_h(ax, 4075, 5300, 7300, txt="1.225", fs=5.2)
    cota_h(ax, 19250, 19925, 7300, txt="675", fs=5.2)
    titulo_vista(p, PL_X, 283.5, "LAYOUT COTADO — cotas em mm a partir das faces internas (X 75 / 19.925; Y 75 / 7.775)   ·   "
                 "ESCALA 1:75")
    yn = bloco(p, COL_X, 280, "Notas de cotas e folgas", [
        "Faixa útil entre as pernas do pórtico: Y 1.418 a 6.432 (5.014 mm) = perna (520) + 500 de folga de cada lado.",
        "Lado A: 1.882 (S1–S4) e 2.282 na lavagem (S5). Corredor central 1.250 (1.200 junto à lavagem). "
        "Lado B: 1.882 (S8–S10) e 1.532 em S6–S7.",
        "*Envelopes de equipamento (E07 2.500 x 2.200 x 2.600; cavaletes 1.600 x 1.000) dependem do modelo; "
        "se o escolhido for maior, revisar S5 (folga de 100 atrás da lavadora + corredor de 1.100).",
        "*Curso e alcance do pórtico dependem do truque das pernas (premissa 1.000) e da aproximação do gancho "
        "(premissa 750 do eixo do trilho).",
        "Pernas a 1.225 do escritório e a 675 da parede do portão (mínimo 500).",
        "Instalações aéreas dentro do curso do pórtico: só acima de ~4.600 (viga a 4.460) e abaixo do forro "
        "(~5.380).",
        "Duto da lavadora a ~3.000, entre as faixas das pernas, saindo pela parede de fundos (Desenho 5).",
        "*Altura útil do gancho: depende da talha — verificar com o maior conjunto + cintas + dispositivo.",
        "Blocos de fundação ficam 150 acima do piso junto às paredes: nada encostado nelas entre os pilares "
        "sem conferir o ressalto.",
    ], larg=COL_L, fs=5.2, passo=2.45)
    bloco(p, COL_X, yn - 2, "Premissas que mudam as cotas", [
        "*Truque das pernas do pórtico 1.000 x 520 e aproximação do gancho 750 (confirmar com o fornecedor).",
        "*Lavadora: envelope 2.500 x 2.200 x 2.600 + carro extraível de 1.600 (confirmar com o modelo).",
        "*Cavaletes: 1.600 x 1.000 com o maior motor (confirmar pesos e dimensões dos motores atendidos).",
        "Corredores: central ≥ 1.200; transversal 1.100; rota da porta de fundos 1.250.",
        "Folgas: 500 entre perna em movimento e objeto fixo; 800–900 na frente das bancadas; "
        "1.000 na frente do QDO.",
    ], larg=COL_L, fs=5.2, passo=2.45)
    ac = ax_escala(fig, BX0, 20, (-1300, 9900), (-1800, 6600), 100)
    corte_transversal(ac, completo=True)
    titulo_vista(p, BX0, yb - 6, "CORTE TRANSVERSAL ESQUEMÁTICO (folgas verticais)", 100)
    quadro_cotas(p, 140, yb - 6)
    detalhe_faixa(fig, p, BD_X, yb - 6)
    escala_grafica(p, BX0, 13, ESC, rotulo="Escala gráfica da planta 1:75 (A3)")
    escala_grafica(p, 140, 13, 100, rotulo="Escala gráfica do corte 1:100 (A3)")
    pdf_salvar(pdf, fig, "SM-03")


def detalhe_faixa(fig, p, x, y):
    """Planta 1:50 da faixa entre a parede do lado A e a área útil (mostra de onde vêm os 1.418 mm)."""
    titulo_vista(p, x, y, "DETALHE — faixa do pórtico, lado A (planta)", 50, fs=7.5)
    ax = ax_escala(fig, x, y - 46, (4300, 12200), (-500, 1700), 50)
    rect(ax, (4300, 12200, -75, 75), fc="#9c9c9c", ec=PRETO, lw=0.4, zorder=4)
    for xp in (5000, 10000):
        rect(ax, (xp - 350, xp + 350, -350, 350), fc="none", ec=PRETO, lw=0.4, hatch="////", zorder=3)
        rect(ax, (xp - 75, xp + 75, -90, 90), fc=PRETO, ec=PRETO, zorder=6)
    rect(ax, (4300, 12200, LY.PERNA_FAIXA[0][0], LY.PERNA_FAIXA[0][1]), fc="none", ec="#c9a400", lw=0.5,
         ls=(0, (2, 2)), zorder=3)
    rect(ax, (5300, 12200, 75, LY.FAIXA_UTIL_Y[0]), fc="#fff3b0", ec="#c9a400", lw=0.3, hatch="\\\\\\", alpha=0.6, zorder=2)
    for dy in (-25, 25):
        ax.plot([4800, 12200], [658 + dy, 658 + dy], lw=0.4, c="#5a5a5a", zorder=4)
    rect(ax, (8000, 9000, LY.PERNA_FAIXA[0][0], LY.PERNA_FAIXA[0][1]), fc="#fff3b0", ec="#7a6200", lw=0.7, zorder=7)
    ax.text(8500, 658, "perna (truque*)", ha="center", va="center", fontsize=5.2, zorder=8)
    ax.plot([5300, 5300], [300, 1000], lw=1.4, c="#7a6200", zorder=8)
    ax.text(5250, 1050, "batente*", ha="right", fontsize=5.2, color=CONFIRMAR)
    rect(ax, (5800, 8300, 1450, 1700), fc="white", ec=PRETO, lw=0.5, zorder=6)
    ax.text(7050, 1560, "E07 lavadora", ha="center", va="center", fontsize=5.2, zorder=8)
    rect(ax, (11100, 12200, 1500, 1700), fc="white", ec=PRETO, lw=0.5, zorder=6)
    ax.text(11650, 1580, "E06", ha="center", va="center", fontsize=5.2, zorder=8)
    ax.plot([4300, 12200], [LY.FAIXA_UTIL_Y[0]] * 2, lw=0.8, c="#c9a400", zorder=7)
    xc = 10850
    cadeia_v(ax, [0, 75, 658, LY.PERNA_FAIXA[0][1], LY.FAIXA_UTIL_Y[0]], xc, ref=None, fs=5.2)
    for yy, t in ((37, "½ parede"), (366, "face da parede → eixo do trilho"), (788, "½ perna*"), (1168, "folga de segurança")):
        ax.text(xc + 120, yy, t, fontsize=5.2, va="center", color=CONFIRMAR if "*" in t else PRETO, zorder=9)
    ax.text(4400, LY.FAIXA_UTIL_Y[0] + 40, "limite da faixa útil Y 1.418", fontsize=5.2, va="bottom", zorder=9)
    ax.text(6000, -130, "eixo A (parede do lado da porta)", fontsize=5.2, va="top", zorder=9)
    ax.text(5000, -420, "bloco 700 x 700, topo +150", ha="center", fontsize=5, va="bottom", zorder=9)


def identificacao_visual(p, x, y):
    tabela(p, x, y, [44, 50, 408 - x - 94], [
        ("Categoria", "Onde fica", "Identificação visual"),
        ("Aguardando manutenção", "S1 (E02) e carrinhos E05", "Etiqueta da OS laranja; estado de chegada fotografado"),
        ("Sujas aguardando lavagem", "S4 e zona de carga de S5", "Cestos laranja com etiqueta da OS"),
        ("Limpas aguardando inspeção", "S6 (E10)", "Cestos azuis; protetivo aplicado nas ferrosas"),
        ("Aprovadas aguardando montagem", "Kit da OS (E14) em S8/S9", "Etiqueta verde 'aprovado' + nº da OS; capa"),
        ("Reprovadas / em avaliação", "Quarentena E20 (S12)", "Etiqueta vermelha com motivo; rack vermelho"),
        ("Ferramentas e dispositivos", "E19 (S12) e carros por mecânico", "Painel com silhueta; torquímetros com selo"),
        ("Reparados aguardando retirada", "Expedição E18 (S10)", "Etiqueta de liberação do CQ; capa de proteção"),
        ("Resíduos", "Coletor E03 / abrigo externo", "Recipientes identificados por tipo de resíduo"),
    ], titulo="Identificação visual e localização por categoria", fs=5.1, passo=2.3)


def quadro_cotas(p, x, y):
    linhas = [("Setor", "X de–até", "Y de–até", "comp. x larg.", "m²")]
    for sid, nome, zona, rets, _ in LY.A1_SETORES:
        for i, r in enumerate(rets):
            linhas.append((sid if i == 0 else "", f"{br(r[0])}–{br(r[1])}", f"{br(r[2])}–{br(r[3])}",
                           f"{br(r[1] - r[0])} x {br(r[3] - r[2])}", br(V.area(r), 2)))
    for nm, rs in LY.A1_CORREDORES:
        for r in rs:
            linhas.append((nm.replace("Corredor ", "Corr. ").replace("Rota até a porta de fundos", "Rota fundos"),
                           f"{br(r[0])}–{br(r[1])}", f"{br(r[2])}–{br(r[3])}",
                           f"{br(r[1] - r[0])} x {br(r[3] - r[2])}", br(V.area(r), 2)))
    return tabela(p, x, y, [17, 21, 19, 23, 10], linhas, titulo="Quadro de cotas (mm; origem eixo 1 / eixo A)",
                  fs=5.0, passo=2.25)


# ---------------------------------------------------------------- SM-04 fluxo
def caixa(p, x, y, w, h, txt, fc, ec, fs=5.4):
    p.add_patch(Rectangle((x, y), w, h, fc=fc, ec=ec, lw=0.6))
    p.text(x + w / 2, y + h / 2, txt, ha="center", va="center", fontsize=fs, linespacing=0.95)


def flecha(p, a, b, cor=PRETO, ls="-"):
    p.add_patch(FancyArrowPatch(a, b, arrowstyle="-|>", mutation_scale=7, lw=0.7, color=cor, ls=ls,
                                shrinkA=0, shrinkB=0))


def f3_fluxo(pdf):
    fig, p = folha("SM-04", "Desenho 3 — Fluxo operacional (Alternativa 1)", "1:75 (A3)")
    ax, yb = planta(fig)
    desenhar_setores(ax, LY.A1_SETORES, alfa=0.55)
    desenhar_galpao(ax, rotulos=False)
    desenhar_equip(ax, LY.A1_EQUIP, carro=LY.A1_CARRO_LAVADORA, so_tag=True, fs=5.2)
    for fid, nome, cor, ls, pts in LY.A1_FLUXOS:
        seta(ax, pts, cor, ls=ls, lw=1.5 if fid in ("F1", "F2") else 1.0)
    r = mm_txt(ax, 6.5)
    for n, txt, (x, y) in LY.A1_ETAPAS:
        ax.add_patch(Circle((x, y), r, fc="#263238", ec="white", lw=0.5, zorder=14))
        ax.text(x, y, str(n), ha="center", va="center", fontsize=6, color="white", weight="bold", zorder=15)
    ax.text(20250, 2250, "entrada", fontsize=5.8, ha="left", va="top")
    ax.text(20250, 5600, "saída", fontsize=5.8, ha="left", va="bottom")
    ax.text(-900, 1500, "pessoas", fontsize=5.8, ha="center", va="bottom")
    titulo_vista(p, PL_X, 283.5, "FLUXO DE COMPONENTES, PESSOAS E RESÍDUOS — sem cruzamento entre sujo e limpo   ·   "
                 "ESCALA 1:75")
    yy = legenda_simbolos(p, COL_X, 280, [(sim_linha(c, ls, 1.3), f"{fid}  {nm}") for fid, nm, c, ls, _ in LY.A1_FLUXOS],
                          titulo="Fluxos")
    yy = bloco(p, COL_X, yy, "Etapas (círculos)", [f"{n}. {t}" for n, t, _ in LY.A1_ETAPAS], larg=COL_L, fs=5.3,
               passo=2.6)
    yy = bloco(p, COL_X, yy - 2, "Rotas de fuga", [
        "Corredor central → portão; corredor transversal → porta de fundos (inverter o sentido de abertura "
        "para fora, a confirmar com o projeto de PCI).",
        "Distância máxima até uma saída ≈ 13 m.",
    ], larg=COL_L, fs=5.2, passo=2.45)
    bloco(p, COL_X, yy - 2, "Pontos de contato sujo/limpo e como são tratados", [
        "• Boca da lavadora (carro extraível): carrega pelo lado sujo e descarrega direto para S6, do outro lado "
        "do corredor; limpar o carro e o piso entre ciclos.",
        "• Portão: recebimento (lado A) e expedição (lado B) separados pelo corredor central; não receber e "
        "expedir ao mesmo tempo; conjunto liberado sai com capa.",
        "• Pórtico único para os dois lados: carga limpa não passa sobre posição suja aberta; programar a ordem "
        "das movimentações.",
        "Retorno (retrabalho): peça reprovada vai à quarentena (E20) ou volta à lavadora pelo corredor "
        "transversal; serviço externo (retífica) sai pela expedição e volta pelo recebimento, passando de "
        "novo pela lavagem. Rastreabilidade: etiqueta da OS em cada cesto, carrinho e conjunto.",
    ], larg=COL_L, fs=5.2, passo=2.45)
    fluxograma(p, BX0, yb - 5)
    identificacao_visual(p, BD_X, yb - 5)
    escala_grafica(p, BX0, 13, ESC)
    pdf_salvar(pdf, fig, "SM-04")


def fluxograma(p, x0, ytop):
    p.text(x0, ytop, "FLUXOGRAMA DO PROCESSO", fontsize=8.5, weight="bold")
    suja, lav, limpa, qual, apoio = ("#f6dcc4", "#b8652a"), ("#cfe3f5", "#2c6ea8"), ("#d5ecd6", "#2f7d3c"), \
        ("#e4dcf1", "#6a4c9c"), ("#eeeeee", "#666666")
    w, h, gap = 24.5, 11, 3.3
    yr = ytop - 26          # linha 1
    y2 = yr - 24            # linha 2
    yb = y2 - 22            # ramos
    linha1 = [("Pré-lavagem\nexterna*", apoio), ("1 Recebimento\nOS e fotos", suja), ("2 Drenagem\nde fluidos", suja),
              ("3 Desmontagem\n(S2/S3)", suja), ("Separação e\nidentificação", suja), ("4-5 Lavagem\n(S4/S5)", lav),
              ("6 Escorrimento\ne secagem", lav), ("Inspeção\nvisual", lav)]
    xs = []
    for i, (t, (fc, ec)) in enumerate(linha1):
        x = x0 + i * (w + gap)
        caixa(p, x, yr, w, h, t, fc, ec)
        xs.append(x)
        if i:
            flecha(p, (x - gap, yr + h / 2), (x, yr + h / 2))
    linha2 = [("7 Inspeção\ndimensional", limpa), ("Kit da OS\n(aprovadas + novas)", limpa), ("8 Montagem\n(S8/S9)", limpa),
              ("9 CQ e testes\nfuncionais", qual), ("10 Expedição\n(com capa)", qual)]
    x2 = []
    for i, (t, (fc, ec)) in enumerate(linha2):
        x = xs[-1] - i * (w + gap)
        caixa(p, x, y2, w, h, t, fc, ec)
        x2.append(x)
        if i:
            flecha(p, (x + w + gap, y2 + h / 2), (x + w, y2 + h / 2))
    flecha(p, (xs[-1] + w / 2, yr), (xs[-1] + w / 2, y2 + h))
    caixa(p, x2[0], yb, w, h, "Quarentena (E20)\nreprovada", "#fde0dc", "#b3302d")
    flecha(p, (x2[0] + w / 2 + 4, y2), (x2[0] + w / 2 + 4, yb + h), "#b3302d")
    caixa(p, x2[0] - w - gap, yb, w, h, "Serviço externo\n(retífica)*", apoio[0], apoio[1])
    flecha(p, (x2[0] + 4, y2), (x2[0] - gap - w / 2, yb + h), "#555")
    caixa(p, x2[1] - w - gap, yb, w, h, "Peças novas\n(almoxarifado)", apoio[0], apoio[1])
    flecha(p, (x2[1] - gap - w / 2, yb + h), (x2[1] + 3, y2), "#555")
    caixa(p, x2[3] - w - gap, yb, w, h, "Descarte, garantia\nou recuperação", apoio[0], apoio[1])
    p.plot([x2[0] + w / 2, x2[0] + w / 2], [yb, yb - 4], lw=0.7, c="#b3302d")
    p.plot([x2[0] + w / 2, x2[3] - gap - w / 2], [yb - 4, yb - 4], lw=0.7, c="#b3302d")
    flecha(p, (x2[3] - gap - w / 2, yb - 4), (x2[3] - gap - w / 2, yb), "#b3302d")
    ret = yb - 8
    p.plot([x2[0] - gap - w / 2, x2[0] - gap - w / 2], [yb, ret], lw=0.7, c="#555", ls="--")
    p.plot([x2[0] - gap - w / 2, xs[1] + w / 2], [ret, ret], lw=0.7, c="#555", ls="--")
    p.plot([xs[1] + w / 2, xs[1] + w / 2], [ret, yr - 0.5], lw=0.7, c="#555", ls="--")
    flecha(p, (xs[1] + w / 2, yr - 3), (xs[1] + w / 2, yr), "#555", "--")
    p.text(xs[1] + w / 2 + 2, ret + 1.2, "volta pelo recebimento e passa de novo pela lavagem", fontsize=5.2, color="#555")
    caixa(p, xs[2], yr + h + 4, w + 6, 8, "Resíduos → abrigo externo\n(OLUC, classe I)", "#fff3d6", "#7a5a00", fs=5.1)
    flecha(p, (xs[2] + w / 2, yr + h), (xs[2] + w / 2, yr + h + 4), "#7a5a00")
    caixa(p, xs[5], yr + h + 4, w + 6, 8, "Banho e lodo → destinação\nlicenciada (sem rede pluvial)", "#fff3d6",
          "#7a5a00", fs=5.1)
    flecha(p, (xs[5] + w / 2, yr + h), (xs[5] + w / 2, yr + h + 4), "#7a5a00")
    p.text(100, ret - 7, "*Pré-lavagem externa e retífica: fora do galpão, a confirmar local e fornecedor.",
           fontsize=5.2, color=CONFIRMAR)


# ---------------------------------------------------------------- SM-05 lavagem
ESC_LAV = 40


def f4_lavagem(pdf):
    fig, p = folha("SM-05", "Desenho 4 — Detalhamento da área de lavagem (Alternativa 1)", "1:40 planta; 1:75 corte (A3)")
    xl, yl = (-1800, 10500), (650, 6900)
    h = (yl[1] - yl[0]) / ESC_LAV
    ax = ax_escala(fig, PL_X, PL_TOPO - h, xl, yl, ESC_LAV)
    yb = PL_TOPO - h
    desenhar_setores(ax, LY.A1_SETORES, fs=8)
    rect(ax, (5700, 10000, 1418, 3700), fc="none", ec="#0d47a1", lw=1.4, zorder=8)
    ax.text(5750, 3740, "contenção perimetral: lombada transponível h = 50 mm (ou piso rebaixado 50 mm)*",
            fontsize=6.3, color="#0d47a1", va="bottom", zorder=9)
    desenhar_galpao(ax, rotulos=False)
    desenhar_portico(ax, posicao=None, cobertura=False, fs=6.3)
    desenhar_equip(ax, LY.A1_EQUIP, carro=LY.A1_CARRO_LAVADORA, fs=6.8)
    ax.add_patch(Arc((8300, 3650), 2200, 2200, theta1=270, theta2=360, lw=0.4, ls="--", ec=PRETO, zorder=7))
    ax.text(9450, 3500, "porta*", fontsize=6.3, zorder=9, color=CONFIRMAR)
    for y in (2100, 3000):
        ax.plot([8300, 9900], [y, y], lw=0.5, c=PRETO, zorder=7)
    ax.text(9100, 2400, "trilhos do carro", ha="center", fontsize=6, zorder=9)
    rect(ax, LY.CANALETA, fc="#90a4ae", ec=PRETO, lw=0.6, hatch="||||", zorder=8)
    ax.text(9950, 1330, "CN-1 canaleta com grelha, larg. 150*", ha="right", va="top", fontsize=6.3, zorder=9,
            color=CONFIRMAR)
    for y in (1750, 3350):
        seta(ax, [(8500, y), (9750, y)], "#0d47a1", lw=0.7)
    ax.text(8550, 1500, "caimento ≥ 1% para CN-1*", fontsize=6.3, color="#0d47a1", zorder=12, va="bottom")
    seta(ax, LY.DRENO, "#6d4c41", ls=(0, (5, 2)), lw=1.0)
    ax.add_patch(Circle(LY.CAIXA_INSPECAO, 140, fc="white", ec="#6d4c41", lw=0.8, zorder=12))
    ax.text(LY.CAIXA_INSPECAO[0], LY.CAIXA_INSPECAO[1] + 200, "CI caixa de inspeção", ha="center", fontsize=6.3, zorder=12)
    rect(ax, LY.CAIXA_CONTENCAO, fc="#d7ccc8", ec="#6d4c41", lw=0.8, zorder=12)
    cx = (LY.CAIXA_CONTENCAO[0] + LY.CAIXA_CONTENCAO[1]) / 2
    ax.text(cx, 2700, "CX-1\ncaixa de\ncontenção\nestanque\n(externa)*", ha="center", va="center", fontsize=6,
            zorder=13, color=CONFIRMAR)
    ax.text(2900, 1350, "dreno enterrado, Ø a dimensionar*", fontsize=6.3, color="#6d4c41", zorder=12, va="top")
    ax.plot(*zip(*LY.DUTO_EXAUSTAO), lw=3.5, c="#ce93d8", alpha=0.85, zorder=11)
    ax.text(700, 2640, "EX-1 duto de vapor h ≈ 3.000 → parede de fundos*", fontsize=6.3, color="#6a1b9a", zorder=12,
            va="bottom")
    ax.plot(*zip(*LY.AGUA), lw=0.9, c="#1976d2", zorder=11)
    ax.text(-1300, 3560, "AF água fria", fontsize=6.3, color="#1976d2", zorder=12, va="bottom")
    ax.text(7000, 4300, "CORREDOR — limite sujo/limpo: faixa amarela; limpar o carro antes de descarregar",
            ha="center", va="center", fontsize=6.3, zorder=9)
    ax.text(5850, 4980, "lado limpo", fontsize=7, color="#2f7d3c", weight="bold", zorder=9, va="bottom")
    ax.text(10030, 2950, "lado sujo", fontsize=6.3, color="#b8652a", weight="bold", zorder=9)
    cota_h(ax, 5700, 10000, 950, txt="4.300 área contida (S5)", fs=6)
    cota_v(ax, 1418, 3700, 10400, txt="2.282", fs=6)
    cota_h(ax, 5800, 8300, 4000, txt="2.500*", cor=CONFIRMAR, fs=6)
    cota_v(ax, 1450, 3650, 5500, txt="2.200*", cor=CONFIRMAR, fs=6)
    cota_h(ax, 8300, 9900, 1150, txt="1.600", fs=6)
    cota_h(ax, 4600, 5700, 6650, txt="1.100", fs=6)
    titulo_vista(p, PL_X, 283.5, "PLANTA AMPLIADA DA LAVAGEM — S4, S5, S6 e apoio S12   ·   ESCALA 1:40")
    yy = bloco(p, 336, 280, "Notas da lavagem", [
        "Sequência (sujo → limpo): carrinho sujo (E05) → peças pequenas (E08) ou carro da lavadora (E07) → "
        "ciclo fechado → descarga para S6 (E10) → inspeção visual → estabilização térmica → S7 (E12).",
        "Efluentes e resíduos:",
        "• Banho esgotado: bombeado para IBC/tambor e enviado a destinador licenciado; emulsão não vai para "
        "caixa separadora.",
        "• Óleo do separador da lavadora → OLUC (coletor autorizado). Lodo e filtros → resíduo classe I.",
        "• CN-1 e lava-olhos → CX-1 estanque, esvaziada por caminhão de sucção licenciado",
        "*  ou ligada à estação de tratamento da usina, se licenciada para esse efluente.",
        "• Nenhuma ligação com rede pluvial, sanitária ou solo.",
        "*Piso de S5: revestimento resistente a álcali e calor (ex.: uretano-cimento), juntas seladas, rodapé "
        "sanitário.",
        "Segurança: lava-olhos/chuveiro (E22) a ≤ 10 s; intertravamento da porta; LOTO na manutenção; luvas, "
        "avental e protetor facial na descarga (peças e solução quentes).",
    ], larg=408 - 336, fs=5.2, passo=2.45)
    yy = legenda_simbolos(p, 336, yy - 1, [
        (sim_ret("none", "#0d47a1"), "Limite da contenção (S5)"),
        (sim_ret("#90a4ae", PRETO, "||||"), "Canaleta com grelha CN-1"),
        (sim_linha("#6d4c41", (0, (5, 2)), 0.9), "Dreno enterrado → CX-1"),
        (sim_linha("#ce93d8", "-", 3), "Duto de exaustão EX-1"),
        (sim_linha("#1976d2", "-", 0.8), "Água fria AF"),
    ])
    bloco(p, 336, yy - 1, "Dados para selecionar a lavadora", [
        "*1. Maior peça (blocos, carcaças): dimensões e peso → diâmetro, altura e carga da mesa giratória.",
        "*2. Química e temperatura aprovadas pelos fabricantes dos componentes (ferro, aço, alumínio).",
        "*3. Utilidades: tensão, potência de aquecimento, água de reposição, condensador de vapor ou exaustão.",
        "4. Separador de óleo (skimmer), filtro do banho e drenagem do lodo pela frente da máquina.",
        "5. Carro extraível para carga por pórtico; porta intertravada (NR-12); manual em português.",
        "6. Consumo de detergente e frequência de troca do banho → volume de resíduo por mês.",
    ], larg=408 - 336, fs=5.2, passo=2.45)
    ac = ax_escala(fig, BX0, 34, (-1800, 10600), (-1500, 3600), 75)
    corte_lavagem(ac)
    titulo_vista(p, BX0, yb - 6, "CORTE ESQUEMÁTICO A-A (ao longo do lado A, Y ≈ 2.500)", 75)
    detalhe_contencao(p, 192, yb - 6)
    tabela(p, BD_X, yb - 6, [30, 62, 408 - BD_X - 92], [
        ("Componente", "Processo indicado", "Cuidados"),
        ("Blocos, cabeçotes, cárteres de ferro", "Aspersão aquosa alcalina quente (E07)", "Proteger roscas e galerias; secar e proteger"),
        ("Carcaças e tampas de alumínio", "E07 com detergente compatível com alumínio*", "pH e inibidor conforme fabricante*"),
        ("Engrenagens, eixos, rolamentos", "E07 em cesto ou E08; rolamento a reaproveitar à parte", "Proteção anticorrosiva logo após secar"),
        ("Peças pequenas e fixadores", "E08 (tanque com bomba, solução aquosa)", "Cestos por OS; nada de diesel/solvente"),
        ("Vedações, sensores, chicotes, injeção", "Não lavar na lavadora", "Limpeza manual ou serviço especializado"),
    ], titulo="Processo por tipo de componente", fs=5.1, passo=2.35)
    escala_grafica(p, BX0, 13, ESC_LAV, metros=4, rotulo="Escala gráfica da planta 1:40 (A3)")
    escala_grafica(p, 130, 13, 75, rotulo="Escala gráfica do corte 1:75 (A3)")
    pdf_salvar(pdf, fig, "SM-05")


def corte_lavagem(ac):
    ac.plot([-1800, 10600], [0, 0], lw=0.3, c=CINZA)
    ac.add_patch(Polygon([(5700, 0), (9850, -40), (9850, -100), (5700, -100)], fc="#cfd8dc", ec=PRETO, lw=0.4))
    ac.add_patch(Polygon([(75, 0), (5650, 0), (5650, -100), (75, -100)], fc="#eceff1", ec=PRETO, lw=0.3))
    ac.add_patch(Polygon([(10000, 0), (10600, 0), (10600, -100), (10000, -100)], fc="#eceff1", ec=PRETO, lw=0.3))
    ac.add_patch(Rectangle((5650, 0), 50, 50, fc="#0d47a1", ec="#0d47a1"))
    ac.add_patch(Rectangle((9850, -300), 150, 300, fc="#90a4ae", ec=PRETO, lw=0.4))
    rect(ac, (-75, 75, -100, 3500), fc="#9c9c9c", ec=PRETO, lw=0.3)
    rect(ac, (5800, 8300, 0, 2600), fc="#cfe3f5", ec="#2c6ea8", lw=0.6)
    ac.text(7050, 1300, "E07 lavadora (envelope*)", ha="center", va="center", fontsize=5.6)
    rect(ac, (8300, 9900, 120, 900), fc="none", ec=PRETO, lw=0.4, ls="--")
    ac.text(9100, 1000, "carro extraível", ha="center", fontsize=5.4)
    ac.plot([9925, 9925, 5100, -1050, -1050], [-300, -450, -650, -800, -850], lw=0.8, c="#6d4c41", ls=(0, (5, 2)))
    rect(ac, (-1650, -450, -1400, 50), fc="#d7ccc8", ec="#6d4c41", lw=0.6)
    ac.text(-1050, -600, "CX-1*", ha="center", fontsize=5.6)
    ac.plot([5900, 5900, -700], [2600, 3000, 3000], lw=2.8, c="#ce93d8")
    ac.text(2500, 3080, "EX-1 h ≈ 3.000", fontsize=5.6, va="bottom")
    ac.text(7700, -330, "caimento ≥ 1% →", fontsize=5.6, va="top", color="#0d47a1")
    ac.text(5600, 100, "lombada 50", fontsize=5.4, ha="right", color="#0d47a1")
    ac.text(2000, -1000, "dreno enterrado com caixa de inspeção (CI)", fontsize=5.6, color="#6d4c41")
    ac.text(120, 3300, "parede de fundos", ha="left", fontsize=5.4)


def detalhe_contencao(p, x, y):
    """Detalhe esquemático (sem escala) da borda da contenção e da canaleta."""
    p.text(x, y, "DETALHE — contenção e canaleta (sem escala)", fontsize=7, weight="bold")
    y0 = y - 40
    p.add_patch(Polygon([(x, y0), (x + 30, y0 - 1.2), (x + 30, y0 - 9), (x, y0 - 9)], fc="#cfd8dc", ec=PRETO, lw=0.4))
    p.add_patch(Rectangle((x + 30, y0 - 14), 8, 14 - 1.2, fc="#90a4ae", ec=PRETO, lw=0.4, hatch="||"))
    p.add_patch(Rectangle((x + 38, y0 - 9), 12, 9, fc="#eceff1", ec=PRETO, lw=0.4))
    p.add_patch(Polygon([(x - 6, y0), (x - 2, y0 + 4), (x + 2, y0 + 4), (x + 4, y0)], fc="#0d47a1", ec="#0d47a1"))
    p.plot([x, x + 30], [y0 + 0.6, y0 - 0.6], lw=0.9, c="#ff8f00")
    p.plot([x + 30, x + 38], [y0 - 1.2, y0 - 1.2], lw=1.2, c=PRETO, ls=(0, (1, 0.6)))
    t = [("lombada transponível 50 mm (ou degrau do rebaixo)", x - 4, y0 + 8),
         ("revestimento uretano-cimento, rodapé sanitário*", x + 2, y0 + 3.5),
         ("laje 100 mm com caimento ≥ 1%*", x + 2, y0 - 12),
         ("canaleta com grelha removível,", x + 30, y0 - 17),
         ("fundo inclinado para o dreno*", x + 30, y0 - 20),
         ("juntas seladas (selante resistente a álcali)*", x + 2, y0 - 24)]
    for s, xx, yy in t:
        p.text(xx, yy, s, fontsize=5.2, color=CONFIRMAR if s.endswith("*") else PRETO)


# ---------------------------------------------------------------- SM-06 infraestrutura
def f5_infra(pdf):
    fig, p = folha("SM-06", "Desenho 5 — Infraestrutura preliminar (Alternativa 1)", "1:75 (A3)")
    ax, yb = planta(fig)
    desenhar_setores(ax, LY.A1_SETORES, alfa=0.35, rotulo=False)
    desenhar_galpao(ax, rotulos=False)
    desenhar_equip(ax, LY.A1_EQUIP, carro=LY.A1_CARRO_LAVADORA, so_tag=True, fs=5.0)
    for nm, pts in LY.REDE_ELETRICA:
        ax.plot(*zip(*pts), lw=0.9, c="#d32f2f", ls=(0, (6, 2)) if "enterrado" in nm or not nm else (0, (1, 1.5)), zorder=11)
    for nm, pts in LY.REDE_AR:
        ax.plot(*zip(*pts), lw=0.9, c="#2e7d32", ls=(0, (3, 1.5)), zorder=11)
    ax.plot(*zip(*LY.DUTO_EXAUSTAO), lw=3.2, c="#ce93d8", zorder=11)
    seta(ax, LY.DRENO, "#6d4c41", ls=(0, (5, 2)), lw=1.0)
    ax.plot(*zip(*LY.AGUA), lw=0.9, c="#1976d2", zorder=11)
    rect(ax, LY.CANALETA, fc="#90a4ae", ec=PRETO, lw=0.5, hatch="||||", zorder=12)
    rect(ax, LY.CAIXA_CONTENCAO, fc="#d7ccc8", ec="#6d4c41", lw=0.6, zorder=12)
    ax.text(sum(LY.CAIXA_CONTENCAO[:2]) / 2, 2700, "CX-1", ha="center", va="center", fontsize=5.6, zorder=13)
    rect(ax, LY.COMPRESSOR, fc="#c8e6c9", ec="#2e7d32", lw=0.6, zorder=12)
    ax.text(sum(LY.COMPRESSOR[:2]) / 2, sum(LY.COMPRESSOR[2:]) / 2, "compres-\nsor*", ha="center", va="center",
            fontsize=5.2, zorder=13, color=CONFIRMAR)
    for cp, x, y in LY.CAIXAS_PISO:
        rect(ax, (x - 150, x + 150, y - 150, y + 150), fc="white", ec="#d32f2f", lw=0.7, zorder=13)
        ax.plot([x - 150, x + 150], [y - 150, y + 150], lw=0.4, c="#d32f2f", zorder=13)
        ax.plot([x - 150, x + 150], [y + 150, y - 150], lw=0.4, c="#d32f2f", zorder=13)
        ax.text(x + 200, y, cp, ha="left", va="center", fontsize=5.2, color="#d32f2f", zorder=13)
    for rr, x, y in LY.ENROLADORES:
        ax.add_patch(Circle((x, y), 170, fc="#c8e6c9", ec="#2e7d32", lw=0.7, zorder=13))
        ax.text(x, y, "R", ha="center", va="center", fontsize=5, zorder=14)
        ax.text(x - 220 if x > 19000 else x + 220, y, rr, ha="right" if x > 19000 else "left", va="center",
                fontsize=5.2, color="#2e7d32", zorder=14)
    for x, y in LY.LUMINARIAS:
        ax.add_patch(Circle((x, y), 220, fc="#fff9c4", ec="#795548", lw=0.5, zorder=12))
        ax.plot([x - 220, x + 220], [y, y], lw=0.4, c="#795548", zorder=12)
        ax.plot([x, x], [y - 220, y + 220], lw=0.4, c="#795548", zorder=12)
    for t, x, y in LY.EXTINTORES:
        ax.add_patch(Polygon([(x - 160, y - 140), (x + 160, y - 140), (x, y + 180)], fc="#e53935", ec="#b71c1c",
                             lw=0.5, zorder=13))
        ax.text(x + 220, y, t, fontsize=5, va="center", color="#b71c1c", zorder=13)
    for x, y in ((300, 1950), (19650, 3925), (5150, 6250), (4300, 4100)):
        rect(ax, (x - 140, x + 140, y - 80, y + 80), fc="#a5d6a7", ec="#1b5e20", lw=0.5, zorder=13)
    titulo_vista(p, PL_X, 283.5, "INFRAESTRUTURA PRELIMINAR — traçados indicativos, sujeitos aos projetos complementares   ·   "
                 "ESCALA 1:75")
    yy = legenda_simbolos(p, COL_X, 280, [
        (sim_linha("#d32f2f", (0, (6, 2))), "Eletroduto enterrado (BT)"),
        (sim_linha("#d32f2f", (0, (1, 1.5))), "Circuito aéreo fora do curso"),
        (lambda q, x, y: q.add_patch(Rectangle((x + 2, y - 1.3), 2.6, 2.6, fc="white", ec="#d32f2f", lw=0.6)),
         "CP caixa de piso 220/380 V"),
        (sim_linha("#2e7d32", (0, (3, 1.5))), "Rede de ar comprimido"),
        (lambda q, x, y: q.add_patch(Circle((x + 3.3, y), 1.4, fc="#c8e6c9", ec="#2e7d32", lw=0.6)),
         "R enrolador de ar (15 m)"),
        (sim_linha("#1976d2", "-", 0.9), "AF água fria"),
        (sim_linha("#6d4c41", (0, (5, 2)), 0.9), "Dreno industrial → CX-1"),
        (sim_linha("#ce93d8", "-", 3), "EX-1 exaustão da lavadora"),
        (lambda q, x, y: q.add_patch(Circle((x + 3.3, y), 1.6, fc="#fff9c4", ec="#795548", lw=0.5)),
         "Luminária LED h ≈ 4.800–5.200*"),
        (lambda q, x, y: q.add_patch(Rectangle((x + 2, y - 0.8), 3, 1.6, fc="#a5d6a7", ec="#1b5e20", lw=0.5)),
         "Iluminação de emergência / saída"),
        (lambda q, x, y: q.add_patch(Polygon([(x + 2, y - 1.3), (x + 5, y - 1.3), (x + 3.5, y + 1.5)],
                                              fc="#e53935", ec="#b71c1c", lw=0.4)), "Extintor (posição indicativa)*"),
    ], titulo="Convenções")
    yy = bloco(p, COL_X, yy, "Por que este traçado", [
        "• Dentro do curso do pórtico nada desce do teto abaixo de ~4,6 m e nada cruza a faixa das pernas: "
        "tomadas em caixas de piso (CP) alimentadas por eletrodutos sob o corredor.",
        "• Ar por enroladores nas pontas (R1–R4): 15 m cobrem a faixa útil sem mangueira fixa.",
        "• Lavadora: energia, água e duto de vapor por cima do S12 (fora do curso), saindo pelos fundos.",
        "*• Embutidos no piso (eletrodutos, CP, dreno, CN-1) precisam ser definidos ANTES da concretagem.",
    ], larg=COL_L, fs=5.2, passo=2.45)
    yy = bloco(p, COL_X, yy - 2, "Embutidos a definir antes de concretar", [
        "□ Trilhos do pórtico (X 4.800–19.600), ancoragem e nivelamento",
        "□ Eletrodutos sob o corredor e caixas CP-01 a CP-07",
        "□ Contenção e caimento de S5, CN-1, dreno, CI e CX-1",
        "□ Juntas seladas e revestimento da área de lavagem",
        "□ Base nivelada da lavadora e fixação da prensa",
        "□ Malha de aterramento e descidas",
    ], larg=COL_L, fs=5.2, passo=2.5)
    yt = tabela(p, BX0, yb - 5, [24, 50, 147], [
        ("Ponto", "Atende", "Utilidade e observação"),
        ("QDO (E23)", "Toda a oficina", "Parede de fundos, fora da área úmida, 1.000 livres na frente; DR nas tomadas; aterramento único*"),
        ("C1", "Lavadora E07", "Trifásico dedicado por cima do S12; potência e tensão conforme o modelo escolhido*"),
        ("C2", "Pórtico/talha E01", "Trifásico; festão ou enrolador de cabo ao longo do trilho B (fornecedor)*"),
        ("C3", "CP-01 a CP-07", "Tomadas industriais 220/380 V em caixas de piso de ferro fundido, eletroduto sob o corredor"),
        ("C4 / C5", "Iluminação", "LED industrial acima do pórtico e abaixo do forro; emergência nas saídas e corredores"),
        ("C6", "Escritório S11", "Tomadas, ar-condicionado (metrologia) e iluminação de bancada"),
        ("AF", "E07, E22", "Entrada pelos fundos; lava-olhos/chuveiro e enchimento da lavadora"),
        ("Dreno", "CN-1, E22", "Para CX-1 estanque; nenhuma ligação com rede pluvial ou sanitária"),
        ("Ar", "R1–R4, E10, E06/E15", "Compressor em abrigo externo; rede interna pelos fundos e externa pela fachada A*"),
        ("EX-1", "E07", "Duto de vapor a ~3,0 m até a parede de fundos; avaliar condensador de vapor na lavadora*"),
    ], titulo="Pontos de utilidade (potências e vazões a levantar)", fs=5.6, passo=2.65)
    bloco(p, BX0, yt - 3, "Ventilação, exaustão, iluminação e segurança elétrica", [
        "Ventilação natural pelos vãos superiores das laterais (D12: confirmar fechamento); exaustão localizada "
        "na lavadora (EX-1); anteparo de sopro com coletor de névoa em E10; escritório/metrologia fechado com "
        "ar-condicionado.",
        "*Iluminação: geral LED acima do pórtico e abaixo do forro; reforço localizado nas bancadas E06, E12, E15 e "
        "no S11; níveis a definir em cálculo luminotécnico (referência NBR ISO/CIE 8995-1).",
        "Segurança elétrica: DR nas tomadas, grau de proteção adequado na área úmida (sem caixa de piso em S5), "
        "pontos de bloqueio (LOTO) para lavadora, pórtico e prensa, QDO sinalizado e com acesso restrito (NR-10).",
    ], larg=BX1 - BX0, fs=5.6, passo=2.65)
    yp = bloco(p, BD_X, yb - 5, "Pendências de infraestrutura", [
        "*Tensão, potência disponível e ponto de entrega da rede da usina.",
        "*Demanda de ar (ferramentas, sopro, lavadora) e local do compressor ou rede existente.",
        "*Destino do efluente da CX-1: caminhão de sucção licenciado ou estação de tratamento da usina.",
        "*Fechamento dos vãos superiores das laterais (D12) → ventilação natural e poeira.",
        "Projetos: elétrico NBR 5410 / NR-10, SPDA NBR 5419, iluminação NBR ISO/CIE 8995-1, PCI (CBMMG), "
        "hidráulica/drenagem e exaustão — profissionais habilitados, com ART.",
    ], larg=408 - BD_X, fs=5.4, passo=2.55)
    bloco(p, BD_X, yp - 2, "Sequência sugerida para as instalações", [
        "1. Fechar modelo e cargas dos equipamentos (lavadora, pórtico, compressor).",
        "2. Projetos elétrico, hidráulico/drenagem, exaustão e PCI.",
        "3. Embutidos no piso antes da concretagem (lista na coluna da direita).",
        "4. Luminárias e suportes acima do pórtico; instalações aparentes.",
        "5. Comissionamento: estanqueidade de S5 e CX-1, iluminância, DR, aterramento.",
    ], larg=408 - BD_X, fs=5.4, passo=2.55)
    escala_grafica(p, BX0, 13, ESC)
    pdf_salvar(pdf, fig, "SM-06")


# ---------------------------------------------------------------- SM-07 alternativa 2
def f6_alt2(pdf):
    fig, p = folha("SM-07", "Alternativa 2 — layout otimizado com ponte rolante (comparação)", "1:75 (A3)")
    ax, yb = planta(fig)
    desenhar_setores(ax, LY.A2_SETORES)
    desenhar_galpao(ax, trilhos=None)
    rotulos_galpao(ax)
    desenhar_equip(ax, LY.A2_EQUIP, carro=LY.A2_CARRO_LAVADORA)
    gx, gy = LY.PONTE_GANCHO_X, LY.PONTE_GANCHO_Y
    rect(ax, (gx[0], gx[1], gy[0], gy[1]), fc="none", ec="#1565c0", lw=0.6, ls=(0, (6, 3)), zorder=9)
    ax.text(10500, 3750, "CORREDOR CENTRAL 1.600", ha="center", va="center", fontsize=6, color="#555", zorder=9)
    ax.text(10500, 3350, "alcance do gancho da ponte rolante (premissa): X 700–19.300, Y 900–6.950", ha="center",
            va="center", fontsize=5.8, color="#1565c0", zorder=9)
    for y in (300, 7550):
        ax.plot([0, G.L], [y, y], lw=0.9, c="#7a6200", ls=(0, (10, 3)), zorder=9)
    ax.text(10000, 8050, "vigas de rolamento sobre pilares reforçados ou independentes*", ha="center", va="bottom",
            fontsize=6, color=CONFIRMAR, zorder=9)
    ax.text(2075, 5350, "S11\nESCRITÓRIO\nTÉCNICO E\nMETROLOGIA", ha="center", va="center", fontsize=6.3, zorder=8,
            color="#8a7d4f", weight="bold")
    titulo_vista(p, PL_X, 283.5, "ALTERNATIVA 2 — sem trilhos no piso; bancadas e utilidades junto às paredes   ·   ESCALA 1:75")
    yy = legenda_setores(p, COL_X, 280, LY.A2_SETORES, 2)
    yy = bloco(p, COL_X, yy, "Recomendação", [
        "Implantar a Alternativa 1: usa o pórtico já previsto e não depende de reforço estrutural.",
        "A Alternativa 2 só se justifica se o galpão ainda não foi executado e o estudo estrutural confirmar a "
        "viga de rolamento a custo aceitável: devolve a largura perdida para as pernas do pórtico "
        "(+74% de área produtiva) e permite uma 2ª posição de montagem.",
        "*Em qualquer caso, decidir antes de concretar o piso: trilhos, caixas de piso, dreno e canaleta "
        "são embutidos.",
    ], larg=COL_L, fs=5.2, passo=2.45)
    bloco(p, COL_X, yy - 2, "Se a Alternativa 2 for estudada", [
        "*• Reações das rodas da ponte e esforços horizontais nos pilares atuais (hoje dimensionados para a "
        "cobertura) ou pilares próprios da via.",
        "*• Fundações: hoje um bloco 600 x 600 sobre uma estaca Ø 250 x 4 m por pilar.",
        "*• Altura: viga de rolamento + ponte + talha abaixo do forro (~5,38 m) com altura útil de gancho suficiente.",
        "• Batentes, limitadores, alimentação por barramento e inspeção (NR-11/NR-12).",
    ], larg=COL_L, fs=5.2, passo=2.45)
    y2 = tabela(p, BX0, yb - 5, [56, 82, 83], comparativo(), titulo="Comparação das alternativas de layout", fs=5.2,
                passo=2.4)
    tabela(p, BX0, y2 - 3, [27, 49, 49, 48, 48], matriz_icamento(), titulo="Matriz — equipamento de içamento",
           fs=5.0, passo=2.25)
    bloco(p, BD_X, yb - 5, "Critérios de escolha do içamento", [
        "1. Peso do maior conjunto + dispositivo + cintas ≤ capacidade nominal (com o fator do fabricante).",
        "2. Área a cobrir: todas as posições que exigem içamento devem estar no alcance do gancho.",
        "3. Altura útil do gancho com o maior conjunto suspenso e folga sobre bancadas e cavaletes.",
        "4. Capacidade da estrutura (galpão ou estrutura própria) e das fundações.",
        "5. Interferência com piso, iluminação, dutos e circulação.",
        "*Dimensionamento, especificação e ART: profissional habilitado (NR-11, NR-12, ABNT NBR 8400).",
    ], larg=408 - BD_X, fs=5.2, passo=2.45)
    escala_grafica(p, BD_X, 67, ESC)
    pdf_salvar(pdf, fig, "SM-07")


def comparativo():
    q1 = {i: a for i, _, _, a, _ in V.quadro_areas(1)}
    q2 = {i: a for i, _, _, a, _ in V.quadro_areas(2)}
    prod1 = sum(q1[f"S{i}"] for i in range(1, 11))
    prod2 = sum(q2[f"S{i}"] for i in range(1, 11))
    return [
        ("Critério", "Alt. 1 — pórtico do projeto (recomendada)", "Alt. 2 — ponte rolante"),
        ("Área dos setores produtivos S1–S10", f"{br(prod1, 1)} m²", f"{br(prod2, 1)} m² (+{br((prod2 / prod1 - 1) * 100)}%)"),
        ("Área perdida para o içamento", f"{br(q1['S13'], 1)} m² (faixas S13)", "≈ 0 (só pilares da viga)"),
        ("Posições de motor (desmontagem + montagem)", "1 + 1", "1 + 2"),
        ("Corredor central", "1,20–1,25 m", "1,60 m (paleteira)"),
        ("Utilidades nas estações", "caixas de piso + enroladores nas pontas", "pelas paredes e pilares"),
        ("Dependência estrutural", "nenhuma nova (pórtico autoportante)", "viga de rolamento, pilares e fundações*"),
        ("Situação no projeto atual", "já previsto (desenho CRV)", "exige revisão do projeto estrutural"),
        ("Principais riscos", "esmagamento nas pernas; trilho no piso", "carga suspensa sobre todo o vão"),
        ("Custo relativo (sem cotação)", "menor", "maior (estrutura + ponte)"),
    ]


def matriz_icamento():
    return [
        ("Critério", "Ponte rolante univiga 2 t", "Pórtico móvel 2 t sobre trilhos (projeto)", "Monovia central + talha",
         "Pórtico com rodízios + guindaste de oficina"),
        ("Cobertura", "Quase todo o piso", "Faixa entre as pernas, no curso dos batentes", "Só a linha central",
         "Pontual; reposicionado à mão"),
        ("Ocupação de piso", "Nenhuma", "Trilhos + faixas das pernas (≈ 37,5 m²)", "Nenhuma", "Pernas e base no ponto de uso"),
        ("Exigência estrutural", "Viga de rolamento e pilares/fundações*", "Autoportante; trilhos ancorados no piso",
         "Viga apoiada na estrutura (verificar)*", "Nenhuma"),
        ("Flexibilidade", "Alta", "Média", "Baixa (layout em linha)", "Baixa para conjuntos pesados"),
        ("Riscos principais", "Carga sobre pessoas", "Esmagamento nas pernas", "Carga sobre o corredor",
         "Tombamento; esforço manual"),
        ("Indicado quando", "Galpão ainda em projeto e estrutura viável", "Projeto atual mantido",
         "Fluxo linear e poucos pontos", "Apoio pontual (ex.: carga da E08)"),
    ]


def main():
    with PdfPages(PDF) as pdf:
        for f in (f0_existente, f1_planta, f2_cotas, f3_fluxo, f4_lavagem, f5_infra, f6_alt2):
            f(pdf)
        d = pdf.infodict()
        d["Title"] = "Oficina de motores e transmissões — pranchas do estudo de layout"
        d["Subject"] = "Estudo conceitual sobre o desenho CRV(MG)20102026 rev. 00"
    print(PDF)


if __name__ == "__main__":
    main()
