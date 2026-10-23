"""Planta em CAD (DXF R2010, mm, escala 1:1) com camadas separadas.

python exportar_dxf.py  ->  Planta_Sala_Motores_OUT2026.dxf
Três quadros no espaço do modelo, todos com a origem no cruzamento dos eixos 1 e A:
  Alternativa 1 (recomendada) em Y 0; existente deslocado +16.000 em Y; Alternativa 2 deslocada -16.000.
Abre no AutoCAD, BricsCAD, DraftSight, LibreCAD e QCAD; no AutoCAD, "Salvar como" DWG se preferir.
"""
import json
from pathlib import Path

import ezdxf
from ezdxf import recover
from ezdxf import units
from ezdxf.enums import TextEntityAlignment

import galpao as G
import layout as LY
import verificacao as V

AQUI = Path(__file__).parent
DXF = AQUI / "Planta_Sala_Motores_OUT2026.dxf"
DY_EXISTENTE, DY_ALT2 = 16000, -16000
H_TXT = 150            # altura de texto no modelo (≈2 mm impresso em 1:75)

CAMADAS = {   # nome: (cor ACI, tipo de linha)
    "00-EIXOS": (8, "DASHDOT"), "00-PAREDES": (8, "CONTINUOUS"), "00-PILARES": (7, "CONTINUOUS"),
    "00-BLOCOS": (8, "CONTINUOUS"), "00-PORTAS": (7, "CONTINUOUS"), "00-TRILHOS": (9, "CONTINUOUS"),
    "00-ORIGINAL-PDF": (151, "CONTINUOUS"), "00-AREAS-DEMARCADAS": (8, "DASHED"), "00-A-CONFIRMAR": (6, "CONTINUOUS"),
    "SETORES": (3, "CONTINUOUS"), "SETORES-TEXTO": (3, "CONTINUOUS"), "EQUIPAMENTOS": (7, "CONTINUOUS"),
    "EQUIPAMENTOS-TEXTO": (7, "CONTINUOUS"), "PORTICO-FAIXA": (2, "CONTINUOUS"), "PORTICO-GANCHO": (5, "DASHED"),
    "CORREDORES": (8, "CONTINUOUS"), "COTAS": (1, "CONTINUOUS"), "FLUXOS": (30, "CONTINUOUS"),
    "INFRA-ELETRICA": (1, "DASHED"), "INFRA-AR": (94, "DASHED"), "INFRA-AGUA": (5, "CONTINUOUS"),
    "INFRA-DRENO": (34, "DASHED"), "INFRA-EXAUSTAO": (212, "CONTINUOUS"), "INFRA-ILUMINACAO": (42, "CONTINUOUS"),
    "TITULOS": (7, "CONTINUOUS"),
}
COR_ZONA = {"suja": 30, "lavagem": 150, "limpa": 92, "qualidade": 202, "apoio": 8, "escritorio": 52, "exclusao": 2}


def ret(r, dy=0):
    x0, x1, y0, y1 = r
    return [(x0, y0 + dy), (x1, y0 + dy), (x1, y1 + dy), (x0, y1 + dy)]


def texto(msp, s, x, y, camada, h=H_TXT, alinh=TextEntityAlignment.MIDDLE_CENTER, rot=0):
    t = msp.add_mtext(s, dxfattribs={"layer": camada, "char_height": h, "style": "ARIAL", "rotation": rot})
    t.set_location((x, y), attachment_point={TextEntityAlignment.MIDDLE_CENTER: 5, TextEntityAlignment.MIDDLE_LEFT: 4,
                                             TextEntityAlignment.TOP_LEFT: 1, TextEntityAlignment.BOTTOM_LEFT: 7}[alinh])
    return t


def poli(msp, pts, camada, fechada=True, **k):
    return msp.add_lwpolyline(pts, close=fechada, dxfattribs={"layer": camada, **k})


def galpao(msp, dy, trilhos):
    e = G.PAREDE / 2
    for x in G.PILARES_X:
        msp.add_line((x, -1300 + dy), (x, G.B + 1300 + dy), dxfattribs={"layer": "00-EIXOS", "ltscale": 20})
    for i, x in enumerate(G.PILARES_X, 1):
        texto(msp, str(i), x, -1700 + dy, "00-EIXOS", h=250)
    for y, nm in ((0, "A"), (G.B, "B")):
        msp.add_line((-1300, y + dy), (G.L + 1300, y + dy), dxfattribs={"layer": "00-EIXOS", "ltscale": 20})
        texto(msp, nm, -1700, y + dy, "00-EIXOS", h=250)
    L, B = G.L, G.B
    paredes = [(-e, e, -e, G.PORTA_FUNDOS[0]), (-e, e, G.PORTA_FUNDOS[1], B + e),
               (L - e, L + e, -e, G.PORTAO[0]), (L - e, L + e, G.PORTAO[1], B + e),
               (-e, G.ABERTURA_LATERAL[0], -e, e), (G.ABERTURA_LATERAL[1], L + e, -e, e), (-e, L + e, B - e, B + e)]
    x0, x1, y0, y1 = G.ESCRITORIO
    p = G.PAREDE
    paredes += [(x0, x1, y0, y0 + p), (x0, x1, y1 - p, y1), (x1 - p, x1, y0, G.ESCRITORIO_PORTA[0]),
                (x1 - p, x1, G.ESCRITORIO_PORTA[1], y1)]
    for r in paredes:
        pl = poli(msp, ret(r, dy), "00-PAREDES")
        h = msp.add_hatch(color=8, dxfattribs={"layer": "00-PAREDES"})
        h.paths.add_polyline_path(pl.get_points("xy"), is_closed=True)
    hb = G.BLOCO / 2
    for x in G.PILARES_X:
        for y in (0, G.B):
            poli(msp, ret((x - hb, x + hb, y - hb, y + hb), dy), "00-BLOCOS")
            px, py = G.PILAR
            pl = poli(msp, ret((x - px / 2, x + px / 2, y - py / 2, y + py / 2), dy), "00-PILARES")
            h = msp.add_hatch(color=7, dxfattribs={"layer": "00-PILARES"})
            h.paths.add_polyline_path(pl.get_points("xy"), is_closed=True)
    a, b = G.PORTA_FUNDOS
    msp.add_arc((0, a + dy), b - a, 0, 90, dxfattribs={"layer": "00-PORTAS"})
    msp.add_line((0, a + dy), (b - a, a + dy), dxfattribs={"layer": "00-PORTAS"})
    d0, d1 = G.ESCRITORIO_PORTA
    xh = G.ESCRITORIO[1] - G.PAREDE / 2
    msp.add_arc((xh, d1 + dy), d1 - d0, 180, 270, dxfattribs={"layer": "00-PORTAS"})
    msp.add_line((xh - (d1 - d0), d1 + dy), (xh, d1 + dy), dxfattribs={"layer": "00-PORTAS"})
    for dx in (-35, 30):
        msp.add_line((L + dx, G.PORTAO[0] + dy), (L + dx, G.PORTAO[1] + dy), dxfattribs={"layer": "00-PORTAS"})
    texto(msp, "PORTÃO DE ARRASTE 4.000 x 5.000", L + 450, (G.PORTAO[0] + G.PORTAO[1]) / 2 + dy, "00-PORTAS", rot=90)
    msp.add_line((G.ABERTURA_LATERAL[0], dy), (G.ABERTURA_LATERAL[1], dy), dxfattribs={"layer": "00-A-CONFIRMAR"})
    texto(msp, "vão 2.250 a identificar (D08)", sum(G.ABERTURA_LATERAL) / 2, -450 + dy, "00-A-CONFIRMAR")
    if trilhos:
        for y in G.TRILHOS_Y:
            for d in (-25, 25):
                msp.add_line((trilhos[0], y + d + dy), (trilhos[1], y + d + dy), dxfattribs={"layer": "00-TRILHOS"})


def setores(msp, lista, dy):
    for sid, nome, zona, rets, _ in lista:
        for r in rets:
            pl = poli(msp, ret(r, dy), "SETORES", color=COR_ZONA[zona])
            if zona == "exclusao":
                h = msp.add_hatch(color=2, dxfattribs={"layer": "PORTICO-FAIXA"})
                h.set_pattern_fill("ANSI31", scale=120)
                h.paths.add_polyline_path(pl.get_points("xy"), is_closed=True)
        r = rets[0]
        if zona != "exclusao":
            texto(msp, sid, r[0] + 80, r[3] - 80 + dy, "SETORES-TEXTO", h=160, alinh=TextEntityAlignment.TOP_LEFT)
        else:
            texto(msp, f"{sid} {nome}", (r[0] + r[1]) / 2, (r[2] + r[3]) / 2 + dy, "SETORES-TEXTO", h=130)
    # legenda à direita da planta: setor, nome e área (mesmos números do quadro de áreas)
    linhas = [f"{sid}  {nome} — {sum(V.area(q) for q in rets):.2f} m²".replace(".", ",")
              for sid, nome, zona, rets, _ in lista]
    texto(msp, "SETORES\\P" + "\\P".join(linhas), 22800, 8200 + dy, "SETORES-TEXTO", h=140,
          alinh=TextEntityAlignment.TOP_LEFT)


def legenda_equip(msp, lista, dy):
    tags = []
    for t, _, _ in lista:
        if t not in tags:
            tags.append(t)
    linhas = [f"{t}  {LY.EQUIP[t][0]} — {sum(1 for e in lista if e[0] == t)}x {LY.EQUIP[t][3]}" for t in tags]
    texto(msp, "EQUIPAMENTOS (envelope de reserva, mm)\\P" + "\\P".join(linhas), 22800, 4600 + dy,
          "EQUIPAMENTOS-TEXTO", h=110, alinh=TextEntityAlignment.TOP_LEFT)


def equipamentos(msp, lista, carro, dy):
    for tag, r, curto in lista:
        poli(msp, ret(r, dy), "EQUIPAMENTOS")
        if tag == "E07":
            msp.add_line((r[0], r[2] + dy), (r[1], r[3] + dy), dxfattribs={"layer": "EQUIPAMENTOS"})
            msp.add_line((r[0], r[3] + dy), (r[1], r[2] + dy), dxfattribs={"layer": "EQUIPAMENTOS"})
        if tag in ("E04", "E17"):
            cx, cy = (r[0] + r[1]) / 2, (r[2] + r[3]) / 2
            poli(msp, ret((cx - 600, cx + 600, cy - 350, cy + 350), dy), "EQUIPAMENTOS", linetype="DASHED")
        rot = 90 if (r[3] - r[2]) > 1.6 * (r[1] - r[0]) else 0
        texto(msp, tag + (f"\\P{curto}" if curto else ""), (r[0] + r[1]) / 2, (r[2] + r[3]) / 2 + dy,
              "EQUIPAMENTOS-TEXTO", h=100, rot=rot)
    poli(msp, ret(carro, dy), "EQUIPAMENTOS", linetype="DASHED")


def cota(msp, p1, p2, base, dy, ang=0, txt="<>"):
    d = msp.add_linear_dim(base=(base[0], base[1] + dy), p1=(p1[0], p1[1] + dy), p2=(p2[0], p2[1] + dy), angle=ang,
                           text=txt, dimstyle="SM", dxfattribs={"layer": "COTAS"})
    d.render()


def alternativa1(msp):
    dy = 0
    setores(msp, LY.A1_SETORES, dy)
    galpao(msp, dy, LY.TRILHOS_PROPOSTOS_X)
    equipamentos(msp, LY.A1_EQUIP, LY.A1_CARRO_LAVADORA, dy)
    legenda_equip(msp, LY.A1_EQUIP, dy)
    for (y0, y1) in LY.PERNA_FAIXA:
        poli(msp, ret((LY.PERNAS_X[0], LY.PERNAS_X[1], y0, y1)), "PORTICO-FAIXA", linetype="DASHED")
    gx, gy = LY.GANCHO_X, LY.GANCHO_Y
    poli(msp, ret((gx[0], gx[1], gy[0], gy[1])), "PORTICO-GANCHO")
    texto(msp, f"alcance do gancho (premissa) X {gx[0]:.0f}-{gx[1]:.0f}, Y {gy[0]:.0f}-{gy[1]:.0f}", gx[0] + 100,
          gy[1] + 120, "PORTICO-GANCHO", h=110, alinh=TextEntityAlignment.BOTTOM_LEFT)
    for x in LY.PERNAS_X:
        for (y0, y1) in LY.PERNA_FAIXA:
            msp.add_line((x, y0 - 80), (x, y1 + 80), dxfattribs={"layer": "PORTICO-FAIXA", "lineweight": 50})
    for nm, rs in LY.A1_CORREDORES:
        for r in rs:
            poli(msp, ret(r), "CORREDORES", linetype="DASHED")
    # cotas
    xs_a = [75, 4600, 5700, 10000, 11000, 13400, 17300, 19925]
    for a, b in zip(xs_a, xs_a[1:]):
        cota(msp, (a, LY.FAIXA_UTIL_Y[0]), (b, LY.FAIXA_UTIL_Y[0]), (a, -1300), 0)
    xs_b = [75, 4075, 4600, 5700, 8100, 10000, 13400, 17300, 19925]
    for a, b in zip(xs_b, xs_b[1:]):
        cota(msp, (a, LY.FAIXA_UTIL_Y[1]), (b, LY.FAIXA_UTIL_Y[1]), (a, 9000), 0)
    ys = [75, LY.FAIXA_UTIL_Y[0], 3300, 4550, LY.FAIXA_UTIL_Y[1], 7775]
    for a, b in zip(ys, ys[1:]):
        cota(msp, (19925, a), (19925, b), (21300, a), 0, ang=90)
    ys = [75, LY.FAIXA_UTIL_Y[0], 3700, 4900, LY.FAIXA_UTIL_Y[1], 7775]
    for a, b in zip(ys, ys[1:]):
        cota(msp, (75, a), (75, b), (-800, a), 0, ang=90)
    cota(msp, (0, 0), (G.L, 0), (0, -2300), 0, txt="20.000 (eixos)")
    # fluxos
    for fid, nome, cor, ls, pts in LY.A1_FLUXOS:
        poli(msp, pts, "FLUXOS", fechada=False)
        texto(msp, fid, pts[0][0], pts[0][1] + 150, "FLUXOS", h=110)
    # infraestrutura
    for nm, pts in LY.REDE_ELETRICA:
        poli(msp, pts, "INFRA-ELETRICA", fechada=False)
    for cp, x, y in LY.CAIXAS_PISO:
        poli(msp, ret((x - 150, x + 150, y - 150, y + 150)), "INFRA-ELETRICA")
        texto(msp, cp, x + 200, y, "INFRA-ELETRICA", h=100, alinh=TextEntityAlignment.MIDDLE_LEFT)
    for nm, pts in LY.REDE_AR:
        poli(msp, pts, "INFRA-AR", fechada=False)
    for rr, x, y in LY.ENROLADORES:
        msp.add_circle((x, y), 170, dxfattribs={"layer": "INFRA-AR"})
        texto(msp, rr, x + 250, y, "INFRA-AR", h=100, alinh=TextEntityAlignment.MIDDLE_LEFT)
    poli(msp, LY.AGUA, "INFRA-AGUA", fechada=False)
    poli(msp, LY.DRENO, "INFRA-DRENO", fechada=False)
    poli(msp, ret(LY.CANALETA), "INFRA-DRENO")
    poli(msp, ret(LY.CAIXA_CONTENCAO), "INFRA-DRENO")
    texto(msp, "CX-1", sum(LY.CAIXA_CONTENCAO[:2]) / 2, sum(LY.CAIXA_CONTENCAO[2:]) / 2, "INFRA-DRENO", h=120)
    msp.add_circle(LY.CAIXA_INSPECAO, 140, dxfattribs={"layer": "INFRA-DRENO"})
    poli(msp, LY.DUTO_EXAUSTAO, "INFRA-EXAUSTAO", fechada=False, const_width=200)
    poli(msp, ret(LY.COMPRESSOR), "INFRA-AR")
    for x, y in LY.LUMINARIAS:
        msp.add_circle((x, y), 220, dxfattribs={"layer": "INFRA-ILUMINACAO"})
    texto(msp, "ALTERNATIVA 1 (RECOMENDADA) — pórtico do projeto, fluxo em U — origem: eixos 1/A", 0, 10600, "TITULOS",
          h=350, alinh=TextEntityAlignment.BOTTOM_LEFT)


def existente(msp, orig):
    dy = DY_EXISTENTE
    for s in orig["planta_segmentos_mm"]:
        msp.add_line((s[0][0], s[0][1] + dy), (s[1][0], s[1][1] + dy), dxfattribs={"layer": "00-ORIGINAL-PDF"})
    galpao(msp, dy, G.TRILHOS_X)
    for nm, x0, x1 in G.AREAS_DEMARCADAS:
        poli(msp, ret((x0, x1, G.AREAS_Y[0], G.AREAS_Y[1]), dy), "00-AREAS-DEMARCADAS")
        texto(msp, nm, (x0 + x1) / 2, 3900 + dy, "00-AREAS-DEMARCADAS", h=150)
    texto(msp, "EXISTENTE — redesenho sobre as linhas do PDF CRV(MG)20102026 rev. 00 (camada 00-ORIGINAL-PDF)",
          0, 10600 + dy, "TITULOS", h=350, alinh=TextEntityAlignment.BOTTOM_LEFT)


def alternativa2(msp):
    dy = DY_ALT2
    setores(msp, LY.A2_SETORES, dy)
    galpao(msp, dy, None)
    equipamentos(msp, LY.A2_EQUIP, LY.A2_CARRO_LAVADORA, dy)
    legenda_equip(msp, LY.A2_EQUIP, dy)
    gx, gy = LY.PONTE_GANCHO_X, LY.PONTE_GANCHO_Y
    poli(msp, ret((gx[0], gx[1], gy[0], gy[1]), dy), "PORTICO-GANCHO")
    for y in (300, 7550):
        msp.add_line((0, y + dy), (G.L, y + dy), dxfattribs={"layer": "00-A-CONFIRMAR", "linetype": "DASHED"})
    texto(msp, "ALTERNATIVA 2 — ponte rolante (comparação; depende de estudo estrutural)", 0, 10600 + dy, "TITULOS",
          h=350, alinh=TextEntityAlignment.BOTTOM_LEFT)


def main():
    doc = ezdxf.new("R2010", setup=True)
    doc.units = units.MM
    doc.header["$MEASUREMENT"] = 1
    doc.header["$LTSCALE"] = 20
    doc.styles.add("ARIAL", font="arial.ttf")
    for nome, (cor, lt) in CAMADAS.items():
        doc.layers.add(nome, color=cor, linetype=lt)
    ds = doc.dimstyles.duplicate_entry("EZDXF", "SM")
    ds.dxf.dimtxt = 110
    ds.dxf.dimasz = 90
    ds.dxf.dimexe = 60
    ds.dxf.dimexo = 60
    ds.dxf.dimgap = 30
    ds.dxf.dimtsz = 60
    ds.dxf.dimlfac = 1
    ds.dxf.dimpost = ""
    ds.dxf.dimdec = 0
    ds.dxf.dimdsep = ord(",")
    ds.dxf.dimtxsty = "ARIAL"
    msp = doc.modelspace()
    orig = json.loads((AQUI / "desenho_original.json").read_text(encoding="utf-8"))
    alternativa1(msp)
    existente(msp, orig)
    alternativa2(msp)
    texto(msp, "ESTUDO CONCEITUAL — NÃO EXECUTIVO — NÃO LIBERADO PARA OBRA. Unidade: mm (1:1). Base: desenho "
               "CRV(MG)20102026 rev. 00. Envelopes de equipamento são reservas, não especificação.",
          0, -2900 + DY_ALT2, "TITULOS", h=200, alinh=TextEntityAlignment.TOP_LEFT)
    for nome in CAMADAS:   # abre limpo: fluxos e instalações ficam desligados até o usuário ligar
        if nome.startswith("INFRA-") or nome == "FLUXOS" or nome == "00-ORIGINAL-PDF":
            doc.layers.get(nome).off()
    texto(msp, "Camadas desligadas ao abrir: FLUXOS, INFRA-* (elétrica, ar, água, dreno, exaustão, iluminação) e "
               "00-ORIGINAL-PDF. Ligar no gerenciador de camadas.", 0, -3400 + DY_ALT2, "TITULOS", h=200,
          alinh=TextEntityAlignment.TOP_LEFT)
    doc.saveas(DXF)
    _, auditor = recover.readfile(DXF)
    print(DXF, "| erros de auditoria:", len(auditor.errors), "| entidades:", len(msp))


if __name__ == "__main__":
    main()
