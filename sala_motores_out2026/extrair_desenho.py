"""Lê o desenho CRV(MG)20102026 (PDF vetorial, A1, escala 1:50) e grava as linhas da planta em mm.

python extrair_desenho.py  ->  desenho_original.json

O PDF não tem camadas úteis nem cotas associativas: as medidas saem das próprias linhas.
1 pt no papel = 25,4/72 mm; na escala 1:50 isso vale 17,639 mm no galpão.
Origem adotada: eixo da parede de fundos (porta de 1.000) = X 0; eixo da parede do lado da
porta = Y 0. X cresce até o portão (X 20.000), Y cresce até o lado do escritório (Y 7.850).
"""
import json
import math
from pathlib import Path

from pypdf import PdfReader
from pypdf.generic import ContentStream

AQUI = Path(__file__).parent
PDF = AQUI / "dados" / "SALA DE MOTORES.pdf"
SAIDA = AQUI / "desenho_original.json"

MM_POR_PT = 25.4 / 72 * 50
# eixos de referência na página (pt), tirados das linhas de parede (pares de faces de 150 mm)
EIXO_FUNDOS_PY = (1706.5 + 1715.0) / 2      # parede de fundos
EIXO_LADO_A_PX = (590.6 + 599.2) / 2        # parede do lado da porta
JANELA_PLANTA = (540, 1110, 520, 1760)      # x0, x1, y0, y1 da "VISTA SUPERIOR - PISO"
JANELA_CORTE = (1140, 1640, 1770, 2300)     # "SECÇÃO S1-S1"


def mul(a, b):
    a0, a1, a2, a3, a4, a5 = a
    b0, b1, b2, b3, b4, b5 = b
    return [a0 * b0 + a1 * b2, a0 * b1 + a1 * b3, a2 * b0 + a3 * b2, a2 * b1 + a3 * b3,
            a4 * b0 + a5 * b2 + b4, a4 * b1 + a5 * b3 + b5]


def segmentos(pagina, leitor):
    """Todos os segmentos traçados (S, s, b, B), em pt da página, com curvas aproximadas."""
    ctm, pilha, sub, cur, out = [1, 0, 0, 1, 0, 0], [], [], [], []

    def p(x, y):
        return (ctm[0] * x + ctm[2] * y + ctm[4], ctm[1] * x + ctm[3] * y + ctm[5])

    def tracar(fechar):
        nonlocal sub, cur
        if cur:
            sub.append(cur)
        for pts in sub:
            if fechar and len(pts) > 1 and pts[0] != pts[-1]:
                pts = pts + [pts[0]]
            out.extend(zip(pts, pts[1:]))
        sub, cur = [], []

    for ops, op in ContentStream(pagina.get_contents(), leitor).operations:
        op = op.decode("latin-1")
        v = [float(x) for x in ops] if op in ("cm", "m", "l", "c", "re") else None
        if op == "q":
            pilha.append(ctm[:])
        elif op == "Q":
            ctm = pilha.pop()
        elif op == "cm":
            ctm = mul(v, ctm)
        elif op == "m":
            if cur:
                sub.append(cur)
            cur = [p(*v)]
        elif op == "l":
            cur.append(p(*v))
        elif op == "c":
            (x0, y0), pts = cur[-1], [p(v[i], v[i + 1]) for i in (0, 2, 4)]
            for k in range(1, 9):
                t = k / 8
                m = 1 - t
                cur.append(tuple(m ** 3 * a + 3 * m * m * t * b + 3 * m * t * t * c + t ** 3 * d
                                 for a, b, c, d in zip((x0, y0), *pts)))
        elif op == "re":
            x, y, w, h = v
            sub.append([p(x, y), p(x + w, y), p(x + w, y + h), p(x, y + h), p(x, y)])
        elif op == "h" and cur:
            cur.append(cur[0])
        elif op in ("S", "s"):
            tracar(op == "s")
        elif op in ("b", "b*", "B", "B*"):
            tracar(True)
        elif op in ("n", "f", "F", "f*"):
            sub, cur = [], []
    return out


def dentro(seg, jan):
    x0, x1, y0, y1 = jan
    return all(x0 <= x <= x1 and y0 <= y <= y1 for x, y in seg)


def planta_mm(seg):
    """pt da página -> mm do galpão (a planta está girada 90° na folha)."""
    return [[round((EIXO_FUNDOS_PY - y) * MM_POR_PT, 1), round((x - EIXO_LADO_A_PX) * MM_POR_PT, 1)]
            for x, y in seg]


def linhas_horizontais_corte(segs):
    """No corte, as linhas horizontais da vista têm x da página constante; devolve as posições."""
    xs = {}
    for (a, b) in segs:
        if dentro((a, b), JANELA_CORTE) and abs(a[0] - b[0]) < 0.05 and abs(a[1] - b[1]) > 20:
            xs.setdefault(round(a[0], 1), 0)
            xs[round(a[0], 1)] += abs(a[1] - b[1])
    return xs


def main():
    leitor = PdfReader(PDF)
    pag = leitor.pages[0]
    segs = segmentos(pag, leitor)

    planta = [planta_mm(s) for s in segs if dentro(s, JANELA_PLANTA)]
    planta = [s for s in planta if math.dist(*s) > 1]

    # alturas do corte S1-S1 (x da página das linhas horizontais), conferidas na leitura:
    # 1261.7 topo do piso | 1270.2 topo do bloco | 1497.0 base da viga | 1514.5 topo da viga | 1567.0 forro
    piso = 1261.7
    alturas = {nome: round((x - piso) * MM_POR_PT)
               for nome, x in {"topo_bloco": 1270.2, "base_viga_portico": 1497.0,
                               "topo_viga_portico": 1514.5, "forro": 1567.0}.items()}
    horiz = linhas_horizontais_corte(segs)
    for x in (1261.7, 1270.2, 1497.0, 1514.5, 1567.0):
        assert any(abs(x - k) < 0.3 for k in horiz), f"linha do corte não encontrada em x={x}"

    def medida(a, b):
        return round(abs(a - b) * MM_POR_PT)

    conferencia = [
        ("Comprimento entre eixos (cota 20.000)", 20000, medida(1710.75, 577.0)),
        ("Largura entre eixos (cota 7.850)", 7850, medida(594.9, (1035.7 + 1044.2) / 2)),
        ("Vão dos trilhos do pórtico (cota 6.534)", 6534, medida((630.8 + 633.6) / 2, (1001.3 + 1004.0) / 2)),
        ("Escritório, profundidade (cota 3.000)", 3000, medida(817.4, 987.5)),
        ("Portão (cota 4.000)", 4000, round(226.8 * MM_POR_PT)),
        ("Espessura da parede (cota 150)", 150, medida(590.6, 599.2)),
        ("Bloco de fundação na planta (cotas 600 e 700)", 700, medida(557.0, 596.8)),
        ("Altura da viga do pórtico (cota 4.460)", 4460, alturas["topo_viga_portico"]),
        ("Espessura da laje (detalhe: 100)", 100, medida(1256.0, 1261.7)),
    ]
    print(f"{'item':52s} {'cota':>7s} {'medido':>7s}")
    for nome, cota, med in conferencia:
        print(f"{nome:52s} {cota:7d} {med:7d}")

    SAIDA.write_text(json.dumps({
        "fonte": "CRV(MG)20102026 - GALPÃO PARA MOTOR A DIESEL - rev. 00 de 16/10/2026",
        "mm_por_pt": MM_POR_PT,
        "planta_segmentos_mm": planta,
        "alturas_corte_mm": alturas,
        "conferencia": [{"item": a, "cota": b, "medido": c} for a, b, c in conferencia],
    }, ensure_ascii=False), encoding="utf-8")
    print(SAIDA, len(planta), "segmentos")


if __name__ == "__main__":
    main()
