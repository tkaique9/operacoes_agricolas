"""Quadro de áreas e verificação de consistência das alternativas (item 9 do escopo).

python verificacao.py  ->  imprime o quadro de áreas e o resultado de cada verificação
"""
import itertools

import galpao as G
import layout as LY

INTERIOR = (G.XI0, G.XI1, G.YI0, G.YI1)
# peças que dependem do gancho para entrar ou sair da posição
PRECISA_GANCHO = {"E02", "E04", "E06", "E07", "E10", "E12", "E15", "E17", "E18"}


def area(r):
    return max(0, r[1] - r[0]) * max(0, r[3] - r[2]) / 1e6


def inter(a, b):
    return (max(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), min(a[3], b[3]))


def cruza(a, b, tol=1):
    i = inter(a, b)
    return i[1] - i[0] > tol and i[3] - i[2] > tol


def contido(r, cont, tol=1):
    return r[0] >= cont[0] - tol and r[1] <= cont[1] + tol and r[2] >= cont[2] - tol and r[3] <= cont[3] + tol


def blocos():
    """Blocos de fundação nas paredes laterais (topo 150 acima do piso), em planta."""
    h = G.BLOCO / 2
    return [(x - h, x + h, y - h, y + h) for x in G.PILARES_X for y in (0, G.B)]


def alternativa(n):
    if n == 1:
        return LY.A1_SETORES, LY.A1_EQUIP, LY.A1_CORREDORES, LY.A1_CARRO_LAVADORA, (LY.GANCHO_X, LY.GANCHO_Y)
    return LY.A2_SETORES, LY.A2_EQUIP, LY.A2_CORREDORES, LY.A2_CARRO_LAVADORA, (LY.PONTE_GANCHO_X, LY.PONTE_GANCHO_Y)


def quadro_areas(n):
    """Linhas (id, nome, zona, área m², % do interior); circulação = interior - setores (sem dupla contagem)."""
    setores = alternativa(n)[0]
    linhas, soma = [], 0.0
    for sid, nome, zona, rets, _ in setores:
        a = sum(area(r) for r in rets)
        if sid == "S11":   # separa as paredes do escritório da área útil da sala
            x0, x1, y0, y1 = G.ESCRITORIO
            util = (x1 - G.PAREDE - x0) * (y1 - y0 - 2 * G.PAREDE) / 1e6
            linhas.append((sid, nome + " — área interna", zona, util))
            linhas.append(("—", "Paredes do escritório", "apoio", a - util))
        else:
            linhas.append((sid, nome, zona, a))
        soma += a
    linhas.append(("—", "Circulação e acessos (corredores, rota de fuga, áreas livres)", "circulacao",
                   G.AREA_INTERNA - soma))
    return [(i, nm, z, a, a / G.AREA_INTERNA * 100) for i, nm, z, a in linhas]


def verificar(n):
    """Lista de (item, ok, detalhe)."""
    setores, equip, corredores, carro, (gx, gy) = alternativa(n)
    res = []
    rets_set = [(sid, r) for sid, _, _, rets, _ in setores for r in rets]

    fora = [sid for sid, r in rets_set if not contido(r, INTERIOR)]
    res.append(("Setores dentro das faces internas do galpão (19,85 x 7,70 m)", not fora, ", ".join(fora) or "todos"))
    sobre = [f"{a}x{b}" for (a, ra), (b, rb) in itertools.combinations(rets_set, 2) if cruza(ra, rb)]
    res.append(("Setores sem sobreposição (áreas sem dupla contagem)", not sobre, ", ".join(sobre) or "nenhuma"))
    tot = sum(a for *_, a, _ in quadro_areas(n))
    res.append(("Soma do quadro de áreas = área interna", abs(tot - G.AREA_INTERNA) < 0.01,
                f"{tot:.2f} m² x {G.AREA_INTERNA:.2f} m²"))

    so_eq = [r for _, r, _ in equip]
    validos = [r for sid, r in rets_set if sid != "S13"]
    soltos = [t for t, r, _ in equip if not any(contido(r, s) for s in validos)]
    res.append(("Cada equipamento dentro de um setor", not soltos, ", ".join(soltos) or "todos"))
    choque = [f"{a}x{b}" for (a, ra, _), (b, rb, _) in itertools.combinations(equip, 2) if cruza(ra, rb)]
    res.append(("Equipamentos sem sobreposição entre si", not choque, ", ".join(choque) or "nenhuma"))
    bl = [t for t, r, _ in equip if any(cruza(r, b) for b in blocos())]
    res.append(("Equipamentos afastados dos blocos de fundação (ressalto de 150 mm)", not bl, ", ".join(bl) or "nenhum"))
    corr = [f"{t} no {nm}" for t, r, _ in equip for nm, rs in corredores for c in rs if cruza(r, c)]
    corr += [f"carro da lavadora no {nm}" for nm, rs in corredores for c in rs if cruza(carro, c)]
    res.append(("Corredores e rota de fuga desobstruídos", not corr, ", ".join(corr) or "livres"))
    larg = [(nm, min(c[1] - c[0], c[3] - c[2])) for nm, rs in corredores for c in rs]
    estreito = [f"{nm} {w:.0f} mm" for nm, w in larg if w < 1100]
    res.append(("Largura mínima dos corredores ≥ 1.100 mm (central ≥ 1.200)",
                not estreito and all(w >= 1200 for nm, w in larg if nm == "Corredor central"),
                "; ".join(f"{nm}: {w:.0f} mm" for nm, w in larg)))
    porta = (0, 1000, G.PORTA_FUNDOS[0], G.PORTA_FUNDOS[1])
    res.append(("Giro da porta de fundos livre", not any(cruza(r, porta) for r in so_eq), "X 0–1.000"))
    sem = [t for t, r, _ in equip if t in PRECISA_GANCHO and t != "E07"
           and not (gx[0] <= (r[0] + r[1]) / 2 <= gx[1] and gy[0] <= (r[2] + r[3]) / 2 <= gy[1])]
    cx, cy = (carro[0] + carro[1]) / 2, (carro[2] + carro[3]) / 2
    if not (gx[0] <= cx <= gx[1] and gy[0] <= cy <= gy[1]):
        sem.append("E07 (carro)")
    res.append(("Posições de içamento dentro do alcance do gancho", not sem, ", ".join(sem) or
                f"todas (gancho X {gx[0]:.0f}–{gx[1]:.0f}, Y {gy[0]:.0f}–{gy[1]:.0f})"))
    lav = [r for t, r, _ in equip if t == "E07"][0]
    s5 = [r for sid, r in rets_set if sid == "S5"][0]
    res.append(("Lavadora + zona de carga cabem no setor de lavagem", contido(lav, s5) and contido(carro, s5),
                f"envelope {lav[1]-lav[0]:.0f} x {lav[3]-lav[2]:.0f}; carro {carro[1]-carro[0]:.0f} mm"))
    if n == 1:
        exc = [r for sid, r in rets_set if sid == "S13"]
        noexc = [t for t, r, _ in equip if any(cruza(r, e) for e in exc)]
        res.append(("Nada fixo nas faixas das pernas do pórtico", not noexc, ", ".join(noexc) or "nenhum"))
        folga_esc = LY.PERNAS_X[0] - G.ESCRITORIO[1]
        folga_frente = G.XI1 - LY.PERNAS_X[1]
        res.append(("Pernas do pórtico a ≥ 500 mm do escritório e da parede do portão",
                    folga_esc >= 500 and folga_frente >= 500,
                    f"escritório {folga_esc:.0f} mm; parede do portão {folga_frente:.0f} mm"))
        y_duto = LY.DUTO_EXAUSTAO[0][1]
        res.append(("Duto de exaustão não cruza a faixa das pernas",
                    LY.PERNA_FAIXA[0][1] < y_duto < LY.PERNA_FAIXA[1][0], f"duto em Y {y_duto:.0f}"))
    return res


def main():
    for n in (1, 2):
        print(f"\n== Alternativa {n}")
        for i, nm, z, a, p in quadro_areas(n):
            print(f"{i:4s} {nm[:60]:60s} {a:7.2f} m² {p:5.1f}%")
        for item, ok, det in verificar(n):
            print(("OK   " if ok else "FALHA"), item, "—", det)


if __name__ == "__main__":
    main()
