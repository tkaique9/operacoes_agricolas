"""Escala rural (planilha "ESCALA RURAL", out/2026): contagens por departamento e pelo grupo que roda 3 turnos.

python rural.py  ->  rural.json  (só contagens; a planilha, com nome e matrícula, fica em dados/rural/ e não sobe para o git)

Colunas usadas: Departamento (F, o de hoje), Nome da Função (I), Dpto (J, o agrupamento novo anotado na planilha), K (escala
de hoje, 5X2) e L ("ADM" = segunda a sexta; "3t" = 3 turnos nas 24 h). Quem foi marcado "3t" (atreladores de cana picada e
de pátio, lavagem de colhedora e apoio regiões) fica junto em Atreladores Pátio, no Dpto anotado. Linhas sem matrícula
("CONTRATAR") são vagas a contratar.

Fim da safra: a planilha "quadro rural" (dados/rural/quadro_rural.xlsx) tem todo mundo ativo hoje; quem está nela e não
está na escala (pela matrícula) sai no fim da safra.
"""
import collections
import json
from pathlib import Path

import openpyxl

AQUI = Path(__file__).parent
ARQ = AQUI / "dados" / "rural" / "escala_rural.xlsx"
QUADRO = AQUI / "dados" / "rural" / "quadro_rural.xlsx"
SAIDA = AQUI / "rural.json"

NOMES = {  # como a diretoria lê (a planilha vem sem acento)
    "130.002.002": "Turma Rural · Ituiutaba II", "130.002.003": "Turma Rural · Ituiutaba III",
    "130.002.004": "Turma Rural · Capinópolis I", "130.002.050": "Turma Rural · Zona 3", "130.006.001": "Reflorestamento",
    "130.022.001": "Auxiliares · Adubação Pátio", "130.022.012": "Auxiliares · Dosador Herbicida",
    "130.022.025": "Auxiliares · Fertirrigação I", "130.022.028": "Auxiliares · Irrigação Localizada",
    "130.022.050": "Auxiliar · Apoio Regiões", "130.023.003": "Auxiliares · Atreladores Cana Picada",
    "130.023.004": "Auxiliares · Atreladores Pátio", "130.023.005": "Auxiliares · Lavagem Colhedora",
    "130.028.009": "Auxiliares · Pátio / Outros", "130.023.010": "Auxiliares · Queimadores",
}
CURTO = {"130.022.050": "Apoio Regiões", "130.023.003": "Atreladores Cana Picada", "130.023.004": "Atreladores Pátio",
         "130.023.005": "Lavagem Colhedora", "130.002.050": "Zona 3", "130.022.028": "Irrigação Localizada",
         "130.023.010": "Queimadores", "130.022.001": "Adubação Pátio", "130.028.009": "Pátio / Outros",
         "130.002.004": "Capinópolis I"}
AREAS = {  # grupo do código do departamento, com a cor da página
    "130.002": ("Turmas rurais", "#2F6FD0"), "130.006": ("Reflorestamento", "#1FA37A"),
    "130.022": ("Auxiliares de tratos", "#E0A21A"), "130.023": ("Auxiliares do CCT", "#E06A3A"),
    "130.028": ("Pátio e outros", "#8A93A6"),
}
FUNCOES = {"TRABALHADOR RURAL": "trab. rurais", "AUXILIAR AGRICOLA": "aux. agrícolas"}
UM = {"trab. rurais": "trab. rural", "aux. agrícolas": "aux. agrícola"}
GRUPOS = {"130.023.004": "Atreladores, lavagem e apoio · 3 turnos"}   # Dpto anotado dos marcados "3t"


def postos_5x2(n):
    """Maior número de pessoas por turno que a equipe de hoje cobre na 5x2 rotativa com turma fixa (3 x ceil(p x 7/5) <= n)."""
    p = 0
    while 3 * -(-(p + 1) * 7 // 5) <= n:
        p += 1
    return max(1, p)


def matricula(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def safra(linhas):
    """Quadro de hoje x escala: quem sai no fim da safra (está no quadro, não está na escala) e as vagas a contratar."""
    if not QUADRO.exists():
        return None
    quadro = {}
    for r in list(openpyxl.load_workbook(QUADRO, read_only=True, data_only=True).worksheets[0].iter_rows(values_only=True))[1:]:
        m = matricula(r[1]) if r else None
        if m and str(r[4]).strip() == "A":
            quadro[m] = str(r[6] or "").strip()
    escala = {matricula(r[1]) for r in linhas[1:] if r and matricula(r[1])}
    vagas = collections.Counter(str(r[9] or r[5] or "").split()[0] for r in linhas[1:] if r and r[1] and not matricula(r[1]))
    sai = collections.Counter(quadro[m] for m in quadro.keys() - escala)
    nome = lambda k: CURTO.get(k) or NOMES.get(k, k)
    return {"quadro": len(quadro), "fica": len(quadro.keys() & escala), "demitir": sum(sai.values()), "contratar": sum(vagas.values()),
            "demitir_por": [[nome(k), v] for k, v in sai.most_common()], "contratar_por": [[nome(k), v] for k, v in vagas.most_common()]}


def main():
    linhas = list(openpyxl.load_workbook(ARQ, read_only=True, data_only=True).worksheets[0].iter_rows(values_only=True))
    deps, grupos = {}, {}
    funcoes = collections.Counter()
    for r in linhas[1:]:
        if not r[1]:
            continue
        origem = str(r[5] or "").strip()
        dpto = str(r[9] or "").split()[0] if r[9] else origem
        funcao = FUNCOES.get(str(r[8] or "").strip(), "outros")
        hoje = str(r[10] or "").strip().lower()
        turno = str(r[11] or "").strip().lower()
        funcoes[funcao] += 1
        if turno == "3t":
            g = grupos.setdefault(dpto, {"id": dpto, "nome": GRUPOS.get(dpto, NOMES.get(dpto, dpto)), "n": 0,
                                         "hoje": hoje, "deps": collections.Counter(), "funcoes": collections.Counter()})
            g["n"] += 1
            g["deps"][origem] += 1
            g["funcoes"][funcao] += 1
            continue
        d = deps.setdefault(dpto, {"dep": dpto, "nome": NOMES.get(dpto, dpto), "area": dpto[:7], "n": 0, "dia": 0,
                                   "funcoes": collections.Counter(), "vindos": collections.Counter()})
        d["n"] += 1
        d["dia"] += 1
        d["funcoes"][funcao] += 1
        if origem != dpto:
            d["vindos"][origem] += 1
    lista = []
    for d in sorted(deps.values(), key=lambda x: x["dep"]):
        comp = [f"{v} {k if v > 1 else UM.get(k, k)}" for k, v in d["funcoes"].most_common()]
        comp += [f"{v} vindos de {CURTO.get(k, NOMES.get(k, k))}" for k, v in d["vindos"].most_common()]
        lista.append({"dep": d["dep"], "nome": d["nome"], "area": d["area"], "n": d["n"], "dia": d["dia"], "comp": " · ".join(comp)})
    saida_grupos = []
    for g in grupos.values():
        saida_grupos.append({"id": g["id"], "nome": g["nome"], "n": g["n"], "hoje": g["hoje"], "postos": postos_5x2(g["n"]),
                             "deps": [[CURTO.get(k, NOMES.get(k, k)), v] for k, v in g["deps"].most_common()],
                             "funcoes": dict(g["funcoes"])})
    saida = {"fonte": "ESCALA RURAL.xlsx (26/10/2026)", "n": len(linhas) - 1 - sum(1 for r in linhas[1:] if not r[1]),
             "hoje_txt": "5x2", "funcoes": dict(funcoes),
             "areas": {k: {"nome": v[0], "cor": v[1]} for k, v in AREAS.items()}, "deps": lista, "grupos": saida_grupos,
             "safra": safra(linhas)}
    SAIDA.write_text(json.dumps(saida, ensure_ascii=False, indent=1), encoding="utf-8")
    return saida


if __name__ == "__main__":
    s = main()
    print(s["n"], "pessoas;", s["funcoes"], ";", len(s["deps"]), "departamentos de segunda a sexta,",
          sum(d["dia"] for d in s["deps"]), "pessoas;", ", ".join(f'{g["nome"]} {g["n"]} ({g["postos"]} por turno, hoje {g["hoje"]})' for g in s["grupos"]))
    print("fim da safra:", s["safra"])
