"""Escala ADM da área agrícola (planilha "ESCALA ADM", out/2026): contagens por departamento e por central 24 h.

python adm.py  ->  adm.json  (só contagens; a planilha, com nome e matrícula, fica em dados/adm/ e não sobe para o git)

Colunas usadas: Departamento, Qtde, ESCALA (5X2 ou 12X36) e TIPO. O TIPO separa quem fica de segunda a sexta (ADM e
"sem cartão", sem cartão de ponto) das três centrais que rodam 24 h em 12x36 (Monitoramento, em amarelo na planilha,
COA e COI), cada uma com 4 pessoas: 1 por turno.
"""
import collections
import json
from pathlib import Path

import openpyxl

AQUI = Path(__file__).parent
ARQ = AQUI / "dados" / "adm" / "escala_adm.xlsx"
SAIDA = AQUI / "adm.json"

NOMES = {  # como a diretoria lê (a planilha vem sem acento)
    "101.005.002": "Controle Agrícola", "101.005.003": "Topografia", "101.005.005": "Fiscais de Área",
    "101.005.006": "Cercas", "101.005.007": "Estradas", "101.005.009": "Gestão de Transporte de Pessoal",
    "101.006.001": "Gerência CCT", "101.006.002": "Coordenação CCT", "101.006.003": "Supervisão CCT · Mecanizado",
    "101.006.005": "Supervisão · Central / Pátio", "101.006.006": "Controle CCT", "101.007.003": "Gerência de Tratos",
    "101.007.004": "Fiscais · Herbicida", "101.007.005": "Fiscais · Irrigação", "101.007.006": "Fiscais · Adubação",
    "101.007.008": "Fiscais · Controle de Pragas", "101.007.009": "Fiscais · Outros Tratos",
    "101.007.010": "Fiscais · Irrigação Localizada", "101.007.012": "Fiscais · Reflorestamento",
    "101.008.001": "Controladoria", "101.014.001": "Central de Operações Agrícolas",
    "101.014.002": "Central de Operações de Irrigação", "101.014.003": "Central de Rastreamento e Sensor de Fadiga",
    "101.014.004": "Central de Agricultura de Precisão",
}
AREAS = {  # grupo do código do departamento, na ordem da planilha, com a cor da página
    "101.005": ("Controle agrícola e apoio", "#2F6FD0"), "101.006": ("CCT · corte, carregamento e transporte", "#E06A3A"),
    "101.007": ("Tratos culturais", "#1FA37A"), "101.008": ("Controladoria", "#8A93A6"),
    "101.014": ("Centrais de operações", "#8E5BB5"),
}
CURTO = {"101.014.001": "Op. Agrícolas", "101.014.002": "Op. Irrigação", "101.014.003": "Rastreamento"}   # origem das centrais
CENTRAIS = {"MONITORAMENTO": "Monitoramento", "COA": "COA · Central de Operações Agrícolas",
            "COI": "COI · Central de Operações de Irrigação"}
TIPOS = {"ADM": "adm", "SEM CARTÃO": "sem_cartao"}


def main():
    ws = openpyxl.load_workbook(ARQ, read_only=True, data_only=True).worksheets[0]
    linhas = list(ws.iter_rows(values_only=True))
    col = {str(c).strip(): i for i, c in enumerate(linhas[0]) if c}
    deps, centrais = {}, {}
    for r in linhas[1:]:
        if not r[col["Departamento"]]:
            continue
        cod, _, nome = str(r[col["Departamento"]]).partition(" - ")
        cod, qtd = cod.strip(), int(r[col["Qtde"]] or 1)
        esc = str(r[col["ESCALA"]] or "").strip().upper().replace("X", "x")
        tipo = str(r[col["TIPO"]] or "").strip().upper()
        d = deps.setdefault(cod, {"dep": cod, "nome": NOMES.get(cod) or nome.strip().title(), "area": cod[:7], "n": 0,
                                  "esc": collections.Counter(), "tipo": collections.Counter()})
        d["n"] += qtd
        d["esc"][esc] += qtd
        if tipo in CENTRAIS:
            c = centrais.setdefault(tipo, {"id": tipo, "nome": CENTRAIS[tipo], "n": 0, "esc": collections.Counter(),
                                           "deps": collections.Counter()})
            c["n"] += qtd
            c["esc"][esc] += qtd
            c["deps"][cod] += qtd
            d["central"] = d.get("central", 0) + qtd
        else:
            d["tipo"][TIPOS.get(tipo, "outro")] += qtd
    for c in centrais.values():
        # hoje em 12x36: 4 pessoas por posto 24 h (2 de dia e 2 de noite se revezando)
        c["postos"] = max(1, round(c["esc"].get("12x36", 0) / 4))
        c["hoje"] = "12x36"
        c["deps"] = [[CURTO.get(k) or NOMES.get(k, k), v] for k, v in c["deps"].most_common()]
        c["esc"] = dict(c["esc"])
    lista = []
    for d in sorted(deps.values(), key=lambda x: x["dep"]):
        d["esc"], d["tipo"] = dict(d["esc"]), dict(d["tipo"])
        d["dia"] = d["n"] - d.get("central", 0)   # quem fica de segunda a sexta
        t = d["tipo"]
        d["comp"] = " · ".join(x for x in [f'{t["adm"]} ADM' if t.get("adm") else "", f'{t["sem_cartao"]} sem cartão' if t.get("sem_cartao") else "",
                                          f'+{d["central"]} nas centrais 24 h' if d.get("central") else ""] if x)
        lista.append(d)
    saida = {"fonte": "ESCALA ADM.xlsx (26/10/2026)", "n": sum(d["n"] for d in lista), "hoje_txt": "5x2 / 12x36",
             "areas": {k: {"nome": v[0], "cor": v[1]} for k, v in AREAS.items()},
             "deps": lista, "grupos": [centrais[k] for k in CENTRAIS if k in centrais]}
    SAIDA.write_text(json.dumps(saida, ensure_ascii=False, indent=1), encoding="utf-8")
    return saida


if __name__ == "__main__":
    s = main()
    print(s["n"], "pessoas;", len(s["deps"]), "departamentos;",
          ", ".join(f'{c["nome"]} {c["n"]} ({c["postos"]} posto)' for c in s["grupos"]))
    print("de dia:", sum(d["dia"] for d in s["deps"]), "| tipos:",
          dict(sum((collections.Counter(d["tipo"]) for d in s["deps"]), collections.Counter())))
