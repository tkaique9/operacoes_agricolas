"""FAT: motoristas e operadores com direito (planilha "validar fat", out/2026) e a fila de quem vai.

python fat.py  ->  fat.json                    (sem nomes: a fila só com papel e atividade, para a aba FAT da página)
               ->  FAT_lista_nominal.html      (com nome e matrícula: fica na máquina, fora do git)
               ->  build/fat_nominal_artifact.html  (a mesma lista no formato da página privada do claude.ai)

A planilha (dados/fat/validar_fat.xlsx) tem nome, matrícula, idade e admissão e não sobe para o git: o repositório é
público. Quem sobra sem atividade na escala escolhida (painel da frota) pode ir para o FAT; a fila diz quem vai primeiro
(pedido de 28/10/2026): a colheita (trator transbordo, colhedora e caminhão canavieiro, que param na entressafra) e
depois as outras atividades; dentro de cada grupo, menor tempo de casa primeiro. O papel vem da função, como no
pessoal_frota.py: "MOTORISTA ..." é motorista, o resto (OP. MÁQUINAS AGR.) é operador.
"""
import datetime
import json
import re
from pathlib import Path

import openpyxl

from pessoal_frota import atividade

AQUI = Path(__file__).parent
ARQ = AQUI / "dados" / "fat" / "validar_fat.xlsx"
SAIDA = AQUI / "fat.json"
TEMPLATE = AQUI / "fat_nominal_template.html"
LOCAL = AQUI / "FAT_lista_nominal.html"
ARTIFACT = AQUI / "build" / "fat_nominal_artifact.html"
FONTE = "validar fat.xlsx (28/10/2026)"
# atividades da colheita, que param na entressafra: vão primeiro (pelo departamento da planilha)
COLHEITA = re.compile(r"TRANSBORDO|COLHEDORA|CANAVIEIRO")
CRITERIO = ("Colheita primeiro (trator transbordo, colhedora e caminhão canavieiro, que param na entressafra), depois as "
            "outras atividades; dentro de cada grupo, menor tempo de casa primeiro")
# necessidade de partida na lista nominal: o que fica sem atividade na escala escolhida que está nos dados da página
# (painel da frota, 28/10/2026). Na lista dá para mudar; o número certo é o da aba FAT da página
NEC_PADRAO = {"operador": 83, "motorista": 124}


def data_iso(v):
    return v.date().isoformat() if isinstance(v, datetime.datetime) else None


def ler():
    ws = openpyxl.load_workbook(ARQ, read_only=True, data_only=True).worksheets[0]
    linhas = list(ws.iter_rows(values_only=True))
    col = {str(c).strip(): i for i, c in enumerate(linhas[0]) if c}
    gente = []
    for r in linhas[1:]:
        if not r[col["Mat."]] or str(r[col["FAT"]]).strip() != "Tem direito":
            continue
        dep = str(r[col["Departamento"]] or "").strip()
        gente.append({
            "mat": int(r[col["Mat."]]), "nome": str(r[col["Nome"]]).strip(), "funcao": str(r[col["Função"]]).strip(),
            "papel": "motorista" if "MOTORISTA" in str(r[col["Função"]]).upper() else "operador",
            "ativ": atividade(dep), "grupo": 1 if COLHEITA.search(dep.upper()) else 2,
            "ger": str(r[col["Gerência"]] or "").strip().strip("–-").strip() or "–",
            "adm": data_iso(r[col["Admissão"]]), "tempo": str(r[col["Tempo"]] or ""), "idade": r[col["Idade"]],
            "ferias": str(r[col["Férias"]] or ""), "limite": data_iso(r[col["Limite férias"]]),
        })
    # fila: colheita primeiro; dentro do grupo, menor tempo de casa (admissão mais nova) primeiro
    gente.sort(key=lambda p: (p["grupo"], -(datetime.date.fromisoformat(p["adm"]).toordinal() if p["adm"] else 0), -p["mat"]))
    return gente


def main():
    gente = ler()
    publico = {"fonte": FONTE, "criterio": CRITERIO, "n": len(gente),
               "fila": [[p["papel"], p["ativ"], p["grupo"]] for p in gente]}
    cab = json.dumps({k: v for k, v in publico.items() if k != "fila"}, ensure_ascii=False)[:-1]
    SAIDA.write_text(cab + ', "fila": [\n' + ",\n".join(json.dumps(f, ensure_ascii=False) for f in publico["fila"]) + "\n]}\n",
                     encoding="utf-8")
    dados = {"fonte": FONTE, "criterio": CRITERIO, "nec": NEC_PADRAO, "gente": gente}
    js = json.dumps(dados, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    pagina = TEMPLATE.read_text(encoding="utf-8").replace("/*FAT*/null", js)
    ARTIFACT.parent.mkdir(exist_ok=True)
    ARTIFACT.write_text(pagina, encoding="utf-8")
    LOCAL.write_text("<!doctype html>\n<html lang=\"pt-BR\" translate=\"no\" class=\"notranslate\"><head>"
                     "<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1, "
                     "viewport-fit=cover\">\n</head><body>\n" + pagina + "\n</body></html>\n", encoding="utf-8")
    por = {}
    for p in gente:
        por[p["papel"]] = por.get(p["papel"], 0) + 1
    print(f"{len(gente)} com direito: {por}; colheita {sum(1 for p in gente if p['grupo'] == 1)}")
    return publico


if __name__ == "__main__":
    main()
