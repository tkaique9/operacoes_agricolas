"""Atualiza o bloco de dados do index.html a partir dos relatórios RSC0819R.

Coloque o relatório de cada semana (RSC0819R — Relatório de Operações, exportado
em .xlsx) na pasta dados/ e rode:

    python atualizar_dados.py

Cada arquivo vira uma semana (o período vem do cabeçalho do relatório). O script
lê todos os .xlsx da pasta, monta window.OPERACOES e substitui o bloco entre os
marcadores DADOS-INICIO e DADOS-FIM do index.html. Todo o resto — totais,
indicadores, gráficos e textos — é calculado no navegador.
"""
import glob, json, os, re, sys, datetime
import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
HTML = os.path.join(AQUI, "index.html")

# rótulo curto de cada operação (o relatório traz a descrição do sistema, longa e
# truncada). Operação nova sem rótulo aqui usa a descrição limpa.
ROTULOS = {
    24565: "Broca · drone · 1ª aplic.",
    24566: "Broca · drone · 2ª aplic.",
    24593: "Cigarrinha · drone · 1ª aplic.",
    14501: "Formiga · manual · 1ª aplic.",
    24003: "Adubação cocho · 1ª aplic.",
    24004: "Adubação cocho · 2ª aplic.",
    14250: "Vinhaça enriquecida · 1ª aplic.",
    14330: "Herb. Uniport barra total · 1ª aplic.",
    24330: "Herb. Uniport barra total · 1ª aplic.",
    24331: "Herb. Uniport barra total · 2ª aplic.",
    14327: "Herb. Uniport bordadura · 1ª aplic.",
    24327: "Herb. Uniport bordadura · 1ª aplic.",
    14308: "Herb. trator bordadura · 2ª aplic.",
    14320: "Herb. arrastão bordadura · 1ª aplic.",
    14340: "Herb. manual · 1ª aplic.",
    14341: "Herb. manual · 2ª aplic.",
    14342: "Herb. manual · 3ª aplic.",
}
# tipo de aplicação: o nível entre o grupo e a operação na página 1 (herbicida ->
# Uniport, manual, trator...). Sai da descrição do sistema pela primeira palavra-chave
# que aparecer; TIPOS força o tipo de um código quando a palavra não resolve.
PALAVRAS_TIPO = [("DRONE", "Drone"), ("UNIPORT", "Uniport"), ("TRATOR", "Trator"),
                 ("ARRASTAO", "Arrastão"), ("MANUAL", "Manual"), ("COCHO", "Cocho"),
                 ("VINHACA", "Vinhaça enriquecida")]
TIPOS = {}
UNIDADES = {"LT": "L", "KG": "kg", "TON": "t", "UN": "un", "HA": "ha"}
MINUSCULAS = {"de", "do", "da", "dos", "das", "e", "d'"}
# o sistema grava sem acento algumas palavras que vão para a capa
ACENTOS = {"Condominio": "Condomínio", "Agricola": "Agrícola"}


def s(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def num(v):
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return None


def nome_bonito(t):
    """FAZ. DOS BAÚS OLHOS D' ÁGUA -> Faz. dos Baús Olhos D' Água"""
    t = s(t).replace("( ", "(").replace(" )", ")")
    t = re.sub(r"^FAZ\.(?=\S)", "FAZ. ", t, flags=re.I)     # "FAZ.DO SALTO" -> "FAZ. DO SALTO"
    out = []
    for i, w in enumerate(t.split(" ")):
        lw = w.lower()
        if re.fullmatch(r"l{1,3}", w):            # "SUCURI ll" = Sucuri II
            out.append(w.upper().replace("L", "I"))
        elif re.fullmatch(r"[ivx]+", lw) and i > 0:
            out.append(w.upper())
        elif lw in MINUSCULAS and i > 0:
            out.append(lw)
        else:
            out.append(re.sub(r"^(\(?)(\w)", lambda m: m.group(1) + m.group(2).upper(), lw))
    return " ".join(ACENTOS.get(w, w) for w in out)


def rotulo_generico(desc):
    d = re.sub(r"\(HA\)$", "", desc).strip(" (")
    return d.capitalize()


def tipo_aplicacao(cod, desc):
    if cod in TIPOS:
        return TIPOS[cod]
    for palavra, tipo in PALAVRAS_TIPO:
        if palavra in desc.upper():
            return tipo
    return "Outros"


def iso(d):
    return d.date().isoformat() if isinstance(d, datetime.datetime) else ""


def le_relatorio(path, base):
    ws = openpyxl.load_workbook(path, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    semana = {
        "ini": iso(rows[0][4]), "fim": iso(rows[0][6]), "emissao": iso(rows[0][10]),
        "empresa": nome_bonito(rows[0][1]), "arquivo": os.path.basename(path),
    }
    faz = tal = ha_tal = None
    op = None
    exec_id = None
    for r in rows[2:]:
        c0 = s(r[0])
        if c0 == "PROPRIEDADE:":
            faz = int(r[1])
            fundo = s(r[5])
            base["fazendas"][str(faz)] = nome_bonito(fundo.split(" - ", 1)[-1] if " - " in fundo else fundo)
        elif c0 == "TALHÃO:":
            tal, ha_tal = int(r[1]), num(r[4])
        elif c0 == "Operação:":
            op = int(r[1])
            desc = s(r[3])
            base["operacoes"][str(op)] = {
                "desc": desc, "rotulo": ROTULOS.get(op) or rotulo_generico(desc),
                "tipo": tipo_aplicacao(op, desc),
                "grupo": s(r[7]), "un": UNIDADES.get(s(r[5]), s(r[5])),
            }
        elif c0 == "OS" or r[0] is None:
            continue
        else:
            if c0 != "":                      # linha nova de OS: área e data
                base["n_exec"] += 1
                exec_id = base["n_exec"]
                cur = {"os": int(c0), "data": iso(r[2]), "ha": num(r[5])}
            prod = int(r[6])
            base["produtos"][str(prod)] = {"nome": s(r[7]), "un": UNIDADES.get(s(r[8]), s(r[8]))}
            # campos: e, data, faz, talhao, haTalhao, op, os, ha, prod, qtde, rec
            base["linhas"].append([
                exec_id, cur["data"], faz, tal, ha_tal, op, cur["os"], cur["ha"],
                prod, num(r[10]), num(r[11]),
            ])
    return semana


def main():
    arquivos = sorted(glob.glob(os.path.join(AQUI, "dados", "*.xlsx")))
    if not arquivos:
        sys.exit("Nenhum .xlsx em dados/")
    base = {"fazendas": {}, "operacoes": {}, "produtos": {}, "linhas": [], "n_exec": 0}
    semanas = []
    for a in arquivos:
        sem = le_relatorio(a, base)
        if any(x["ini"] == sem["ini"] and x["fim"] == sem["fim"] for x in semanas):
            print("aviso: semana repetida, ignorada:", a)
            continue
        semanas.append(sem)
        print(f"{os.path.basename(a)}: {sem['ini']} a {sem['fim']}")
    semanas.sort(key=lambda x: x["ini"])
    dados = {
        "fonte": "RSC0819R — Relatório de Operações",
        "semanas": semanas,
        "fazendas": base["fazendas"],
        "operacoes": base["operacoes"],
        "produtos": base["produtos"],
        "campos": ["e", "data", "faz", "talhao", "haTalhao", "op", "os", "ha", "prod", "qtde", "rec"],
        "linhas": base["linhas"],
    }
    js = "window.OPERACOES = " + json.dumps(dados, ensure_ascii=False, separators=(",", ":"))
    js = js.replace('],[', '],\n[')     # uma linha de produto por linha de texto
    html = open(HTML, encoding="utf-8").read()
    novo, n = re.subn(r"(/\*DADOS-INICIO\*/\n).*?(\n/\*DADOS-FIM\*/)",
                      lambda m: m.group(1) + js + ";" + m.group(2), html, flags=re.S)
    if n != 1:
        sys.exit("Marcadores DADOS-INICIO / DADOS-FIM não encontrados no index.html")
    open(HTML, "w", encoding="utf-8", newline="\n").write(novo)
    print(f"{len(base['linhas'])} linhas de produto, {base['n_exec']} apontamentos, "
          f"{len(semanas)} semana(s) -> index.html")


if __name__ == "__main__":
    main()
