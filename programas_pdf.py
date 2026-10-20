"""Tabelas dos programas de tratos lidas da impressão das planilhas (PDF).

O PDF semanal (ex.: 19.10.2026.pdf) traz as abas de cada programa como são impressas:
pré-emergência cana soca, reforço bordadura (3ª molécula), 1ª catação química e
adubação de soqueira (localizada e com cocho). Cada tabela tem as fazendas, o subtotal
por região (ou por bloco, na adubação) e o total, e é lida assim mesmo, sem refazer
conta: a apresentação mostra o que a planilha mostra.

A leitura é pela posição das palavras: cada número vai para a coluna cujo título
termina mais perto da borda direita dele (os números são alinhados à direita).
Precisa de: pip install pymupdf
"""
import re
import unicodedata

DATA = re.compile(r"^\d{2}/\d{2}/\d{4}$")
NUM = re.compile(r"^-?[\d.]+,\d+$|^-$")


def sem_acento(t):
    return "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn").upper()


def num(t):
    if t == "-":
        return 0.0
    return float(t.replace(".", "").replace(",", "."))


# colunas numéricas de cada tipo de tabela: (chave, título, palavra do título que fecha a
# coluna). A palavra é procurada no cabeçalho da esquerda para a direita, cada uma
# depois da anterior.
COLUNAS = {
    "pre": [("colhida", "Área colhida", "COLHIDA"), ("aColher", "Área a colher", "COLHER"),
            ("reforma", "Área reforma", "REFORMA"), ("realizada", "Área realizada", "REALIZADA"),
            ("aRealizar", "A realizar", "REALIZAR"), ("_obs", "", "OBSERVACAO"),
            ("fornecedor", "Área fornecedor", "FORNECEDOR")],
    "bord": [("aberta", "Área aberta", "ABERTA"), ("realizada", "Área realizada", "REALIZADA"),
             ("aRealizar", "A realizar", "REALIZAR")],
    "cat1": [("corte", "Corte", "CORTE"), ("aberta", "Área aberta", "ABERTA"), ("reforma", "Área reforma", "REFORMA"),
             ("cataMatoMec", "Cata mato MEC", "MATO"), ("realMec", "Realizada MEC", "REALIZADA"),
             ("catManual", "Catação manual", "MANUAL"), ("realMan", "Realizada MAN", "REALIZADA"),
             ("realTotal", "Realizado total", "TOTAL"), ("maior20", "A realizar > 20 dias", "DIAS"),
             ("menor20", "A realizar < 20 dias", "20")],
    "adub": [("aTratar", "Área a tratar", "TRATAR"), ("canaMuda", "Aplicada cana muda", "MUDA"),
             ("reforma", "Área de reforma", "REFORMA"), ("localizada", "Aplicada localizada", "LOCALIZADA"),
             ("cocho", "Aplicada cocho", "COCHO"), ("aRealizar", "A realizar", "REALIZAR")],
}
TEXTO = {  # colunas de texto: (chave, título)
    "pre": [("reg", "Região"), ("cod", "Cód Faz"), ("bloco", "Bloco"), ("prop", "Propriedade"),
            ("data", "Data colheita"), ("atraso", "Dias de atraso")],
    "bord": [("reg", "Região"), ("cod", "Cód Faz"), ("bloco", "Bloco"), ("prop", "Propriedade"),
             ("data", "Data colheita"), ("atraso", "Dias de atraso")],
    "cat1": [("reg", "Região"), ("cod", "Cód Faz"), ("bloco", "Bloco"), ("prop", "Propriedade"),
             ("data", "Data colheita"), ("atraso", "Dias de atraso")],
    "adub": [("reg", "Região"), ("cod", "Cód Faz"), ("prop", "Propriedade"), ("bloco", "Bloco")],
}


def linhas_da_pagina(pg):
    """Palavras agrupadas em linhas (mesma altura, ±3 pt), da esquerda para a direita."""
    L = []
    for w in sorted(pg.get_text("words"), key=lambda w: (w[3], w[0])):
        if L and abs(L[-1][0] - w[3]) <= 3:
            L[-1][1].append(w)
        else:
            L.append([w[3], [w]])
    return [(y, sorted(ws, key=lambda w: w[0])) for y, ws in L]


def texto(ws):
    return " ".join(w[4] for w in ws)


class Layout:
    """Posições das colunas, tiradas do cabeçalho da página (linha 'Região Cód Faz ...'
    e as duas linhas de baixo, onde ficam as segundas linhas dos títulos)."""

    def __init__(self, tipo, cab):
        self.tipo = tipo
        ws = sorted(cab, key=lambda w: w[2])
        self.ancoras, x = [], -1
        for chave, _t, palavra in COLUNAS[tipo]:
            w = next(w for w in ws if w[2] > x and sem_acento(w[4]) == palavra)
            self.ancoras.append((chave, w[2]))
            x = w[2]
        pos = {sem_acento(w[4]): w for w in sorted(cab, key=lambda w: -w[0])}   # primeira ocorrência
        self.x_cod = pos["COD"][0]
        if tipo == "adub":
            # títulos centralizados: a fronteira entre propriedade e bloco fica no meio deles
            self.x_bloco = (pos["PROPRIEDADE"][0] + pos["PROPRIEDADE"][2] + pos["BLOCO"][0] + pos["BLOCO"][2]) / 4 + 6
            self.x_prop = pos["FAZ"][2] + 2
        else:
            self.x_bloco = pos["BLOCO"][0] - 12
            self.x_prop = (pos["BLOCO"][2] + pos["PROPRIEDADE"][0]) / 2 + 20
        self.x_num = self.ancoras[0][1] - 42                 # números começam perto da 1ª coluna
        self.x_obs = self.ancoras[-1][1] + 6                 # depois da última, observação

    def coluna(self, w):
        chave, x = min(self.ancoras, key=lambda a: abs(a[1] - w[2]))
        return chave if abs(x - w[2]) <= 30 else None

    def valores(self, ws):
        v = {}
        for w in ws:
            if w[0] >= self.x_num and NUM.match(w[4]):
                c = self.coluna(w)
                if c and not c.startswith("_"):
                    v[c] = num(w[4])
        return v


def linha_fazenda(lay, ws, blocos):
    r = {"t": "faz", "reg": "", "cod": "", "bloco": "", "prop": "", "data": "", "atraso": None, "obs": ""}
    i_cod = next(i for i, w in enumerate(ws) if re.fullmatch(r"9\d{3}", w[4]) and w[0] >= lay.x_cod - 25)
    r["reg"] = " ".join(w[4] for w in ws[:i_cod])
    r["cod"] = ws[i_cod][4]
    prop, bloco, obs = [], [], []
    resto = [w for w in ws[i_cod + 1:]]
    for k, w in enumerate(resto):
        t = w[4]
        if DATA.match(t):
            r["data"] = t
        elif t == "DIAS" and k and re.fullmatch(r"\d+", resto[k - 1][4]):
            r["atraso"] = int(resto[k - 1][4])
        elif re.fullmatch(r"\d+", t) and k + 1 < len(resto) and resto[k + 1][4] == "DIAS":
            continue
        elif w[0] >= lay.x_num and (NUM.match(t) or (lay.tipo == "cat1" and re.fullmatch(r"\d+", t))):
            continue
        elif w[0] >= lay.x_obs and lay.tipo == "adub":
            obs.append(t)
        elif lay.tipo == "adub":
            if w[0] < lay.x_bloco <= w[2] - 4:
                # nome da propriedade invade a coluna do bloco ("...CONCEIÇUSINA"): separa
                # pelo nome de bloco conhecido no fim da palavra
                b = next((b for b in sorted(blocos, key=len, reverse=True) if t.endswith(b) and len(t) > len(b)), None)
                if b:
                    prop.append(t[:-len(b)]), bloco.append(b)
                    continue
            (bloco if w[0] >= lay.x_bloco else prop).append(t)
        else:
            (prop if w[0] >= lay.x_prop else bloco).append(t)
    r["prop"], r["bloco"], r["obs"] = " ".join(prop), " ".join(bloco), " ".join(obs)
    v = lay.valores(ws)
    if "corte" in v:
        r["corte"] = int(v.pop("corte"))
    else:
        # o corte é um inteiro sem vírgula: pega pela posição
        for w in ws:
            if lay.tipo == "cat1" and re.fullmatch(r"\d+", w[4]) and lay.coluna(w) == "corte":
                r["corte"] = int(w[4])
    r["v"] = v
    return r


def linha_total(lay, ws, tipo_linha):
    nums = [w for w in ws if w[0] >= lay.x_num and NUM.match(w[4])]
    rot = " ".join(w[4] for w in ws if w not in nums and w[0] < lay.x_num)
    v = lay.valores(ws)
    v.pop("corte", None)
    return {"t": tipo_linha, "rot": re.sub(r"\s+", " ", rot).strip(), "v": v}


PROGRAMAS = [  # título da página -> (id, tipo de tabela, nome)
    ("PRE-EMERGENCIA - CANA SOCA", "pre", "pre", "Pré-emergência — cana soca"),
    ("REFORCO BORDADURA", "bord", "bord", "Reforço bordadura (3ª molécula) — cana soca"),
    ("CATACAO QUIMICA - CANA SOCA", "cat1", "cat1", "1ª catação química — cana soca"),
]


def resumo_final(prog, linhas, i):
    """Quadro "RESUMO FINAL ADUBAÇÃO": a linha de títulos e, logo abaixo, os valores."""
    nomes = [("aTratar", "TRATAR"), ("reforma", "REFORMA"), ("localizada", "LOCALIZADA"), ("cocho", "COCHO"), ("aRealizar", "REALIZAR")]
    x0 = min(w[0] for w in linhas[i][1] if "RESUMO" in sem_acento(w[4])) - 60
    for k in range(i + 1, min(i + 6, len(linhas))):
        cab = [w for w in linhas[k][1] if w[0] >= x0]
        if any(sem_acento(w[4]) == "TRATAR" for w in cab):
            anc = [(c, next(w[2] for w in cab if sem_acento(w[4]) == p)) for c, p in nomes]
            for kk in range(k + 1, min(k + 4, len(linhas))):
                vals = [w for w in linhas[kk][1] if w[0] >= x0 and NUM.match(w[4])]
                if len(vals) >= 4:
                    prog["extras"]["resumo"] = {min(anc, key=lambda a: abs(a[1] - w[2]))[0]: num(w[4]) for w in vals}
                    return


def regioes_do_grafico(prog, linhas, i):
    """1ª catação: os valores por região só existem no gráfico (rótulos das barras).
    Cada região tem 5 barras, na ordem da legenda: área colhida, área realizada,
    a realizar < 20 dias, a realizar > 20 dias e falta aplicar."""
    rot = [w for w in linhas[i][1] if sem_acento(w[4]) == "REGIAO"]
    if len(rot) < 3 or "regioes" in prog["extras"]:
        return
    centros = [(w[0] + w[2]) / 2 + 3 for w in rot]
    y_reg = linhas[i][0]
    # rótulos numéricos acima das regiões, até o título do gráfico de cima
    # entre o título do gráfico de baixo e a linha das regiões
    y_top = max([y for y, ws in linhas[:i] if "GRAFICO" in sem_acento(texto(ws))] or [y_reg - 230])
    vals = [w for y, ws in linhas[:i] if y_top < y < y_reg for w in ws
            if re.fullmatch(r"[\d.]+,\d{2}", w[4]) and w[0] > 90]
    chaves = ["colhida", "realizada", "menor20", "maior20", "faltaAplicar"]
    regs = {}
    for n, c in enumerate(centros):
        meio = [(centros[n - 1] + c) / 2 if n else -1e9, (c + centros[n + 1]) / 2 if n + 1 < len(centros) else 1e9]
        dela = sorted([w for w in vals if meio[0] <= (w[0] + w[2]) / 2 < meio[1]], key=lambda w: w[0])
        if len(dela) == 5:
            regs[str(n + 1)] = {k: num(w[4]) for k, w in zip(chaves, dela)}
    if len(regs) == len(centros):
        prog["extras"]["regioes"] = regs


def ler(path):
    import pymupdf
    doc = pymupdf.open(path)
    progs = {}
    atualizacao = None
    secao = None          # adubação: seção corrente (localizada, área de queima, cocho por equipe e região)
    lay_ant = {}          # layout da página anterior do mesmo programa (a tabela continua na seguinte)
    # nomes de bloco da adubação, tirados dos subtotais: servem para separar a propriedade
    # que a impressão grudou no bloco ("RETIRINHO E BOA ESPERANÇABEBEDOURO")
    blocos = {re.sub(r"\s+[\d.,-]+(\s+[\d.,-]+)*$", "", re.sub(r"^SUBTOTAL\s*-\s*", "", texto(ws))).strip()
              for pg in doc for _y, ws in linhas_da_pagina(pg) if texto(ws).startswith("SUBTOTAL -")}
    for pg in doc:
        linhas = linhas_da_pagina(pg)
        todo = sem_acento(" ".join(texto(ws) for _y, ws in linhas[:8]))
        m = re.search(r"ATUALIZA\S*\s*(\d{2})/(\d{2})/(\d{4})", sem_acento(" ".join(texto(ws) for _y, ws in linhas[:8])))
        if m and not atualizacao:
            atualizacao = f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
        if "ADUBACAO SOCARIA" in todo:
            pid, tipo, nome = "adub", "adub", "Adubação de soqueira"
        else:
            achou = next((p for p in PROGRAMAS if p[0] in todo), None)
            if not achou:
                continue
            _t, pid, tipo, nome = achou
        prog = progs.setdefault(pid, {"id": pid, "nome": nome, "tipo": tipo, "secoes": [], "extras": {}})
        lay = lay_ant.get(pid)
        for i, (y, ws) in enumerate(linhas):
            t = sem_acento(texto(ws))
            if sem_acento(ws[0][4]) == "REGIAO" and "COD" in t:
                cab = [w for yy, l in linhas[i:i + 4] if yy - y <= 16 for w in l]
                lay = lay_ant[pid] = Layout(tipo, cab)
                continue
            # seções da adubação
            if tipo == "adub":
                ms = re.match(r"^(ADUBACAO LOCALIZADA|AREA DE QUEIMA|ADUBACAO COM COCHO)(.*)$", t)
                if ms and not any(NUM.match(w[4]) and w[0] > 300 for w in ws):
                    titulo = texto([w for w in ws if not re.fullmatch(r"\d{2}/\d{2}/\d{4}", w[4]) and w[0] < 600])
                    secao = {"titulo": titulo.title().replace(" Com ", " com ").replace(" Do ", " do ").replace(" De ", " de "), "linhas": []}
                    prog["secoes"].append(secao)
                    continue
                if ms and lay:      # "ADUBAÇÃO COM COCHO - ALENCAR  6.961,83 ..." = total da equipe
                    tot = linha_total(lay, ws, "tot")
                    tot["rot"] = "TOTAL " + tot["rot"]
                    if secao is not None:
                        secao["linhas"].append(tot)
                    continue
            if tipo == "adub" and "RESUMO FINAL" in t:
                resumo_final(prog, linhas, i)
                continue
            if tipo == "cat1" and t.startswith("REGIAO 1 REGIAO 2"):
                regioes_do_grafico(prog, linhas, i)
                continue
            if not lay:
                continue
            if tipo == "pre" and ws[0][0] > 500:
                for chave in ("TOTAL ENCERRADO", "TOTAL ABERTO", "GERAL"):
                    if t.startswith(chave) and NUM.match(ws[-1][4]):
                        prog["extras"][chave.title()] = num(ws[-1][4])
                continue
            if ws[0][0] > 80:
                continue
            if tipo != "adub" and not prog["secoes"]:
                prog["secoes"].append({"titulo": "", "linhas": []})
            alvo = prog["secoes"][-1]["linhas"] if prog["secoes"] else None
            if alvo is None:
                continue
            if t.startswith("SUBTOTAL"):
                r = linha_total(lay, ws, "sub")
                if tipo == "adub":
                    blocos.add(re.sub(r"^SUBTOTAL - ", "", r["rot"]))
                alvo.append(r)
            elif t.startswith("TOTAL"):
                alvo.append(linha_total(lay, ws, "tot"))
            elif any(re.fullmatch(r"9\d{3}", w[4]) for w in ws[:3]) and any(NUM.match(w[4]) for w in ws):
                alvo.append(linha_fazenda(lay, ws, blocos))
    # região com erro de fórmula da planilha (#N/D, #REF!...): vale a da fazenda de cima
    # no mesmo quadro (ex.: a 9849 São João, no meio das áreas Cargill)
    for p in progs.values():
        for sec in p["secoes"]:
            ant = None
            for l in sec["linhas"]:
                if l["t"] != "faz":
                    continue
                if l["reg"].startswith("#") and ant:
                    l["reg"] = ant
                ant = l["reg"] if not l["reg"].startswith("#") else ant
    # nome da coluna para a apresentação
    for p in progs.values():
        p["cols"] = [{"k": k, "t": tt} for k, tt in TEXTO[p["tipo"]]] + \
                    [{"k": k, "t": tt, "n": True} for k, tt, _w in COLUNAS[p["tipo"]] if not k.startswith("_")]
    ordem = ["pre", "bord", "cat1", "adub"]
    return {"atualizacao": atualizacao, "programas": [progs[k] for k in ordem if k in progs]}


# ------------------------------------------------ resumo semanal (página 1)
# A primeira página do PDF é a aba Semanal-Acumulado impressa: para cada atividade, o
# realizado de segunda a sábado, o quadro de meta (nº de equipamentos/pessoas, meta por
# unidade, meta do dia) e os realizados do dia, da semana e do mês, mais o mês anterior.
# Os dias vão para a coluna da data mais perto; o resto pela borda direita do número,
# que é sempre a mesma na impressão (conferido: a soma dos dias bate com a semana).
DIREITA = [("equip", 423), ("metaUnid", 461), ("metaDia", 492), ("realDia", 524), ("_vDia", 548),
           ("metaSemana", 586), ("realSemana", 624), ("_vSem", 648), ("metaMes", 688),
           ("realMes", 726), ("_vMes", 751), ("mesAnt", 790), ("_vAnt", 818)]
NAO_ATIVIDADE = ("PESSOAS DIA", "RENDIMENTO", "OBSERVACAO", "ATIVIDADE", "PAGINA")


def ler_semana(path):
    import datetime
    import pymupdf
    doc = pymupdf.open(path)
    pg = next(p for p in doc if "RESUMO SEMANAL" in sem_acento(p.get_text()))
    linhas = linhas_da_pagina(pg)
    texto_pg = sem_acento(" ".join(texto(ws) for _y, ws in linhas[:6]))
    m = re.search(r"DIA REF:?\s*(\d{2})/(\d{2})/(\d{4})", texto_pg)
    dia_ref = f"{m.group(3)}-{m.group(2)}-{m.group(1)}" if m else None
    m = re.search(r"ATUALIZACAO:?\s*(\d{2})/(\d{2})/(\d{4})", texto_pg)
    atualizacao = f"{m.group(3)}-{m.group(2)}-{m.group(1)}" if m else None
    ano = int(dia_ref[:4]) if dia_ref else datetime.date.today().year
    datas = []          # (centro x, data iso) das colunas de dia do bloco corrente
    frente = None
    ativs, atual = [], None
    for _y, ws in linhas:
        t = sem_acento(texto(ws))
        dd = [w for w in ws if re.fullmatch(r"\d{2}/\d{2}", w[4]) and 150 < w[0] < 380]
        if len(dd) >= 5:
            datas = [((w[0] + w[2]) / 2, f"{ano}-{w[4][3:5]}-{w[4][:2]}") for w in dd]
            continue
        nome = " ".join(w[4] for w in ws if w[2] < 160)
        nums = [w for w in ws if w[0] >= 150 and NUM.match(w[4])]
        if not nome:
            continue
        if not nums:
            if nome.isupper() and len(nome) > 8:
                frente = nome
            continue
        dias, quadro = {}, {}
        for w in nums:
            cx = (w[0] + w[2]) / 2
            if cx < 380 and datas:
                d = min(datas, key=lambda c: abs(c[0] - cx))[1]
                if w[4] != "-":
                    dias[d] = num(w[4])
            elif cx >= 380:
                ref = w[2] if w[4] != "-" else cx
                k, x = min(DIREITA, key=lambda a: abs(a[1] - ref))
                if abs(x - ref) <= 14 and not k.startswith("_") and w[4] != "-":
                    quadro[k] = num(w[4])
        tn = sem_acento(nome)
        if tn.startswith("PESSOAS DIA") and atual is not None:
            atual["pessoas"] = dias
            atual["pessoasPlan"] = quadro.get("equip")
            atual["metaPessoa"] = quadro.get("metaUnid")
        elif not tn.startswith(NAO_ATIVIDADE):
            atual = {"nome": nome, "frente": frente, "dias": dias, "pessoas": None, **quadro}
            ativs.append(atual)
    return {"diaRef": dia_ref, "atualizacao": atualizacao, "datas": [d for _x, d in datas], "atividades": ativs}


if __name__ == "__main__":
    import json, sys
    d = ler(sys.argv[1])
    for p in d["programas"]:
        print("==", p["nome"], [s["titulo"] for s in p["secoes"]], p["extras"])
        for s in p["secoes"]:
            for r in s["linhas"]:
                if r["t"] == "faz":
                    print("  ", r["reg"], r["cod"], "|", r["bloco"], "|", r["prop"], "|", r["data"], r["atraso"], r.get("corte", ""), r["v"], r["obs"])
                else:
                    print("  **", r["t"], r["rot"], r["v"])
