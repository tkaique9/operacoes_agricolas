"""Atualiza a reunião semanal de Tratos Culturais (index.html).

Coloque na pasta dados/ as duas planilhas da semana e rode:

    python atualizar_tratos.py

- "Operações Tratos Culturais - 2025-2026" (.xlsb): aba Semanal-Acumulado, com o
  realizado por dia de cada atividade, as metas e a semana de referência.
- "Operações Tratos Culturais - SAFRA 2026" (.xlsx): as abas de cada programa
  (pré-emergente, catações, 3ª molécula, adubação), com o realizado e o saldo a
  realizar por fazenda e por bloco.

O script acha cada planilha pelas abas que ela tem, não pelo nome do arquivo.
Grava window.TRATOS entre os marcadores DADOS-INICIO e DADOS-FIM do index.html;
todo o resto (totais, percentuais, gráficos) é calculado no navegador.
Precisa de: pip install openpyxl pyxlsb
"""
import calendar, datetime, glob, json, os, re, sys
import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
HTML = os.path.join(AQUI, "index.html")
EPOCA = datetime.datetime(1899, 12, 30)

# nome curto de cada frente (o bloco da aba Semanal-Acumulado)
FRENTES = [
    ("APLICAÇÃO MANUAL", "manual", "Manual e semi-mecanizada"),
    ("UNIPORT", "uniport", "Uniport"),
    ("AÉREA", "aerea", "Aplicação com drone"),
    ("QUEBRA LOMBO", "lombo", "Quebra lombo"),
    ("COCHO", "cocho", "Adubação com cocho"),
]
# frentes que não entram na apresentação (nem nas páginas, nem nos totais). Quebra
# lombo saiu a pedido em 13/10/2026: não é aplicação desta operação. Para voltar,
# tire o id daqui.
FRENTES_FORA = {"lombo"}
# atividades que saem da frente da planilha e viram uma operação própria, logo depois
# dela: (frente de origem, trecho do nome da atividade, id, nome, título). O drone de
# monitoramento saiu da aplicação aérea a pedido em 20/10/2026.
SEPARA = [("aerea", "monitoramento", "monit", "Monitoramento com drone", "MONITORAMENTO COM DRONE")]


def separa_frentes(frentes):
    for orig, trecho, fid, nome, titulo in SEPARA:
        i = next((k for k, fr in enumerate(frentes) if fr["id"] == orig), None)
        if i is None or any(fr["id"] == fid for fr in frentes):
            continue
        mov = [a for a in frentes[i]["atividades"] if trecho in a["nome"].lower()]
        if mov:
            frentes[i]["atividades"] = [a for a in frentes[i]["atividades"] if a not in mov]
            frentes.insert(i + 1, {"id": fid, "nome": nome, "titulo": titulo, "atividades": mov})
    return frentes
# meta de hectares por pessoa/dia que vale no lugar da que vem da planilha. Cata Mato:
# a linha "Pessoas Dia" traz 8 ha (120 ha ÷ 15 pessoas), mas a meta é 4 ha/pessoa
# (corrigido em 13/10/2026). Para voltar à da planilha, tire a atividade daqui.
META_PESSOA = {"Cata Mato": 4.0}
# nº de equipamentos que vale no lugar do da planilha; a meta do dia, da semana e do
# mês acompanham (equipamentos × meta por equipamento). Quadriciclo: a planilha traz
# 1, mas operam 2, com meta de 10 ha cada: 20 ha/dia (corrigido em 13/10/2026).
# Catação de Formiga: 3 equipamentos de 10 ha cada, 30 ha/dia (a planilha traz 1 de
# 20 ha; eram 2 desde 13/10/2026, 3 desde 27/10/2026).
EQUIP = {"Quadriciclo": 2, "Catação de Formiga": 3}
# meta de hectares por equipamento/dia que vale no lugar da planilha; aplicada antes
# de EQUIP, que refaz a meta do dia com ela
META_EQUIP = {"Catação de Formiga": 10.0}
# dias da semana sem meta para a atividade, além do domingo (que nunca tem): 5 = sábado
# (datetime.weekday). O drone de monitoramento não voa no sábado (20/10/2026): a meta
# do sábado sai e a da semana e a do mês contam só os dias em que ele voa.
SEM_META_DIAS = {"Área Monitoramento": [5]}
# datas avulsas sem meta para a atividade (dia em que ela não operou e a meta não vale),
# a pedido; a meta da semana e a do mês contam sem elas. O drone de monitoramento não
# voou na sexta 23/10/2026 (pedido em 27/10/2026; nas outras sextas ele voa).
SEM_META_DATAS = {"Área Monitoramento": ["2026-10-23"]}
# atividades não planejadas: entram com o realizado e as pessoas apontadas, mas sem
# meta (nem na operação, nem no total). O Roço passou a ser não planejado em 20/10/2026.
NAO_PLANEJADAS = {"Roço"}
# linhas que a planilha não soma na semana (sem "Realizado Semana" e fora do total da
# frente), mas que entram no relatório: o realizado da semana é a soma dos dias. Nome da
# planilha → (nome no relatório, atividade de baixo da qual a linha aparece). A bordadura
# (3ª molécula) do Uniport próprio entrou em 27/10/2026.
# O Uniport Pingente Inseticida entrou do mesmo jeito em 03/11/2026 (426,26 ha na semana de 26 a 31/10).
SOMA_DIAS = {"Uniport Próprio - Terceira Molécula": ("Uniport Próprio - Terceira Molécula (Bordadura)", "Uniport Próprio"),
             "Uniport Pingente Inseticida": ("Uniport Pingente Inseticida", "Uniport Pingente-Terceiro")}
# hectares realizados que a planilha não traz, lançados a pedido: (atividade, dia) -> ha a mais no dia e no mês
LANCAMENTOS = {("Drone Inseticida", "2026-10-30"): 1316.31}   # pedido em 03/11/2026


def marca_nao_planejada(a):
    if a["nome"] not in NAO_PLANEJADAS:
        return
    a["naoPlanejada"] = True
    a["equip"] = a["metaUnid"] = a["pessoasPlan"] = a["metaPessoa"] = None
    a["metaDia"] = a["metaSemana"] = a["metaMes"] = 0.0
    a.pop("diasSemMeta", None)


def ajusta_dias_sem_meta(a, dias, dia_ref):
    """Meta da semana e do mês só com os dias que têm meta (a planilha conta todos os
    dias menos o domingo). dias: os dias da semana (iso ou data); dia_ref: data."""
    import calendar
    fora = SEM_META_DIAS.get(a["nome"], [])
    datas = SEM_META_DATAS.get(a["nome"], [])
    if not (fora or datas) or not a.get("metaDia"):
        return
    if fora:
        a["diasSemMeta"] = [(w + 1) % 7 for w in fora]   # no navegador: getDay (0 = domingo)
    if datas:
        a["datasSemMeta"] = sorted(datas)
    ref = dia_ref.date() if isinstance(dia_ref, datetime.datetime) else dia_ref
    ds = [d.date() if isinstance(d, datetime.datetime) else
          (datetime.date.fromisoformat(d) if isinstance(d, str) else d) for d in dias]
    vale = lambda d: d.weekday() != 6 and d.weekday() not in fora and d.isoformat() not in datas
    a["metaSemana"] = r2(a["metaDia"] * sum(1 for d in ds if d <= ref and vale(d)))
    n_mes = sum(1 for k in range(1, calendar.monthrange(ref.year, ref.month)[1] + 1)
                if vale(datetime.date(ref.year, ref.month, k)))
    a["metaMes"] = r2(a["metaDia"] * n_mes)
MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def s(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def num(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    try:
        return float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        return None


def data_xl(v):
    if isinstance(v, datetime.datetime):
        return v
    if isinstance(v, (int, float)) and 40000 < v < 60000:
        return EPOCA + datetime.timedelta(days=float(v))
    return None


def iso(d):
    return d.strftime("%Y-%m-%d")


def r2(v):
    return None if v is None else round(v, 2)


# ------------------------------------------------------------------ leitura
def le_xlsb(path):
    from pyxlsb import open_workbook
    out = {}
    with open_workbook(path) as wb:
        for nome in wb.sheets:
            with wb.get_sheet(nome) as sh:
                out[nome] = [[c.v for c in r] for r in sh.rows()]
    return out


def le_xlsx(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    return {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}


def le(path):
    return le_xlsb(path) if path.lower().endswith(".xlsb") else le_xlsx(path)


def cel(linha, j):
    return linha[j] if j < len(linha) else None


# ------------------------------------------------------- semana (xlsb)
def semana(ws):
    """Aba Semanal-Acumulado: atividades, metas e realizado por dia."""
    # cabeçalho com as datas (linha 'Atividade') e o quadro de resumo à direita
    hdr_i = next(i for i, r in enumerate(ws) if s(cel(r, 2)) == "Atividade")
    hdr = ws[hdr_i]
    colData = {j: data_xl(v) for j, v in enumerate(hdr) if data_xl(v)}
    rot = {s(v): j for j, v in enumerate(hdr) if isinstance(v, str)}
    C = {
        "equip": rot.get("N° Equipamentos/Pessoas"), "metaUnid": rot.get("Meta - Equipamentos/Pessoas"),
        "metaDia": rot.get("Meta Dia"), "metaSemana": rot.get("Meta Semana"), "metaMes": rot.get("Meta Mês"),
    }
    # dia de referência e quantidade de dias da semana (quadro "DIA REF" / "QTD DIAS")
    dia_ref = qtd = None
    for r in ws[:hdr_i]:
        for j, v in enumerate(r):
            if s(v) == "DIA REF:":
                dia_ref = data_xl(cel(r, j + 1))
            if s(v) == "QTD DIAS":
                qtd = int(num(cel(r, j + 1)) or 0)
    if not dia_ref:
        sys.exit("Semanal-Acumulado: não achei o DIA REF")
    seg = dia_ref - datetime.timedelta(days=dia_ref.weekday())
    semana_toda = [seg + datetime.timedelta(days=k) for k in range(7)]
    col_de = {iso(d): j for j, d in colData.items()}

    def valores(r, ds):
        return [num(cel(r, col_de.get(iso(d), -1))) if iso(d) in col_de else None for d in ds]

    frentes, atual, ativs = [], None, []
    linhas_pess = []      # (índice da atividade dona, pessoas[], rendimento[])
    for i in range(hdr_i + 1, len(ws)):
        r = ws[i]
        nome = s(cel(r, 2))
        if not nome or nome == "Atividade" or re.match(r"^\d{2}/\d{2}/\d{4} a", nome) or nome.startswith("Observação"):
            continue
        fr = next((f for f in FRENTES if f[0] in nome.upper()), None)
        if fr and nome.upper() == nome:                      # cabeçalho de frente
            atual = {"id": fr[1], "nome": fr[2], "titulo": nome, "atividades": []}
            frentes.append(atual)
            continue
        if atual is None:
            continue
        if nome in ("Pessoas Dia", "Rendimento H/Dia"):
            if nome == "Pessoas Dia":
                rend = ws[i + 1] if i + 1 < len(ws) and s(cel(ws[i + 1], 2)).startswith("Rendimento") else None
                linhas_pess.append((atual, r, rend))
            continue
        a = {
            "nome": nome, "linha": i + 1,
            "equip": num(cel(r, C["equip"])) if C["equip"] is not None else None,
            "metaUnid": num(cel(r, C["metaUnid"])) if C["metaUnid"] is not None else None,
            "metaDia": num(cel(r, C["metaDia"])) or 0.0,
            "metaSemana": num(cel(r, C["metaSemana"])) or 0.0,
            "metaMes": num(cel(r, C["metaMes"])) or 0.0,
            "dias": valores(r, semana_toda),
            "ant": valores(r, [d - datetime.timedelta(days=7) for d in semana_toda]), "pessoas": None,
            "pessoasPlan": None, "metaPessoa": None,
            "_r": r,
        }
        if nome in META_EQUIP:
            a["metaUnid"] = META_EQUIP[nome]
        if nome in EQUIP and a["metaDia"] and a["metaUnid"]:
            fator = EQUIP[nome] * a["metaUnid"] / a["metaDia"]
            a["equip"] = float(EQUIP[nome])
            for k in ("metaDia", "metaSemana", "metaMes"):
                a[k] *= fator
        atual["atividades"].append(a)

    # os valores vêm como a planilha soma: a semana vai de segunda ao DIA REF e leva também
    # os dias seguintes da semana que tiveram produção (o domingo trabalhado entra na
    # semana e no mês, sem meta: a meta da semana segue a da planilha)
    fim = max(k for k, d in enumerate(semana_toda)
              if d <= dia_ref or any((a["dias"][k] or 0) > 0 for fr in frentes if fr["id"] not in FRENTES_FORA
                                     for a in fr["atividades"]))
    dias = semana_toda[:fim + 1]
    dias_ant = [d - datetime.timedelta(days=7) for d in dias]
    for fr in frentes:
        for a in fr["atividades"]:
            a["dias"], a["ant"] = a["dias"][:fim + 1], a["ant"][:fim + 1]

    # pessoas/dia: a linha pertence à atividade da mesma frente cujo realizado
    # dividido pelas pessoas bate com a linha de rendimento logo abaixo
    avisos = []
    for fr, rp, rr in linhas_pess:
        pess = valores(rp, dias)
        rend = valores(rr, dias) if rr else [None] * len(dias)
        melhor, erro_min, invertido = None, None, False
        for a in fr["atividades"]:
            e1, e2 = [], []
            for v, p, rd in zip(a["dias"], pess, rend):
                if v and p and rd:
                    e1.append(abs(v / p - rd))          # ha por pessoa (o certo)
                    e2.append(abs(p / v - rd))          # pessoas por ha (fórmula invertida)
            if e1:
                e = min(sum(e1), sum(e2)) / len(e1)
                if erro_min is None or e < erro_min:
                    melhor, erro_min, invertido = a, e, sum(e2) < sum(e1)
        if melhor is None:      # sem rendimento para conferir: a atividade com valor mais próxima acima
            melhor = next((a for a in reversed(fr["atividades"]) if any(a["dias"])), None)
        if melhor is not None:
            melhor["pessoas"] = pess
            # planejado de pessoas: a linha "Pessoas Dia" pode trazer o próprio quadro de
            # meta (Cata Mato: 15 pessoas); sem ele, vale o da atividade. META_PESSOA
            # tem a palavra final sobre a meta por pessoa
            pp, pm = num(cel(rp, C["equip"])), num(cel(rp, C["metaUnid"]))
            melhor["pessoasPlan"] = r2(pp) if pp else melhor["equip"]
            melhor["metaPessoa"] = r2(pm) if pm else melhor["metaUnid"]
            if melhor["nome"] in META_PESSOA:
                melhor["metaPessoa"] = META_PESSOA[melhor["nome"]]
            if melhor["nome"] in EQUIP:
                melhor["pessoasPlan"] = float(EQUIP[melhor["nome"]])
            if invertido:
                avisos.append(f"{melhor['nome']}: o \"Rendimento H/Dia\" da planilha está como pessoas ÷ ha; "
                              f"na apresentação o rendimento é ha ÷ pessoa.")

    # dias em que uma atividade repete a linha de pessoas ou de rendimento: é linha
    # colada fora do lugar na planilha. Esses valores saem da conta e viram aviso.
    refs = []
    for _, rp, rr in linhas_pess:
        refs.append(rp)
        if rr:
            refs.append(rr)
    meses = sorted({(d.year, d.month) for d in colData.values() if d <= dias[-1]})
    for fr in frentes:
        for a in fr["atividades"]:
            r = a.pop("_r")
            ruins = set()
            for ref in refs:
                if ref is r:
                    continue
                seq = []
                for j in sorted(colData):
                    v, w = num(cel(r, j)), num(cel(ref, j))
                    if v and w and abs(v - w) < 1e-9:
                        seq.append(j)
                    else:
                        if len(seq) >= 3:
                            ruins.update(seq)
                        seq = []
                if len(seq) >= 3:
                    ruins.update(seq)
            if ruins:
                ds = sorted(colData[j] for j in ruins)
                avisos.append(f"{a['nome']}: de {ds[0]:%d/%m} a {ds[-1]:%d/%m} a linha repete pessoas/rendimento "
                              f"de outra atividade ({len(ds)} dias) — valores fora da conta, conferir a planilha.")
                for k, d in enumerate(dias):
                    if col_de.get(iso(d)) in ruins:
                        a["dias"][k] = None
                for k, d in enumerate(dias_ant):
                    if col_de.get(iso(d)) in ruins:
                        a["ant"][k] = None
            mes = {}
            for (y, m) in meses:
                tot = sum(num(cel(r, j)) or 0 for j, d in colData.items()
                          if d.year == y and d.month == m and d <= dias[-1] and j not in ruins)
                mes[f"{y}-{m:02d}"] = round(tot, 2)
            a["mes"] = mes
            a["dias"] = [r2(v) for v in a["dias"]]
            a["ant"] = [r2(v) for v in a["ant"]]
            if a["pessoas"]:
                a["pessoas"] = [r2(v) for v in a["pessoas"]]
            for k in ("equip", "metaUnid", "metaDia", "metaSemana", "metaMes"):
                a[k] = r2(a[k])
            ajusta_dias_sem_meta(a, dias, dia_ref)
            marca_nao_planejada(a)
    frentes = separa_frentes([fr for fr in frentes if fr["id"] not in FRENTES_FORA])
    # atividade sem meta e sem nada realizado no ano não entra
    for fr in frentes:
        fr["atividades"] = [a for a in fr["atividades"] if a["metaSemana"] or any(a["mes"].values())]
    return {
        "diaRef": iso(dia_ref), "qtdDias": qtd or len(dias),
        "dias": [iso(d) for d in dias], "diasAnt": [iso(d) for d in dias_ant],
        "meses": [f"{y}-{m:02d}" for (y, m) in meses],
    }, frentes, avisos


# ---------------------------------------------------- programas (xlsx)
def tabela_fazendas(ws, cab_req):
    """Linhas de fazenda entre o cabeçalho (que tem as colunas pedidas) e o 'Total'."""
    for i, r in enumerate(ws):
        rot = {s(v): j for j, v in enumerate(r) if isinstance(v, str)}
        if all(c in rot for c in cab_req):
            lin = []
            for r2_ in ws[i + 1:]:
                if s(cel(r2_, 0)) in ("Total", "TOTAL", "TOTAL GERAL"):
                    break
                cod = num(cel(r2_, rot["Cód Faz"]))
                if cod and 9000 <= cod < 10000:
                    lin.append(r2_)
            return rot, lin, i
    return None, [], None


# o mesmo bloco vem escrito de jeitos diferentes em cada aba ("SHEKINHA-I",
# "SHEKINAH l", "MARCO-AMUY", "BR-154", "SCALON 7 PLACAS"...): um nome só para todos
SINONIMOS_BLOCO = {"SHEKINHA I": "SHEKINAH I", "SHEKINAH L": "SHEKINAH I", "BR 154": "154",
                   "SCALON 7 PLACAS": "SCALON SETE PLACAS"}


def bloco_norm(b):
    b = re.sub(r"\s+", " ", s(b).upper().replace("-", " ")).strip()
    return SINONIMOS_BLOCO.get(b, b)


def junta(dic, k, v):
    if v:
        dic[k] = round(dic.get(k, 0) + v, 2)


def quadro_bloco(ws, rotulo="BLOCO"):
    """Quadro 'BLOCO | Á REALIZAR' (ou REGIÃO) que a equipe monta na própria aba."""
    for i, r in enumerate(ws):
        if s(cel(r, 0)) == rotulo and s(cel(r, 1)).upper().startswith("Á REALIZAR"):
            out = {}
            for r2_ in ws[i + 1:]:
                k, v = s(cel(r2_, 0)), num(cel(r2_, 1))
                if not k or v is None:
                    break
                junta(out, bloco_norm(k) if rotulo == "BLOCO" else k, v)
            return out
    return {}


def soma(lin, j):
    return round(sum(num(cel(r, j)) or 0 for r in lin), 2)


def por_bloco(lin, jb, jv):
    out = {}
    for r in lin:
        b, v = s(cel(r, jb)), num(cel(r, jv)) or 0
        if b and not b.startswith("#"):
            junta(out, bloco_norm(b), v)
    return out


def programas(wb):
    P = []
    # pré-emergente cana soca: base = área colhida; saldo pelo quadro BLOCO da aba
    ws = wb["Pré-Emerg -Cana Soca SAFRA"]
    rot, lin, _ = tabela_fazendas(ws, ["Cód Faz", "Bloco", "Área Colhida", "Área Realizada", "Á Realizar"])
    blocos = quadro_bloco(ws)
    P.append({"id": "pre", "nome": "Pré-emergente cana soca", "base": soma(lin, rot["Área Colhida"]), "baseRot": "área colhida",
              "reforma": soma(lin, rot["Área Reforma"]), "realizado": soma(lin, rot["Área Realizada"]),
              "aRealizar": round(sum(blocos.values()), 2), "blocos": blocos, "regioes": quadro_bloco(ws, "REGIÃO")})
    # pré-emergente área de queima
    ws = wb["Pré-Emerg - Área Queima"]
    rot, lin, _ = tabela_fazendas(ws, ["Cód Faz", "Bloco", "Área Queimada", "Área Realizada", "Á Realizar"])
    P.append({"id": "queima", "nome": "Pré-emergente área de queima", "base": soma(lin, rot["Área Queimada"]), "baseRot": "área queimada",
              "realizado": soma(lin, rot["Área Realizada"]), "aRealizar": soma(lin, rot["Á Realizar"]),
              "blocos": por_bloco(lin, rot["Bloco"], rot["Á Realizar"])})
    # 3ª molécula (reforço de bordadura)
    ws = wb["Terceira Mol. - Soca SAFRA"]
    rot, lin, _ = tabela_fazendas(ws, ["Cód Faz", "Bloco", "Área Aberta", "Área Realizada", "Á Realizar"])
    P.append({"id": "mol3", "nome": "3ª molécula (bordadura)", "base": soma(lin, rot["Área Aberta"]), "baseRot": "área aberta",
              "realizado": soma(lin, rot["Área Realizada"]), "aRealizar": soma(lin, rot["Á Realizar"]),
              "blocos": por_bloco(lin, rot["Bloco"], rot["Á Realizar"])})
    # 1ª catação química: saldo dividido em mais e menos de 20 dias após a colheita
    ws = wb["1ª Catação Química SAFRA"]
    rot, lin, _ = tabela_fazendas(ws, ["Cód Faz", "Bloco", "Área Aberta", "Realizado Total"])
    j20 = next(j for k, j in rot.items() if k.startswith("Á Realizar") and ">" in k)
    j20m = next(j for k, j in rot.items() if k.startswith("Á Realizar") and "<" in k)
    mais, menos = soma(lin, j20), soma(lin, j20m)
    bl = por_bloco(lin, rot["Bloco"], j20)
    for b, v in por_bloco(lin, rot["Bloco"], j20m).items():
        bl[b] = round(bl.get(b, 0) + v, 2)
    P.append({"id": "cat1", "nome": "1ª catação química", "base": soma(lin, rot["Área Aberta"]), "baseRot": "área aberta",
              "reforma": soma(lin, rot["Área Reforma"]), "realizado": soma(lin, rot["Realizado Total"]),
              "mec": soma(lin, rot["Área Realizada MEC"]), "man": soma(lin, rot["Área Realizada MAN"]),
              "aRealizar": round(mais + menos, 2), "atrasado": mais, "blocos": bl,
              "blocosAtrasado": por_bloco(lin, rot["Bloco"], j20)})
    # 2ª catação química
    ws = wb["2ª Catação Química SAFRA"]
    rot, lin, _ = tabela_fazendas(ws, ["Cód Faz", "Bloco", "Programado", "Área Realizada", "Á Realizar"])
    P.append({"id": "cat2", "nome": "2ª catação química", "base": soma(lin, rot["Programado"]), "baseRot": "programado",
              "reforma": soma(lin, rot["Área Reforma"]), "realizado": soma(lin, rot["Área Realizada"]),
              "aRealizar": soma(lin, rot["Á Realizar"]), "blocos": por_bloco(lin, rot["Bloco"], rot["Á Realizar"])})
    # adubação de soca: resumo por bloco da própria planilha (localizada e cocho)
    ws = wb["ABERTO-ENCERRADO_BLOCO"]
    for titulo, pid, nome in (("LOCALIZADA", "loc", "Adubação localizada"), ("COCHO", "cocho", "Adubação com cocho")):
        ini = next(i for i, r in enumerate(ws) if s(cel(r, 5)) == titulo)
        cab = next(i for i in range(ini, len(ws)) if s(cel(ws[i], 0)) == "BLOCO")
        bl, tot = {}, None
        for r in ws[cab + 1:]:
            b = s(cel(r, 0))
            if not b:
                break
            vals = [num(cel(r, j)) or 0 for j in range(1, 7)]   # aberta, encerrada, total, reforma, realizado, a realizar
            if b.startswith("TOTAL"):
                tot = vals
                break
            junta(bl, bloco_norm(b), vals[5])
        P.append({"id": pid, "nome": nome, "base": round(tot[2], 2), "baseRot": "área com ordem",
                  "reforma": round(tot[3], 2), "realizado": round(tot[4], 2), "aRealizar": round(tot[5], 2),
                  "blocos": {b: v for b, v in bl.items() if v}})
    return P


# ------------------------------------------ fazendas e talhões da semana
# Detalhe por fazenda e talhão, sempre só do que foi feito nos dias da semana:
# - relatório de operações do sistema (RSC0819R, .xlsx) que estiver em dados/: todas
#   as operações, com OS, área executada e dose de cada produto;
# - aba "Realizado - 3ª Mol" da planilha SAFRA: OS da 3ª molécula (Uniport
#   bordadura), com a área de borda aplicada em cada talhão (coluna "Recom.").
# A operação do sistema vai para a frente pela descrição; a atividade só quando não
# há dúvida (o sistema não separa Uniport próprio de terceiro nem as equipes do cocho).
OPER_FRENTE = [("DRONE", "aerea", "Drone Inseticida"), ("UNIPORT", "uniport", None),
               ("COCHO", "cocho", None), ("FORMIGA", "manual", "Catação de Formiga"),
               ("MANUAL", "manual", None)]
# rótulo curto das operações mais comuns (o relatório traz a descrição longa e truncada)
ROTULOS = {
    24565: "Broca · drone · 1ª aplic.", 24566: "Broca · drone · 2ª aplic.",
    24593: "Cigarrinha · drone · 1ª aplic.", 14501: "Formiga · manual · 1ª aplic.",
    24003: "Adubação cocho · 1ª aplic.", 24004: "Adubação cocho · 2ª aplic.",
    14330: "Herb. Uniport barra total · 1ª aplic.", 24330: "Herb. Uniport barra total · 1ª aplic.",
    24331: "Herb. Uniport barra total · 2ª aplic.", 14327: "Herb. Uniport bordadura · 1ª aplic.",
    24327: "Herb. Uniport bordadura · 1ª aplic.", 14340: "Herb. manual · 1ª aplic.",
    14341: "Herb. manual · 2ª aplic.", 14342: "Herb. manual · 3ª aplic.",
}
UNIDADES = {"LT": "L", "KG": "kg", "TON": "t", "UN": "un", "HA": "ha"}
MINUSCULAS = {"de", "do", "da", "dos", "das", "e", "ou", "d'"}


def nome_faz(fundo):
    """'9006 - FAZ. CÓRREGO DAS FLORES II' -> 'Faz. Córrego das Flores II'"""
    t = s(fundo)
    t = t.split(" - ", 1)[-1] if " - " in t else t
    t = re.sub(r"^FAZ\.(?=\S)", "FAZ. ", t, flags=re.I)
    out = []
    for i, w in enumerate(t.split(" ")):
        lw = w.lower()
        if re.fullmatch(r"l{1,3}", w) and i > 0:         # "SUCURI ll" = Sucuri II
            out.append(w.upper().replace("L", "I"))
        elif re.fullmatch(r"[ivx]+", lw) and i > 0:
            out.append(w.upper())
        elif lw in MINUSCULAS and i > 0:
            out.append(lw)
        else:
            out.append(re.sub(r"^(\(?)(\w)", lambda m: m.group(1) + m.group(2).upper(), lw))
    return " ".join(out)


def talhao(v):
    n = num(v)
    return str(int(n)) if n is not None and n == int(n) else s(v)


def oper_frente(desc):
    d = desc.upper()
    for chave, fr, at in OPER_FRENTE:
        if chave in d:
            return fr, at
    return None, None


def rsc_semana(ws, dias):
    """Relatório de operações (RSC0819R): uma linha por OS × talhão, com os produtos."""
    out, nomes = [], {}
    faz = tal = ha_tal = op = cur = None
    for r in ws[2:]:
        c0 = s(cel(r, 0))
        if c0 == "PROPRIEDADE:":
            faz = int(num(cel(r, 1)))
            nomes[str(faz)] = nome_faz(cel(r, 5))
        elif c0 == "TALHÃO:":
            tal, ha_tal = talhao(cel(r, 1)), num(cel(r, 4))
        elif c0 == "Operação:":
            cod, desc = int(num(cel(r, 1))), s(cel(r, 3))
            op = (cod, desc, ROTULOS.get(cod) or re.sub(r"\(HA\)$", "", desc).strip(" (").capitalize())
        elif c0 == "OS" or cel(r, 0) is None:
            continue
        else:
            if c0:                                   # linha nova de OS; as seguintes sem OS são outros produtos dela
                d, cur = data_xl(cel(r, 2)), None
                fr, at = oper_frente(op[1])
                if d and iso(d) in dias and fr:
                    cur = {"f": fr, "a": at, "d": iso(d), "faz": faz, "t": tal, "ht": r2(ha_tal),
                           "ha": r2(num(cel(r, 5))), "os": int(num(c0)), "op": op[2], "prod": []}
                    out.append(cur)
            if cur is not None and s(cel(r, 7)):
                un = UNIDADES.get(s(cel(r, 8)), s(cel(r, 8)))
                dose, rec = num(cel(r, 9)), num(cel(r, 11))
                cur["prod"].append([s(cel(r, 7)), None if dose is None else round(dose, 4),
                                    None if rec is None else round(rec, 4), un])
    return out, nomes


def mol3_semana(wb, dias):
    """OS da 3ª molécula encerradas nos dias da semana (aba Realizado - 3ª Mol)."""
    nome = next((k for k in wb if k.startswith("Realizado - 3")), None)
    if not nome:
        return [], {}
    ws = wb[nome]
    rot = {s(v): j for j, v in enumerate(ws[0]) if isinstance(v, str)}
    out, nomes = [], {}
    for r in ws[1:]:
        d = data_xl(cel(r, rot["Data Op."]))
        faz = num(cel(r, rot["Propriedade"]))
        if not d or iso(d) not in dias or not faz:
            continue
        faz = int(faz)
        nomes[str(faz)] = nome_faz(cel(r, rot["Fundo Agrícola"]))
        out.append({"f": "uniport", "a": None, "d": iso(d), "faz": faz, "t": talhao(cel(r, rot["Talhão"])),
                    "ht": r2(num(cel(r, rot["Plantada"]))), "ha": r2(num(cel(r, rot["Recom."]))),
                    "os": int(num(cel(r, rot["OS"]))), "op": "3ª molécula · Uniport bordadura",
                    "corte": int(num(cel(r, rot["Corte"])) or 0) or None, "prod": []})
    return out, nomes


def locais_semana(wb_prog, rscs, dias):
    linhas, nomes, fontes = [], {}, []
    for arq, ws in rscs:
        lin, nm = rsc_semana(ws, dias)
        if lin:
            linhas += lin
            nomes.update(nm)
            fontes.append(arq)
    # a 3ª molécula entra quando a OS não veio no relatório do sistema
    ja = {(l["os"], l["faz"], l["t"]) for l in linhas}
    lin, nm = mol3_semana(wb_prog, dias)
    lin = [l for l in lin if (l["os"], l["faz"], l["t"]) not in ja]
    if lin:
        linhas += lin
        nomes.update(nm)
        fontes.append("Realizado - 3ª Mol")
    linhas.sort(key=lambda l: (l["f"], l["d"], l["faz"], num(l["t"]) or 0))
    return {"fazendas": nomes, "linhas": linhas, "fontes": fontes}


def data_atualizacao(wb):
    for r in wb["Pré-Emerg -Cana Soca SAFRA"][:6]:
        for j, v in enumerate(r):
            if s(v).startswith("DATA ATUALIZA"):
                d = data_xl(cel(r, j + 1))
                if d:
                    return iso(d)
    return None


# ------------------------------------------------------------------ main
def main():
    arquivos = sorted(glob.glob(os.path.join(AQUI, "dados", "*.xls*")))
    wb_sem = wb_prog = None
    rscs = []
    for a in arquivos:
        if os.path.basename(a).startswith("~$"):
            continue
        try:
            wb = le(a)
        except Exception as e:                       # arquivo de outro relatório na pasta
            print("ignorado:", os.path.basename(a), "-", e)
            continue
        if "Semanal-Acumulado" in wb:
            wb_sem, arq_sem = wb, os.path.basename(a)
        if "Pré-Emerg -Cana Soca SAFRA" in wb:
            wb_prog, arq_prog = wb, os.path.basename(a)
        primeira = next(iter(wb.values()), [])
        if primeira and s(cel(primeira[0], 3)).startswith("Período Acumulado"):   # RSC0819R do sistema
            rscs.append((os.path.basename(a), primeira))
    if not wb_sem or not wb_prog:
        sys.exit("Faltou planilha em dados/: preciso da que tem a aba Semanal-Acumulado (.xlsb) "
                 "e da que tem as abas dos programas (SAFRA 2026 .xlsx)")
    sem, frentes, avisos = semana(wb_sem["Semanal-Acumulado"])
    dados = {
        "fonte": {"semana": arq_sem, "programas": arq_prog, "atualizacaoProgramas": data_atualizacao(wb_prog)},
        "semana": sem, "frentes": frentes, "avisos": avisos, "programas": programas(wb_prog),
        "nomesMes": MESES, "locais": locais_semana(wb_prog, rscs, set(sem["dias"])),
    }
    pdf = pdf_dos_programas()
    if pdf:
        dados["tabelas"] = tabelas_do_pdf(pdf)
        dados["fonte"]["tabelas"] = os.path.basename(pdf)
    grava(dados)
    n_at = sum(len(f["atividades"]) for f in frentes)
    print(f"semana {sem['dias'][0]} a {sem['dias'][-1]} ({sem['qtdDias']} dias) · {len(frentes)} frentes · "
          f"{n_at} atividades · {len(dados['programas'])} programas · {len(avisos)} aviso(s)")
    loc = dados["locais"]
    print(f"fazendas e talhões da semana: {len(loc['linhas'])} linha(s) de "
          f"{len({l['faz'] for l in loc['linhas']})} fazenda(s) · fontes: {', '.join(loc['fontes']) or 'nenhuma'}")
    for a in avisos:
        print("  aviso:", a)


# ------------------------------------------- tabelas dos programas (PDF)
# As páginas de cada programa (pré-emergência, reforço bordadura, 1ª catação, adubação
# localizada e com cocho) e os indicadores do fim saem da impressão das planilhas em
# PDF, como ela vem toda semana (programas_pdf.py). Fica em dados/; vale o mais novo.
def pdf_dos_programas():
    pdfs = sorted(glob.glob(os.path.join(AQUI, "dados", "*.pdf")), key=os.path.getmtime, reverse=True)
    return pdfs[0] if pdfs else None


# Áreas Cargill à parte (pedido em 20/10/2026): fora dos totais da localizada e do total
# geral, com linha e gráfico próprios. Os totais da planilha ("TOTAL REGIÃO 1 + CARGILL"
# e "TOTAL LOCALIZADA") trazem a Cargill em todas as colunas; só o saldo a realizar a
# planilha já soma sem ela (confere com a soma dos outros blocos) — nesse caso ele fica
# como está, senão a Cargill sai dele também.
def cargill_a_parte(tab):
    for p in tab["programas"]:
        if p["tipo"] != "adub":
            continue
        for sec in p["secoes"]:
            cg = next((l for l in sec["linhas"] if l["t"] == "sub" and "CARGILL" in l["rot"].upper()), None)
            if not cg:
                continue
            outros = sum(l["v"].get("aRealizar", 0) or 0 for l in sec["linhas"] if l["t"] == "sub" and l is not cg)
            novas = []
            for l in sec["linhas"]:
                rot = l.get("rot", "").upper()
                if l["t"] == "tot" and ("CARGILL" in rot or rot.endswith("LOCALIZADA")):
                    v = dict(l["v"])
                    for k, x in cg["v"].items():
                        if not x:
                            continue
                        if k == "aRealizar" and abs((v.get(k) or 0) - outros) <= 0.05:
                            continue
                        v[k] = round((v.get(k) or 0) - x, 2)
                    l = dict(l, v=v, rot=l["rot"].replace(" + CARGILL", "") + " (SEM CARGILL)")
                novas.append(l)
            novas.append({"t": "aparte", "rot": "Áreas Cargill · à parte, fora dos totais", "v": dict(cg["v"])})
            sec["linhas"] = novas
            p["extras"]["cargill"] = dict(cg["v"])


# A localizada parou em 27/10/2026 e fica só como histórico da própria (pedido do
# usuário: "área histórico da localizada somente essa 6.715", a do resumo da planilha).
# Nos quadros do cocho (Alencar e Fábio do Prado) a coluna da localizada sai zerada; a
# planilha também não a desconta do saldo a realizar deles nem a soma no resumo.
def localizada_so_propria(tab):
    import programas_pdf
    for p in tab["programas"]:
        if p["tipo"] != "adub":
            continue
        tirou = []
        for sec in p["secoes"]:
            if "COCHO" not in programas_pdf.sem_acento(sec["titulo"]):
                continue
            tot = [l for l in sec["linhas"] if l["t"] == "tot" and l["v"].get("localizada")]
            if tot:
                tirou.append(f"{sec['titulo']} {tot[-1]['v']['localizada']}")
            for l in sec["linhas"]:
                if l["v"].get("localizada"):
                    l["v"]["localizada"] = 0.0
        if tirou:
            print("  localizada fora da própria, zerada (só a da própria fica como histórico):", "; ".join(tirou))


# Saldo a realizar negativo (aplicado mais que a área: 9350 Itaúba, −0,30 ha em
# 27/10/2026) vale zero (pedido do usuário): a sobra num talhão não abate o que falta
# nos outros. O subtotal do bloco e os totais que o somam (o da região e o da equipe)
# sobem o mesmo tanto.
def saldo_sem_negativo(tab):
    import programas_pdf
    for p in tab["programas"]:
        if p["tipo"] != "adub":
            continue
        da_equipe = {}      # equipe -> o que já voltou ao saldo nos quadros anteriores dela
        for sec in p["secoes"]:
            equipe = re.sub(r"\s*-\s*REGIAO.*$", "", programas_pdf.sem_acento(sec["titulo"])).split(" - ")[-1]
            pend = acum = 0.0
            for l in sec["linhas"]:
                v = l["v"]
                x = v.get("aRealizar") or 0.0
                if l["t"] == "faz" and x < 0:
                    print(f"  saldo negativo zerado: {l['cod']} {l['prop']} {x}")
                    pend -= x
                    v["aRealizar"] = 0.0
                elif l["t"] == "sub":
                    x += pend
                    acum += pend + max(-x, 0.0)
                    v["aRealizar"] = r2(max(x, 0.0))
                    pend = 0.0
                elif l["t"] == "tot":
                    mais = acum + (da_equipe.get(equipe, 0.0) if equipe in programas_pdf.sem_acento(l["rot"]) else 0.0)
                    if mais:
                        v["aRealizar"] = r2(x + mais)
            da_equipe[equipe] = da_equipe.get(equipe, 0.0) + acum


# Pré-emergência: o que a planilha deixa a realizar na Cargill vai para a área de
# fornecedor (pedido em 27/10/2026). A área colhida sai junto, porque na planilha a área
# de fornecedor fica fora da colhida (a realizar = colhida − realizada − reforma;
# encerradas = colhida + fornecedor). O total geral acompanha.
def pre_cargill_fornecedor(tab):
    def passa(v, s):
        v["fornecedor"] = r2((v.get("fornecedor") or 0) + s)
        v["colhida"] = r2((v.get("colhida") or 0) - s)
        v["aRealizar"] = r2((v.get("aRealizar") or 0) - s)
    for p in tab["programas"]:
        if p["id"] != "pre":
            continue
        for sec in p["secoes"]:
            fazs, passou = [], 0.0
            for l in sec["linhas"]:
                if l["t"] == "faz":
                    fazs.append(l)
                    continue
                if l["t"] == "sub" and "CARGILL" in l["rot"].upper():
                    for f in fazs:
                        if f["v"].get("aRealizar"):
                            passa(f["v"], f["v"]["aRealizar"])
                    s = l["v"].get("aRealizar") or 0
                    if s:
                        passa(l["v"], s)
                        passou += s
                        print(f"  pré-emergência: a realizar da Cargill ({s} ha) passou para fornecedor")
                elif l["t"] == "tot" and passou:
                    passa(l["v"], passou)
                fazs = []


# fazendas tiradas de um programa a pedido (programa -> códigos): a linha sai e os números dela saem também
# do subtotal do bloco e dos totais do quadro. Pré-emergência, 03/11/2026: 9328 Bugre, 9376 Dos Baús Olhos d'Água e
# 9443 Pântano e Mariano (15 a 33 dias de atraso)
RETIRAR = {"pre": {"9328", "9376", "9443"}}


def retira_fazendas(tab):
    for p in tab["programas"]:
        cods = RETIRAR.get(p["id"])
        if not cods:
            continue
        for sec in p["secoes"]:
            fora, novas = [], []
            for l in sec["linhas"]:
                if l["t"] == "faz" and l["cod"] in cods:
                    fora.append(l)
                    print(f"  {p['id']}: tirada a fazenda {l['cod']} {l['prop']}")
                    continue
                if l["t"] in ("sub", "tot") and fora:
                    tira = fora if l["t"] == "tot" else [f for f in fora if not f.get("_no_sub")]
                    for f in tira:
                        for k, x in f["v"].items():
                            if x:
                                l["v"][k] = round((l["v"].get(k) or 0) - x, 2)
                    if l["t"] == "sub":
                        for f in fora:
                            f["_no_sub"] = True
                novas.append(l)
            sec["linhas"] = novas


# Erros de fórmula da planilha corrigidos (pedido de 10/11/2026). Fica como a planilha o que só parece não fechar:
# os subtotais e totais somam também as ordens encerradas, que a impressão não lista, então "a tratar − aplicada"
# no total não é o "a realizar" (este é a soma das fazendas abertas, e essa soma bate)
def corrige_formulas(tab):
    for p in tab["programas"]:
        for sec in p["secoes"]:
            linhas = sec["linhas"]
            # 1ª catação: fazenda com exatamente 20 dias de atraso entra nas duas colunas (> 20 e < 20 dias) com a
            # mesma área; fica só em "< 20 dias" (20 ainda não passou de 20) e sai do subtotal/total
            if p["tipo"] == "cat1":
                tira = 0.0
                for l in linhas:
                    v = l["v"]
                    if l["t"] == "faz" and (v.get("maior20") or 0) > 0 and (v.get("menor20") or 0) > 0:
                        conta = (v.get("aberta") or 0) - (v.get("reforma") or 0) - (v.get("realTotal") or 0)
                        if abs(v["maior20"] + v["menor20"] - conta) > 0.05 and abs(v["menor20"] - conta) <= 0.05:
                            print(f"  fórmula corrigida: 1ª catação {l['cod']} {l['prop']} ({l.get('atraso')} dias) "
                                  f"estava nas duas colunas; fica só em < 20 dias ({v['menor20']})")
                            tira += v["maior20"]
                            v["maior20"] = 0.0
                    elif l["t"] in ("sub", "tot") and tira:
                        v["maior20"] = r2((v.get("maior20") or 0) - tira)
                        if l["t"] == "sub":
                            tira = 0.0
            # pré-emergência: o total "a colher" tem de ser a soma das regiões (a planilha somou 275,77 a mais)
            if p["tipo"] == "pre":
                subs = [l for l in linhas if l["t"] == "sub"]
                for l in linhas:
                    if l["t"] == "tot" and subs:
                        soma = r2(sum(s["v"].get("aColher") or 0 for s in subs))
                        if abs((l["v"].get("aColher") or 0) - soma) > 0.05:
                            print(f"  fórmula corrigida: pré-emergência, total a colher {l['v']['aColher']} → {soma} (soma das regiões)")
                            l["v"]["aColher"] = soma


def tabelas_do_pdf(pdf):
    import programas_pdf
    tab = programas_pdf.ler(pdf)
    corrige_formulas(tab)
    cargill_a_parte(tab)
    retira_fazendas(tab)
    pre_cargill_fornecedor(tab)
    localizada_so_propria(tab)
    saldo_sem_negativo(tab)
    for p in tab["programas"]:
        n = sum(1 for sec in p["secoes"] for l in sec["linhas"] if l["t"] == "faz")
        print(f"  {p['nome']}: {n} fazenda(s) em {len(p['secoes'])} quadro(s)")
    return tab


def le_dados():
    html = open(HTML, encoding="utf-8").read()
    m = re.search(r"/\*DADOS-INICIO\*/\nwindow\.TRATOS = (.*?);\n/\*DADOS-FIM\*/", html, flags=re.S)
    return json.loads(m.group(1))


def grava(dados):
    js = "window.TRATOS = " + json.dumps(dados, ensure_ascii=False, separators=(",", ":"))
    html = open(HTML, encoding="utf-8").read()
    novo, n = re.subn(r"(/\*DADOS-INICIO\*/\n).*?(\n/\*DADOS-FIM\*/)",
                      lambda m: m.group(1) + js + ";" + m.group(2), html, flags=re.S)
    if n != 1:
        sys.exit("Marcadores DADOS-INICIO / DADOS-FIM não encontrados no index.html")
    open(HTML, "w", encoding="utf-8", newline="\n").write(novo)


def _linha_do_pdf(a, r, s, dias, dia_ref, ym, ym_ant, virou=False):
    """Números de uma atividade pela linha dela na página 1 do PDF: realizado por dia,
    pessoas, quadro de meta e realizado do mês. As correções de meta (META_EQUIP, EQUIP,
    META_PESSOA, SEM_META_DIAS, NAO_PLANEJADAS) valem de novo sobre os números da planilha.
    virou: a planilha já está no mês seguinte ao do DIA REF (atualizada em 03/11 com a semana
    até 31/10): o "mês" dela é o novo, o "mês anterior" é o do DIA REF e a meta do mês dela é a
    do mês novo; a do mês do DIA REF sai da meta do dia × dias de segunda a sábado.
    Devolve o que foi para o dia não impresso (domingo), ou 0."""
    novo = [r["dias"].get(d) for d in dias]
    # a impressão esconde o domingo trabalhado, mas a planilha soma: o que faltar
    # para fechar a semana dela vai para o dia que não foi impresso
    fora = [k for k, d in enumerate(dias) if d not in s["datas"]]
    falta = round((r.get("realSemana") or 0) - sum(v for v in novo if v), 2)
    domingo = 0.0
    if falta > 0.01:
        if len(fora) == 1:
            novo[fora[0]] = falta
            domingo = falta
        else:
            print(f"  aviso: {a['nome']}: {falta} ha da semana fora dos dias impressos")
    a["dias"] = [r2(v) for v in novo]
    a["pessoas"] = [r2(r["pessoas"].get(d)) for d in dias] if r["pessoas"] else None
    for k in ("equip", "metaUnid", "metaDia", "metaSemana", "metaMes"):
        a[k] = r.get(k) if k in ("equip", "metaUnid") else (r.get(k) or 0.0)
    nome = a["nome"]
    if nome in META_EQUIP:
        a["metaUnid"] = META_EQUIP[nome]
    if nome in EQUIP and a["metaDia"] and a["metaUnid"]:
        fator = EQUIP[nome] * a["metaUnid"] / a["metaDia"]
        a["equip"] = float(EQUIP[nome])
        for k in ("metaDia", "metaSemana", "metaMes"):
            a[k] *= fator
    for k in ("equip", "metaUnid", "metaDia", "metaSemana", "metaMes"):
        a[k] = r2(a[k])
    if virou:
        y, m = int(ym[:4]), int(ym[5:])
        uteis = sum(1 for k in range(1, calendar.monthrange(y, m)[1] + 1) if datetime.date(y, m, k).weekday() != 6)
        a["metaMes"] = r2(a["metaDia"] * uteis)
    ajusta_dias_sem_meta(a, dias, dia_ref)
    if a["pessoas"]:
        a["pessoasPlan"] = r2(r.get("pessoasPlan")) or a["equip"]
        a["metaPessoa"] = r2(r.get("metaPessoa")) or a["metaUnid"]
        if nome in META_PESSOA:
            a["metaPessoa"] = META_PESSOA[nome]
        if nome in EQUIP:
            a["pessoasPlan"] = float(EQUIP[nome])
    else:
        a["pessoasPlan"] = a["metaPessoa"] = None
    marca_nao_planejada(a)
    if virou:
        a["mes"][ym] = r2(r.get("mesAnt") or 0.0)
    else:
        a["mes"][ym] = r2(r.get("realMes") or 0.0)
        if r.get("mesAnt") is not None:
            a["mes"][ym_ant] = r2(r["mesAnt"])
    return domingo


def _frente_do_pdf(dados, r):
    """A operação da apresentação de uma linha do PDF (pela frente impressa e SEPARA)."""
    tit = (r.get("frente") or "").upper()
    fid = next((id_ for t, id_, _n in FRENTES if t in tit), None)
    if fid is None or fid in FRENTES_FORA:
        return None
    for orig, trecho, sep, _n, _t in SEPARA:
        if fid == orig and trecho in r["nome"].lower():
            fid = sep
    return next((fr for fr in dados["frentes"] if fr["id"] == fid), None)


def semana_do_pdf(dados, pdf):
    """Números da semana a partir da página 1 do PDF (Semanal-Acumulado impressa).

    Mesma semana do index.html: atualiza realizado por dia, pessoas, quadro de meta e
    realizado do mês de cada atividade, achada pelo nome. Semana seguinte: a semana que
    está na apresentação vira a semana anterior (comparação dia a dia) e a nova entra com
    os números do PDF; atividade nova com realizado na semana entra na operação dela.
    Histórico dos meses e as atividades que a impressão não traz ficam como estão."""
    import programas_pdf
    s = programas_pdf.ler_semana(pdf)
    # impressão de 10/11/2026: os dias impressos são da semana fechada (02 a 08/11), mas o DIA REF já é o dia seguinte
    # (09/11) e "Realizado Dia/Semana/Mês" já contam esse dia. A semana é a dos dias impressos: realizado da semana =
    # soma dos dias e o do mês sem o do DIA REF
    if s["datas"] and s["diaRef"] and s["diaRef"] > max(s["datas"]):
        fim = max(s["datas"])
        print(f"  DIA REF {s['diaRef']} depois dos dias impressos: a semana vai até {fim}, sem o realizado de {s['diaRef']}")
        for r in s["atividades"]:
            if r.get("realMes") is not None:
                r["realMes"] = r2(r["realMes"] - (r.get("realDia") or 0))
            r["realSemana"] = r2(sum(v for v in r["dias"].values() if v)) or None
            r["realDia"] = r["dias"].get(fim)
        s["diaRef"] = fim
    sem = dados["semana"]
    ref = datetime.date.fromisoformat(s["diaRef"])
    seg = ref - datetime.timedelta(days=ref.weekday())
    seg_ap = datetime.date.fromisoformat(sem["dias"][0])
    if seg < seg_ap:
        print(f"semana NÃO atualizada: o PDF é do DIA REF {s['diaRef']}, anterior à semana da apresentação "
              f"({sem['dias'][0]} a {sem['diaRef']})")
        return
    if seg > seg_ap:
        # semana nova: a de agora vira a anterior (se for a semana logo antes) e os dias
        # da nova começam vazios; o detalhe por fazenda e os avisos eram da semana velha
        dias = [(seg + datetime.timedelta(days=k)).isoformat() for k in range(7)]
        seguida = seg - seg_ap == datetime.timedelta(days=7)
        for fr in dados["frentes"]:
            for a in fr["atividades"]:
                velho = (a.get("dias") or [])[:7]
                a["ant"] = (velho + [None] * 7)[:7] if seguida else [None] * 7
                a["dias"], a["pessoas"] = [None] * 7, None
        sem["diasAnt"] = sem["dias"][:7] if seguida else [(seg - datetime.timedelta(days=7 - k)).isoformat() for k in range(7)]
        sem["dias"] = dias
        dados["avisos"] = []
        dados["locais"] = {"fazendas": [], "linhas": [], "fontes": []}
        print(f"semana nova: {dias[0]} a {s['diaRef']} (a de {sem['diasAnt'][0]} vira a semana anterior)")
    dias = sem["dias"]
    sem["diaRef"] = s["diaRef"]
    sem["qtdDias"] = len(s["datas"])
    ym = s["diaRef"][:7]
    y, m = int(ym[:4]), int(ym[5:])
    ym_ant = f"{y if m > 1 else y - 1}-{(m - 1 or 12):02d}"
    virou = bool(s["atualizacao"]) and s["atualizacao"][:7] > ym
    if virou:
        print(f"  a planilha já está em {s['atualizacao'][:7]}: o mês anterior dela é o realizado de {ym}")
    if ym not in sem["meses"]:
        sem["meses"].append(ym)
    for fr in dados["frentes"]:
        for a in fr["atividades"]:
            a["mes"].setdefault(ym, 0)
    # linhas que a planilha não soma: nome do relatório e realizado da semana pela soma dos dias
    for r in s["atividades"]:
        if r["nome"] in SOMA_DIAS:
            r["nome"] = SOMA_DIAS[r["nome"]][0]
            if not r.get("realSemana"):
                r["realSemana"] = r2(sum(v for v in r["dias"].values() if v)) or None
    por_nome = {r["nome"]: r for r in s["atividades"]}
    usados, domingos = set(), []
    for fr in dados["frentes"]:
        for a in fr["atividades"]:
            r = por_nome.get(a["nome"])
            if not r:
                continue
            usados.add(a["nome"])
            dom = _linha_do_pdf(a, r, s, dias, ref, ym, ym_ant, virou)
            if dom:
                domingos.append(f"{a['nome']} {dom}")
    # atividade que a apresentação ainda não tem: entra se a planilha conta o realizado
    # dela na semana; linha sem realizado da semana (só dias soltos, sem meta) fica fora
    for r in s["atividades"]:
        if r["nome"] in usados or r["nome"] == "Média":
            continue
        fr = _frente_do_pdf(dados, r)
        if fr is None:
            continue
        if not (r.get("realSemana") or 0) > 0:
            if any(v for v in r["dias"].values()):
                print(f"  aviso: {r['nome']}: tem dias no PDF, mas a planilha não conta na semana (sem realizado da semana); fica fora")
            continue
        a = {"nome": r["nome"], "linha": None, "equip": None, "metaUnid": None, "metaDia": 0.0, "metaSemana": 0.0,
             "metaMes": 0.0, "dias": [None] * len(dias), "ant": [None] * len(dias), "pessoas": None,
             "pessoasPlan": None, "metaPessoa": None, "mes": {mm: 0 for mm in sem["meses"]}}
        dom = _linha_do_pdf(a, r, s, dias, ref, ym, ym_ant, virou)
        if dom:
            domingos.append(f"{a['nome']} {dom}")
        abaixo = next((v[1] for v in SOMA_DIAS.values() if v[0] == r["nome"]), None)
        k = next((i for i, x in enumerate(fr["atividades"]) if x["nome"] == abaixo), None)
        fr["atividades"].insert(k + 1 if k is not None else len(fr["atividades"]), a)
        usados.add(r["nome"])
        print(f"  atividade nova: {r['nome']} em {fr['nome']}")
    # lançamentos a pedido: somam no dia (se for da semana) e no realizado do mês
    for fr in dados["frentes"]:
        for a in fr["atividades"]:
            for (nome, dia), ha in LANCAMENTOS.items():
                if a["nome"] != nome:
                    continue
                if dia not in dias:
                    # dia de uma semana já passada: o mês dele veio de novo do PDF (sem o lançamento), soma de novo
                    if a["nome"] in usados and dia[:7] in (ym, ym_ant):
                        a["mes"][dia[:7]] = r2((a["mes"].get(dia[:7]) or 0) + ha)
                        print(f"  lançado de novo no mês: {nome} {dia} +{ha} ha")
                    continue
                k = dias.index(dia)
                a["dias"][k] = r2((a["dias"][k] or 0) + ha)
                a["mes"][dia[:7]] = r2((a["mes"].get(dia[:7]) or 0) + ha)
                print(f"  lançado: {nome} {dia} +{ha} ha")
    if domingos:
        print("  domingo (dia não impresso) pela diferença entre a semana da planilha e os dias impressos:", "; ".join(domingos))
    dados["fonte"]["semana"] = os.path.basename(pdf)
    if s["atualizacao"]:
        dados["fonte"]["atualizacaoSemana"] = s["atualizacao"]
    print(f"semana {dias[0]} a {sem['diaRef']}: {len(usados)} atividade(s) atualizada(s) pelo PDF de {s['atualizacao']}")


def pelo_pdf(pdf):
    """Atualiza os números pelo PDF: tabelas dos programas e a semana (a mesma ou a
    seguinte, que vira a semana). Layout e regras não mudam. PDF com só uma parte (ex.:
    só as páginas da adubação) atualiza só os programas que traz; sem a página do
    resumo semanal, a semana fica como está."""
    import pymupdf
    import programas_pdf
    dados = le_dados()
    novas = tabelas_do_pdf(pdf)
    ids = {p["id"] for p in novas["programas"]}
    antes = dados.get("tabelas") or {"programas": []}
    ordem = [p["id"] for p in antes["programas"]] + [i for i in ids if i not in {p["id"] for p in antes["programas"]}]
    por_id = {p["id"]: p for p in antes["programas"]}
    por_id.update({p["id"]: p for p in novas["programas"]})
    dados["tabelas"] = dict(novas, programas=[por_id[i] for i in ordem])
    fonte = dados.setdefault("fonte", {})
    if ids >= set(por_id):
        fonte["tabelas"] = os.path.basename(pdf)
    else:
        parte = f"{', '.join(sorted(ids))} de {os.path.basename(pdf)}"
        fonte["tabelas"] = " · ".join([x for x in fonte.get("tabelas", "").split(" · ") if x and x != parte] + [parte])
    if novas["atualizacao"]:
        dados["fonte"]["atualizacaoProgramas"] = novas["atualizacao"]
    elif antes.get("atualizacao"):
        dados["tabelas"]["atualizacao"] = antes["atualizacao"]
    if any("RESUMO SEMANAL" in programas_pdf.sem_acento(pg.get_text()) for pg in pymupdf.open(pdf)):
        semana_do_pdf(dados, pdf)
    else:
        print("semana como está: o PDF não tem a página do resumo semanal")
    grava(dados)
    print("tabelas dos programas de", dados["tabelas"]["atualizacao"], "lidas de", os.path.basename(pdf))


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--pdf":
        pelo_pdf(sys.argv[2])
    else:
        main()
