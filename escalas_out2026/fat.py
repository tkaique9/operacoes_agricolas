"""FAT: motoristas e operadores com direito, as férias de cada um e a fila de quem vai.

python fat.py  ->  fat.json                    (sem nomes: a fila com papel, atividade e mês do limite das férias,
                                                para a página FAT do estudo)
               ->  FAT_lista_nominal.html      (com nome e matrícula: fica só na máquina, no .gitignore)
               ->  build/fat_nominal_artifact.html  (a mesma lista no formato da página privada do claude.ai)

Fonte: o painel "CRV_Painel_FAT_Ferias" do RH (dados/fat/painel_fat_ferias.xlsx, aba Todos; fora do git: tem nome,
matrícula e idade), de 27/10/2026. Direito ao FAT: efetivo, 12 meses ou mais de empresa e não aposentado. A lista
anterior (validar fat, 237) deixava de fora os 83 com direito e férias vencidas; o painel traz os 320.

Quem fica sem atividade na escala escolhida (painel da frota) pode ir para o FAT; a fila diz quem vai primeiro
(pedido de 28/10/2026): a colheita (trator transbordo, colhedora e caminhão canavieiro, que param na entressafra) e
depois as outras atividades; dentro de cada grupo, menor tempo de casa primeiro. O papel vem da função, como no
pessoal_frota.py: "MOTORISTA ..." é motorista, o resto (OP. MÁQUINAS AGR.) é operador.

Férias e FAT: o FAT suspende o contrato (CLT art. 476-A). O tempo suspenso não conta para as próximas férias (o período
aquisitivo congela), mas as férias já adquiridas não congelam: têm de começar até o "Limite férias" do painel, e não dá
para dar férias com o contrato suspenso. Não haverá férias antes do FAT (decisão de 28/10/2026), então quem vence a
segunda férias na safra de 2027 (limite de abril a novembro de 2027, faixa 2) não vai para o FAT: voltaria e teria de
tirar as férias na safra. A regra fica na página FAT (fat.json leva o mês do limite) e na lista nominal.
"""
import collections
import datetime
import json
import re
from pathlib import Path

import openpyxl

from pessoal_frota import atividade

AQUI = Path(__file__).parent
ARQ = AQUI / "dados" / "fat" / "painel_fat_ferias.xlsx"
SAIDA = AQUI / "fat.json"
TEMPLATE = AQUI / "fat_nominal_template.html"
LOCAL = AQUI / "FAT_lista_nominal.html"
ARTIFACT = AQUI / "build" / "fat_nominal_artifact.html"
FONTE = "Painel FAT & Férias do RH (27/10/2026)"
# a lista nominal publicada no Render, ao lado da página FAT do estudo (aberta: quem tem o link vê os nomes); a página FAT
# só leva o endereço
LISTA_PRIVADA = "https://operacoes-agricolas.onrender.com/escalas_out2026/FAT_lista_nominal.html"
# atividades da colheita, que param na entressafra: vão primeiro (pelo departamento da planilha)
COLHEITA = re.compile(r"TRANSBORDO|COLHEDORA|CANAVIEIRO")
CRITERIO = ("Vão os 258 do grupo 1 da programação do Transporte (06/11/2026); os outros com direito ficam na fila, colheita "
            "primeiro (trator transbordo, colhedora e caminhão canavieiro), depois as outras atividades; dentro de cada grupo, "
            "menor tempo de casa primeiro")
# necessidade de partida na lista nominal: ativos menos a necessidade da escala do Transporte (escala_transporte())
NEC_PADRAO = {"operador": 152, "motorista": 139}
# os escolhidos para o FAT pelo Transporte (relatório "Escala · base de colaboradores", FAT / Operação = FAT; 05/11/2026),
# usados só sem a programação abaixo
ESCOLHIDOS = AQUI / "dados" / "fat" / "fat_escolhidos.pdf"
# programação de FAT e férias de dezembro (Transporte, 06/11/2026): o grupo 1 são os que vão para o FAT (258), com as férias
# de cada um; e a escala de folga da operação de dezembro, com as posições de cada escala (a necessidade)
PROGRAMACAO = AQUI / "dados" / "fat" / "programacao_fat.pdf"
ESCALA = AQUI / "dados" / "frota" / "escala_operacao.pdf"
# janela das férias: a entressafra vai até março; a safra de 2027, de abril a novembro
SAFRA_27 = (datetime.date(2027, 4, 1), datetime.date(2027, 12, 1))
# férias de cada um (05/11/2026), na ordem: períodos aquisitivos do Operacional CTT (exatos), histórico de férias por
# setor (admissão e dias já gozados) e, sem nenhum dos dois, o limite do painel do RH
PERIODOS = AQUI / "dados" / "fat" / "ferias_periodos.xlsx"
HISTORICO = AQUI / "dados" / "fat" / "historico_ferias.xlsx"
HOJE = datetime.date(2026, 11, 5)
# uma turma só (decisão de 09/11/2026): todos entram no FAT depois das férias, de 14/01 a 15/03/2027. "turma" passa a dizer
# só se a pessoa tira férias antes (1: a partir de 15/12, ou 01/12 se o limite vence antes) ou não tem férias a tirar (2)
TURMAS = {1: {"ferias": ("2026-12-15", "2027-01-13"), "fat": ("2027-01-14", "2027-03-15")},
          2: {"ferias": None, "fat": ("2027-01-14", "2027-03-15")}}
INICIO_FAT = datetime.date(2027, 1, 14)


def data_iso(v):
    return v.date().isoformat() if isinstance(v, datetime.datetime) else None


def faixa_ferias(lim):
    """1: o limite cai até mar/27 (vence na entressafra, durante o FAT); 2: na safra de 2027; 3: dez/27 em diante;
    0: sem data."""
    if not lim:
        return 0
    d = datetime.date.fromisoformat(lim)
    return 1 if d < SAFRA_27[0] else 2 if d < SAFRA_27[1] else 3


def _anos(d, n):
    try:
        return d.replace(year=d.year + n)
    except ValueError:
        return d.replace(year=d.year + n, day=28)


def ferias_de_cada_um():
    """Matrícula -> limite (iso) do primeiro período de férias adquirido e ainda não tirado, com a fonte, o saldo e a
    programação que já existir. Sem período pendente (só em aquisição), o limite fica vazio."""
    fer = {}
    if PERIODOS.exists():
        wb = openpyxl.load_workbook(PERIODOS, read_only=True, data_only=True)
        linhas = list(wb["Períodos"].iter_rows(values_only=True))
        por = {}
        for r in linhas[1:]:
            if r[0]:
                por.setdefault(int(r[0]), []).append(dict(zip(linhas[0], r)))
        prog = {}
        linhas_p = list(wb["Programações"].iter_rows(values_only=True))
        for r in linhas_p[1:]:
            if r[0]:
                x = dict(zip(linhas_p[0], r))
                prog[int(r[0])] = f"{x['Início das férias']:%d/%m/%Y} a {x['Fim das férias']:%d/%m/%Y} · {x['Situação']}"
        for m, ps in por.items():
            # todos os períodos ainda não tirados, os adquiridos e o que está em aquisição, pelo limite
            abertos = sorted((x for x in ps if (x["Saldo"] or 0) > 0 and x["Situação"] not in ("Nulo", "Férias gozadas")
                              and x["Data limite"]), key=lambda x: x["Data limite"])
            x = abertos[0] if abertos else None
            fer[m] = {"limite": x["Data limite"].date().isoformat() if x else None, "fonte": "períodos",
                      "saldo": x["Saldo"] if x else None, "prog": prog.get(m) or (str(x["Programação"]) if x and x.get("Programação") else None),
                      "lims": [(y["Data limite"].date().isoformat(), y["Período até"].date().isoformat()) for y in abertos]}
    if HISTORICO.exists():
        h = list(openpyxl.load_workbook(HISTORICO, read_only=True, data_only=True).worksheets[0].iter_rows(values_only=True))
        cab = next(i for i, r in enumerate(h) if r and r[0] == "Setor")
        for r in h[cab + 1:]:
            if not (r[2] and str(r[2]).isdigit()) or int(r[2]) in fer:
                continue
            x = dict(zip(h[cab], r))
            adm = datetime.datetime.strptime(x["Admissão"], "%d/%m/%Y").date()
            k = int((x["Total de dias"] or 0) // 30)                 # períodos já tirados
            # períodos ainda não tirados, do primeiro até o que está em aquisição: fim do período e limite (11 meses depois)
            lims = []
            for j in range(k, k + 4):
                fim = _anos(adm, j + 1) - datetime.timedelta(days=1)
                lims.append(((_anos(fim + datetime.timedelta(days=1), 1) - datetime.timedelta(days=31)).isoformat(), fim.isoformat()))
                if fim >= HOJE:
                    break
            fer[int(r[2])] = {"limite": lims[0][0], "fonte": "histórico", "saldo": None, "prog": None, "lims": lims}
    return fer


# turma 1 (direto) só com nenhuma férias vencendo até 15/12/2027 (decisão de 05/11/2026): quem vai direto não tem como
# tirar férias até o fim da safra de 2027
CORTE_T1 = datetime.date(2027, 12, 15)
INICIO_FERIAS = datetime.date(2026, 12, 15)


def turma_de(lim):
    """1: o primeiro limite de férias em aberto é depois de 15/12/2027 (ou não há): FAT direto; 2: férias e depois FAT."""
    return 1 if not lim or datetime.date.fromisoformat(lim) > CORTE_T1 else 2


def alerta_de(lims, turma):
    """O que as férias de 15/12 a 13/01 da turma 2 não resolvem sozinhas."""
    if turma == 1 or not lims:
        return None
    if datetime.date.fromisoformat(lims[0][0]) < INICIO_FERIAS:
        return "vence antes de 15/12: tirar as férias já"
    if datetime.date.fromisoformat(lims[0][1]) >= INICIO_FERIAS:
        return "período ainda em aquisição em 15/12: férias antecipadas (DP confirmar)"
    if len(lims) > 1 and datetime.date.fromisoformat(lims[1][0]) <= CORTE_T1:
        return f"outro período vence em {lims[1][0][8:10]}/{lims[1][0][5:7]}/{lims[1][0][:4]}: tirar as duas férias antes"
    return None


def _iso(br):
    return f"{br[6:10]}-{br[3:5]}-{br[0:2]}" if br else None


def programacao():
    """Matrícula -> férias do grupo 1 (vão para o FAT) da programação do Transporte: papel, saldo, limite e início e fim
    das férias (vazios para quem não tem férias a tirar)."""
    if not PROGRAMACAO.exists():
        return {}
    import pdfplumber
    dt = r"\d{2}/\d{2}/\d{4}"
    linha = re.compile(rf"^(\d+) (\d{{4,6}}) (.+?) (\d+|) ?({dt}|) ?({dt}|\S) ({dt}|\S) ({dt}) ({dt}) ?(.*)$")
    prog, grupo = {}, None
    with pdfplumber.open(PROGRAMACAO) as pdf:
        for pg in pdf.pages:
            txt = pg.extract_text() or ""
            m = re.search(r"^Grupo (\d)", txt, re.M)
            if m:
                grupo = int(m.group(1))
            if grupo != 1:
                continue
            for l in txt.splitlines():
                m = linha.match(l)
                if m:
                    _, mat, meio, saldo, lim, fi, ff, _, _, _ = m.groups()
                    prog[int(mat)] = {"papel": "motorista" if " Motorista " in f" {meio} " else "operador",
                                      "saldo": int(saldo) if saldo else None, "limite": _iso(lim),
                                      "ini": _iso(fi) if "/" in fi else None, "fim": _iso(ff) if "/" in ff else None}
    return prog


def alerta_prog(f):
    """O que fica fora do FAT único (14/01 a 15/03): quem termina as férias antes de 13/01 ou não tem férias a tirar."""
    if not f["ini"]:
        return "sem férias a tirar: sem atividade de 15/12 a 13/01 (DP)"
    br = lambda iso: f"{iso[8:10]}/{iso[5:7]}"
    if datetime.date.fromisoformat(f["fim"]) < INICIO_FAT - datetime.timedelta(days=1):
        motivo = f"limite {br(f['limite'])}: férias a partir de {br(f['ini'])}" if f["ini"] < "2026-12-15" else f"saldo de {f['saldo']} dias"
        return f"{motivo}; sem atividade de {br((datetime.date.fromisoformat(f['fim']) + datetime.timedelta(days=1)).isoformat())} a 13/01"
    return None


def escala_transporte():
    """Necessidade pela escala de folga da operação de dezembro (Transporte): posições de cada escala (pessoas escaladas
    mais as "a escalar"; "não escalar" não conta), por atividade e papel."""
    if not ESCALA.exists():
        return None
    import pdfplumber
    esc, cur, emit = [], None, None
    with pdfplumber.open(ESCALA) as pdf:
        for pg in pdf.pages[1:]:
            ls = (pg.extract_text() or "").splitlines()
            for j, l in enumerate(ls):
                m = re.match(r"ESCALA DE FOLGA . (.+?) . DEZEMBRO", l)
                if m and j + 1 < len(ls) and re.match(r"^(MOT|OP)-", ls[j + 1]):
                    cur = {"ativ": m.group(1), "papel": "motorista" if ls[j + 1].startswith("MOT") else "operador", "pos": 0}
                    esc.append(cur)
                elif cur and (re.match(r"^\d{4,6}[A-Z]", l) or l.startswith("A ESCALAR")):
                    cur["pos"] += 1
                e = re.search(r"Gerado em (\d{2}/\d{2}/\d{4})", l)
                emit = e.group(1) if e else emit
    # as escalas de linha amarela, ônibus e plantio entram na atividade do resumo do relatório (9 atividades)
    grupo = lambda a: next((g for g in ("LINHA AMARELA", "ÔNIBUS", "PLANTIO") if a.startswith(g)), a)
    ativ = {}
    for e in esc:
        a = ativ.setdefault(grupo(e["ativ"]), {"ativ": grupo(e["ativ"]).capitalize(), "escalas": 0, "operador": 0, "motorista": 0})
        a["escalas"] += 1
        a[e["papel"]] += e["pos"]
    lista = sorted(ativ.values(), key=lambda a: -(a["operador"] + a["motorista"]))
    return {"fonte": f"Escala de folga da operação, dezembro de 2026 (Transporte, {emit})", "atividades": lista,
            "precisa": {p: sum(a[p] for a in lista) for p in ("operador", "motorista")}}


def escolhidos():
    """Matrícula -> papel dos que vão para o FAT: o grupo 1 da programação do Transporte ou, sem ela, os escolhidos do
    relatório "Escala · base de colaboradores"."""
    prog = programacao()
    if prog:
        return {m: f["papel"] for m, f in prog.items()}
    if not ESCOLHIDOS.exists():
        return {}
    import pdfplumber
    esc = {}
    with pdfplumber.open(ESCOLHIDOS) as pdf:
        for pg in pdf.pages:
            for l in (pg.extract_text() or "").splitlines():
                m = re.match(r"^(\d{4,6}) .+? (MOTORISTA|OP\. DE MAQUINAS)", l)
                if m:
                    esc[int(m.group(1))] = "motorista" if m.group(2) == "MOTORISTA" else "operador"
    return esc


ATIVOS = AQUI / "dados" / "frota" / "ativos_transporte.pdf"


def ativos_transporte():
    """Matrículas dos motoristas e operadores ativos no relatório do Transporte (vazio sem o PDF)."""
    if not ATIVOS.exists():
        return set()
    import pdfplumber
    with pdfplumber.open(ATIVOS) as pdf:
        return {int(m.group(1)) for pg in pdf.pages for l in (pg.extract_text() or "").splitlines()
                for m in [re.match(r"^(\d{4,6}) ", l)] if m}


def ler():
    wb = openpyxl.load_workbook(ARQ, read_only=True, data_only=True)
    fer = ferias_de_cada_um()

    def aba(nome):
        linhas = list(wb[nome].iter_rows(values_only=True))
        cab = [str(c).strip() if c else "" for c in linhas[0]]
        return [dict(zip(cab, r)) for r in linhas[1:] if r[0]]

    fora_quadro = {p["Mat."] for p in aba("Divergências") if p["Origem"] == "Só FAT"}
    gente, verificar, safra = [], [], {}
    for r in aba("Todos"):
        dep = str(r["Departamento"] or "").strip()
        p = {"mat": int(r["Mat."]), "nome": str(r["Nome"]).strip(), "funcao": str(r["Função"]).strip(),
             "papel": "motorista" if "MOTORISTA" in str(r["Função"]).upper() else "operador",
             "ativ": atividade(dep), "grupo": 1 if COLHEITA.search(dep.upper()) else 2,
             "ger": str(r["Gerência"] or "").strip().strip("–-").strip() or "–",
             "adm": data_iso(r["Admissão"]), "tempo": str(r["Tempo"] or ""), "idade": r["Idade"],
             "ferias": str(r["Férias"] or ""), "limite": data_iso(r["Limite férias"]), "motivo": str(r["Motivo"] or "")}
        f = fer.get(p["mat"])
        if f:
            p.update(limite=f["limite"], fonte_ferias=f["fonte"], saldo=f["saldo"], prog=f["prog"])
            lims = f["lims"]
        else:
            p.update(fonte_ferias="painel", saldo=None, prog=None)
            lims = [(p["limite"], "2000-01-01")] if p["limite"] else []
        p["faixa"] = faixa_ferias(p["limite"])
        p["turma"] = turma_de(p["limite"])
        p["alerta"] = alerta_de(lims, p["turma"])
        if r["FAT"] == "Tem direito":
            gente.append(p)
        # podem ter direito: efetivos com 12 meses ou mais em que falta conferir a aposentadoria (60+) ou que não
        # entraram no levantamento do FAT
        elif p["motivo"] in ("Aposentadoria a verificar (60+)", "Não consta no levantamento FAT"):
            verificar.append(p)
        # contrato de safra acaba com a safra: sai do quadro na entressafra sem precisar de FAT
        if r["Contrato"] == "SAFRA" and p["mat"] not in fora_quadro:
            safra[p["papel"]] = safra.get(p["papel"], 0) + 1
    # escolhidos do Transporte que não estão no painel do RH: entram com os dados dos períodos e do histórico de férias
    esc = escolhidos()
    ja = {p["mat"] for p in gente + verificar}
    for m, papel in esc.items():
        if m in ja:
            continue
        f = fer.get(m) or {}
        h = next((r for r in openpyxl.load_workbook(HISTORICO, read_only=True, data_only=True).worksheets[0].iter_rows(values_only=True)
                  if r and str(r[2]) == str(m)), None) if HISTORICO.exists() else None
        adm = datetime.datetime.strptime(h[4], "%d/%m/%Y").date() if h else None
        p = {"mat": m, "nome": str(h[1]).strip() if h else str(m), "funcao": str(h[3]).strip() if h else "–", "papel": papel,
             "ativ": atividade(str(h[0])) if h else "–", "grupo": 2, "ger": "–", "adm": adm.isoformat() if adm else None,
             "tempo": "", "idade": None, "ferias": "", "limite": f.get("limite"), "motivo": "Escolhido pelo Transporte, fora do painel do RH",
             "fonte_ferias": f.get("fonte", "–"), "saldo": f.get("saldo"), "prog": f.get("prog")}
        p["faixa"] = faixa_ferias(p["limite"])
        p["turma"] = turma_de(p["limite"])
        p["alerta"] = alerta_de(f.get("lims") or [], p["turma"])
        gente.append(p)
    for p in gente + verificar:
        p["escolhido"] = p["mat"] in esc
        if p["escolhido"] and p["papel"] != esc[p["mat"]]:   # a função do Transporte é a mais nova (mesma base dos ativos)
            p["papel"] = esc[p["mat"]]
    # o Transporte deu direito aos que estavam a conferir (aposentadoria ou fora do levantamento): vão com os escolhidos
    gente += [p for p in verificar if p["escolhido"]]
    verificar = [p for p in verificar if not p["escolhido"]]
    # quem saiu do quadro (não está nos ativos do relatório do Transporte) sai do FAT
    ativos = ativos_transporte()
    if ativos:
        saiu = [p["mat"] for p in gente + verificar if p["mat"] not in ativos]
        gente = [p for p in gente if p["mat"] in ativos]
        verificar = [p for p in verificar if p["mat"] in ativos]
        print(f"fora dos ativos do Transporte (saíram do FAT): {len(saiu)}")
    # uma turma só: as férias de cada um vêm da programação do Transporte (quem vai para o FAT); quem fica, pelo limite
    prog = programacao()
    for p in gente + verificar:
        f = prog.get(p["mat"])
        p["fer_ini"] = p["fer_fim"] = None
        if f:
            p.update(limite=f["limite"] or p["limite"], fonte_ferias="programação do Transporte", saldo=f["saldo"],
                     fer_ini=f["ini"], fer_fim=f["fim"], turma=1 if f["ini"] else 2, alerta=alerta_prog(f))
        else:
            p.update(turma=1 if p["limite"] else 2, alerta=None)
        p["faixa"] = faixa_ferias(p["limite"])
    # fila: os escolhidos do Transporte primeiro; depois colheita; dentro do grupo, menor tempo de casa (admissão mais nova)
    chave = lambda p: (not p["escolhido"], p["grupo"], -(datetime.date.fromisoformat(p["adm"]).toordinal() if p["adm"] else 0), -p["mat"])
    gente.sort(key=chave)
    verificar.sort(key=lambda p: (p["papel"], p["motivo"], chave(p)))
    return gente, verificar, safra


def main():
    gente, verificar, safra = ler()
    conta = lambda lista: {pp: sum(1 for p in lista if p["papel"] == pp) for pp in ("operador", "motorista")}
    escala = escala_transporte()
    publico = {"fonte": FONTE, "criterio": CRITERIO, "n": len(gente), "lista": LISTA_PRIVADA,
               "verificar": conta(verificar), "safra": safra, "escolhidos": conta([p for p in gente if p["escolhido"]]),
               "turmas": TURMAS, "escala": escala,
               "fila": [[p["papel"], p["ativ"], p["grupo"], p["limite"][:7] if p["limite"] else None, p["turma"], 1 if p["alerta"] else 0, 1 if p["escolhido"] else 0] for p in gente]}
    cab = json.dumps({k: v for k, v in publico.items() if k != "fila"}, ensure_ascii=False)[:-1]
    SAIDA.write_text(cab + ', "fila": [\n' + ",\n".join(json.dumps(f, ensure_ascii=False) for f in publico["fila"]) + "\n]}\n",
                     encoding="utf-8")
    # ativos do quadro de motoristas e operadores (só contagens, pessoal_frota.py), para a tabela por papel da lista
    pessoal = AQUI / "pessoal_frota.json"
    ativos = json.loads(pessoal.read_text(encoding="utf-8")).get("ativos", {}) if pessoal.exists() else {}
    nec = {p: max(0, ativos[p] - escala["precisa"][p]) for p in ativos} if escala and ativos else NEC_PADRAO
    dados = {"fonte": FONTE, "criterio": CRITERIO, "nec": nec, "safra27": [d.isoformat() for d in SAFRA_27],
             "turmas": TURMAS, "ativos": ativos, "escala": escala, "gente": gente, "verificar": verificar}
    js = json.dumps(dados, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    pagina = TEMPLATE.read_text(encoding="utf-8").replace("/*FAT*/null", js)
    ARTIFACT.parent.mkdir(exist_ok=True)
    ARTIFACT.write_text(pagina, encoding="utf-8")
    LOCAL.write_text("<!doctype html>\n<html lang=\"pt-BR\" translate=\"no\" class=\"notranslate\"><head>"
                     "<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1, "
                     "viewport-fit=cover\">\n</head><body>\n" + pagina + "\n</body></html>\n", encoding="utf-8")
    t = {k: conta([p for p in gente if p["turma"] == k]) for k in (1, 2)}
    print(f"escala: {escala and escala['precisa']}; vão {conta([p for p in gente if p['escolhido']])}")
    print(f"com férias antes: {t[1]}; sem férias a tirar: {t[2]}; alertas: "
          f"{collections.Counter(p['alerta'] for p in gente if p['alerta'])}; fontes: {collections.Counter(p['fonte_ferias'] for p in gente)}")
    print(f"escolhidos do Transporte: {conta([p for p in gente if p['escolhido']])}; no painel a conferir: "
          f"{conta([p for p in verificar if p['escolhido']])}; fora do painel: {[p['mat'] for p in gente if p['motivo'].startswith('Escolhido')]}")
    print(f"{len(gente)} com direito: {conta(gente)}; colheita {sum(1 for p in gente if p['grupo'] == 1)}; "
          f"a verificar {conta(verificar)}; contrato safra no quadro {safra}")
    return publico


if __name__ == "__main__":
    main()
