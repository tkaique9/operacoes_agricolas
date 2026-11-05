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
CRITERIO = ("Colheita primeiro (trator transbordo, colhedora e caminhão canavieiro, que param na entressafra), depois as "
            "outras atividades; dentro de cada grupo, menor tempo de casa primeiro")
# necessidade de partida na lista nominal: o que fica sem atividade na escala escolhida que está nos dados da página
# (painel da frota, 28/10/2026). Na lista dá para mudar; o número certo é o da aba FAT da página
NEC_PADRAO = {"operador": 83, "motorista": 124}
# janela das férias: a entressafra vai até março; a safra de 2027, de abril a novembro
SAFRA_27 = (datetime.date(2027, 4, 1), datetime.date(2027, 12, 1))
# férias de cada um (05/11/2026), na ordem: períodos aquisitivos do Operacional CTT (exatos), histórico de férias por
# setor (admissão e dias já gozados) e, sem nenhum dos dois, o limite do painel do RH
PERIODOS = AQUI / "dados" / "fat" / "ferias_periodos.xlsx"
HISTORICO = AQUI / "dados" / "fat" / "historico_ferias.xlsx"
HOJE = datetime.date(2026, 11, 5)
# duas turmas (decisão de 05/11/2026): quem vence a segunda férias na safra de 2027 tira as férias antes e entra depois
# no FAT; os outros vão direto. O fim da turma 1 veio escrito "15/0/2027": vale 15/04/2027 (três meses, como a turma 2)
TURMAS = {1: {"ferias": ("2026-12-15", "2027-01-13"), "fat": ("2027-01-14", "2027-04-15")},
          2: {"ferias": None, "fat": ("2026-12-15", "2027-03-15")}}


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
            pend = [x for x in ps if (x["Saldo"] or 0) > 0 and x["Situação"] not in ("Em aquisição", "Nulo", "Férias gozadas")]
            x = min(pend, key=lambda x: x["Data limite"]) if pend else None
            fer[m] = {"limite": x["Data limite"].date().isoformat() if x else None, "fonte": "períodos",
                      "saldo": x["Saldo"] if x else None, "prog": prog.get(m) or (str(x["Programação"]) if x and x.get("Programação") else None)}
    if HISTORICO.exists():
        h = list(openpyxl.load_workbook(HISTORICO, read_only=True, data_only=True).worksheets[0].iter_rows(values_only=True))
        cab = next(i for i, r in enumerate(h) if r and r[0] == "Setor")
        for r in h[cab + 1:]:
            if not (r[2] and str(r[2]).isdigit()) or int(r[2]) in fer:
                continue
            x = dict(zip(h[cab], r))
            adm = datetime.datetime.strptime(x["Admissão"], "%d/%m/%Y").date()
            k = int((x["Total de dias"] or 0) // 30)                 # períodos já tirados
            fim = _anos(adm, k + 1) - datetime.timedelta(days=1)       # fim do primeiro período ainda não tirado
            lim = None if fim >= HOJE else _anos(fim + datetime.timedelta(days=1), 1) - datetime.timedelta(days=31)
            fer[int(r[2])] = {"limite": lim.isoformat() if lim else None, "fonte": "histórico", "saldo": None, "prog": None}
    return fer


def turma_de(lim):
    """1: a segunda férias vence na safra de 2027 (limite de abril a novembro): férias e depois FAT; 2: os outros."""
    return 1 if lim and SAFRA_27[0] <= datetime.date.fromisoformat(lim) < SAFRA_27[1] else 2


def alerta_de(lim, turma):
    """Limite que cai fora do que a turma resolve: antes de começar, durante ou logo depois do FAT da turma 2."""
    if not lim or turma == 1:
        return None
    d = datetime.date.fromisoformat(lim)
    if d < datetime.date(2026, 12, 15):
        return "vence antes de 15/12: tirar as férias já"
    if d <= datetime.date(2027, 3, 15):
        return "vence durante o FAT (15/12 a 15/03)"
    if d < SAFRA_27[0]:
        return "vence logo depois do FAT, antes da safra"
    return None


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
        else:
            p.update(fonte_ferias="painel", saldo=None, prog=None)
        p["faixa"] = faixa_ferias(p["limite"])
        p["turma"] = turma_de(p["limite"])
        p["alerta"] = alerta_de(p["limite"], p["turma"])
        if r["FAT"] == "Tem direito":
            gente.append(p)
        # podem ter direito: efetivos com 12 meses ou mais em que falta conferir a aposentadoria (60+) ou que não
        # entraram no levantamento do FAT
        elif p["motivo"] in ("Aposentadoria a verificar (60+)", "Não consta no levantamento FAT"):
            verificar.append(p)
        # contrato de safra acaba com a safra: sai do quadro na entressafra sem precisar de FAT
        if r["Contrato"] == "SAFRA" and p["mat"] not in fora_quadro:
            safra[p["papel"]] = safra.get(p["papel"], 0) + 1
    # fila: colheita primeiro; dentro do grupo, menor tempo de casa (admissão mais nova) primeiro
    chave = lambda p: (p["grupo"], -(datetime.date.fromisoformat(p["adm"]).toordinal() if p["adm"] else 0), -p["mat"])
    gente.sort(key=chave)
    verificar.sort(key=lambda p: (p["papel"], p["motivo"], chave(p)))
    return gente, verificar, safra


def main():
    gente, verificar, safra = ler()
    conta = lambda lista: {pp: sum(1 for p in lista if p["papel"] == pp) for pp in ("operador", "motorista")}
    publico = {"fonte": FONTE, "criterio": CRITERIO, "n": len(gente), "lista": LISTA_PRIVADA,
               "verificar": conta(verificar), "safra": safra,
               "turmas": TURMAS,
               "fila": [[p["papel"], p["ativ"], p["grupo"], p["limite"][:7] if p["limite"] else None, p["turma"]] for p in gente]}
    cab = json.dumps({k: v for k, v in publico.items() if k != "fila"}, ensure_ascii=False)[:-1]
    SAIDA.write_text(cab + ', "fila": [\n' + ",\n".join(json.dumps(f, ensure_ascii=False) for f in publico["fila"]) + "\n]}\n",
                     encoding="utf-8")
    # ativos do quadro de motoristas e operadores (só contagens, pessoal_frota.py), para a tabela por papel da lista
    pessoal = AQUI / "pessoal_frota.json"
    ativos = json.loads(pessoal.read_text(encoding="utf-8")).get("ativos", {}) if pessoal.exists() else {}
    dados = {"fonte": FONTE, "criterio": CRITERIO, "nec": NEC_PADRAO, "safra27": [d.isoformat() for d in SAFRA_27],
             "turmas": TURMAS, "ativos": ativos, "gente": gente, "verificar": verificar}
    js = json.dumps(dados, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    pagina = TEMPLATE.read_text(encoding="utf-8").replace("/*FAT*/null", js)
    ARTIFACT.parent.mkdir(exist_ok=True)
    ARTIFACT.write_text(pagina, encoding="utf-8")
    LOCAL.write_text("<!doctype html>\n<html lang=\"pt-BR\" translate=\"no\" class=\"notranslate\"><head>"
                     "<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1, "
                     "viewport-fit=cover\">\n</head><body>\n" + pagina + "\n</body></html>\n", encoding="utf-8")
    t = {k: conta([p for p in gente if p["turma"] == k]) for k in (1, 2)}
    print(f"turma 1 (férias e FAT): {t[1]}; turma 2 (FAT direto): {t[2]}; alertas: "
          f"{collections.Counter(p['alerta'] for p in gente if p['alerta'])}; fontes: {collections.Counter(p['fonte_ferias'] for p in gente)}")
    print(f"{len(gente)} com direito: {conta(gente)}; colheita {sum(1 for p in gente if p['grupo'] == 1)}; "
          f"a verificar {conta(verificar)}; contrato safra no quadro {safra}")
    return publico


if __name__ == "__main__":
    main()
