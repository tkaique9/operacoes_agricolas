"""FAT: motoristas e operadores com direito, as férias de cada um e a fila de quem vai.

python fat.py  ->  fat.json                    (sem nomes: a fila com papel, atividade e mês do limite das férias,
                                                para as páginas FAT do estudo)
               ->  FAT_lista_nominal.html      (com nome e matrícula: fica na máquina, fora do git)
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
tirar as férias na safra. A regra fica nas páginas (fat.json leva o mês do limite) e na lista nominal.
"""
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
# a lista nominal publicada como página privada do claude.ai (só abre para quem tem acesso): as páginas FAT do estudo,
# que são públicas, só levam o endereço
LISTA_PRIVADA = "https://claude.ai/artifact/VYcCeLG9VKAHniU3EFaaLh"
# atividades da colheita, que param na entressafra: vão primeiro (pelo departamento da planilha)
COLHEITA = re.compile(r"TRANSBORDO|COLHEDORA|CANAVIEIRO")
CRITERIO = ("Colheita primeiro (trator transbordo, colhedora e caminhão canavieiro, que param na entressafra), depois as "
            "outras atividades; dentro de cada grupo, menor tempo de casa primeiro")
# necessidade de partida na lista nominal: o que fica sem atividade na escala escolhida que está nos dados da página
# (painel da frota, 28/10/2026). Na lista dá para mudar; o número certo é o da aba FAT da página
NEC_PADRAO = {"operador": 83, "motorista": 124}
# janela das férias: a entressafra vai até março; a safra de 2027, de abril a novembro
SAFRA_27 = (datetime.date(2027, 4, 1), datetime.date(2027, 12, 1))


def data_iso(v):
    return v.date().isoformat() if isinstance(v, datetime.datetime) else None


def faixa_ferias(lim):
    """1: o limite cai até mar/27 (vence na entressafra, durante o FAT); 2: na safra de 2027; 3: dez/27 em diante;
    0: sem data."""
    if not lim:
        return 0
    d = datetime.date.fromisoformat(lim)
    return 1 if d < SAFRA_27[0] else 2 if d < SAFRA_27[1] else 3


def ler():
    wb = openpyxl.load_workbook(ARQ, read_only=True, data_only=True)

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
        p["faixa"] = faixa_ferias(p["limite"])
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
               "fila": [[p["papel"], p["ativ"], p["grupo"], p["limite"][:7] if p["limite"] else None] for p in gente]}
    cab = json.dumps({k: v for k, v in publico.items() if k != "fila"}, ensure_ascii=False)[:-1]
    SAIDA.write_text(cab + ', "fila": [\n' + ",\n".join(json.dumps(f, ensure_ascii=False) for f in publico["fila"]) + "\n]}\n",
                     encoding="utf-8")
    dados = {"fonte": FONTE, "criterio": CRITERIO, "nec": NEC_PADRAO, "safra27": [d.isoformat() for d in SAFRA_27],
             "gente": gente, "verificar": verificar}
    js = json.dumps(dados, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    pagina = TEMPLATE.read_text(encoding="utf-8").replace("/*FAT*/null", js)
    ARTIFACT.parent.mkdir(exist_ok=True)
    ARTIFACT.write_text(pagina, encoding="utf-8")
    LOCAL.write_text("<!doctype html>\n<html lang=\"pt-BR\" translate=\"no\" class=\"notranslate\"><head>"
                     "<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1, "
                     "viewport-fit=cover\">\n</head><body>\n" + pagina + "\n</body></html>\n", encoding="utf-8")
    print(f"{len(gente)} com direito: {conta(gente)}; colheita {sum(1 for p in gente if p['grupo'] == 1)}; "
          f"a verificar {conta(verificar)}; contrato safra no quadro {safra}")
    return publico


if __name__ == "__main__":
    main()
