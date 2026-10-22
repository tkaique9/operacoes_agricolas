"""Gera a página "Escalas por departamento" (HTML dinâmico) para a diretoria validar a escala de cada setor.

python gerar_departamentos.py  ->  Escalas_por_Departamento_OUT2026.html  (abre direto no navegador)
                                   build/artifact_departamentos.html       (mesma página, para publicar)
Começa pela oficina (102.xxx). Só números por departamento: nomes e matrículas ficam no oficina.json.
"""
import json
from pathlib import Path

import analise
import oficina
from formatos import cargo

AQUI = Path(__file__).parent
TEMPLATE = AQUI / "departamentos_template.html"
LOCAL = AQUI / "Escalas_por_Departamento_OUT2026.html"
ARTIFACT = AQUI / "build" / "artifact_departamentos.html"

NOTA_OF = {
    "5x2": "Mesmos 3 turnos de 8 h e mesmos horários. Na noite (23:00–07:20) a hora noturna reduzida faz as 7h20 "
           "contarem cerca de 8h23: terminar às 07:00 ou ter norma coletiva com hora de 60 min.",
    "4x2": "Mesmos 3 turnos de 8 h, 4 dias de trabalho e 2 de folga. Mais descanso, mais gente.",
    "12x36": "Dia 07:00–19:00 e noite 19:00–07:00, 1 h de intervalo. CLT art. 59-A com norma coletiva; dispensa a "
             "licença do MTE em atividade insalubre. Na noite a hora reduzida continua valendo e gera pagamento a mais.",
    "3x3": "Mesmo quadro da 12x36, com 3 plantões seguidos e 3 dias de folga. Depende de compensação em norma coletiva "
           "e, em atividade insalubre, de licença do MTE ou cláusula coletiva.",
    "4x2_12": "Não serve na oficina: todos são CLT urbano e, com a hora noturna reduzida, o plantão da noite passa de "
              "10 h contadas (8 h + 2 h extras é o teto).",
}
ROTULOS_GRADE = {
    "5x2": [f"Turma {x}" for x in "ABCDEFG"], "4x2": [f"Turma {x}" for x in "ABC"],
    "12x36": ["A · dia", "B · dia", "C · noite", "D · noite"], "3x3": ["A · dia", "B · dia", "C · noite", "D · noite"],
    "4x2_12": ["A · dia", "B · dia", "C · dia", "D · noite", "E · noite", "F · noite"],
}
NOMES = {  # nomes como a diretoria lê (a planilha vem sem acento)
    "102.001.001": "Gerência da Oficina", "102.001.011": "Planejamento e Controle de Manutenção",
    "102.001.012": "Laboratório de Óleo", "102.001.013": "Serviços Gerais", "102.002.002": "Borracharia",
    "102.002.003": "Elétrica Automotiva", "102.002.004": "Caldeiraria / Soldagem", "102.002.005": "Lanternagem / Pintura",
    "102.002.006": "Lavagem de Veículos", "102.002.007": "Lubrificação Rampa", "102.002.009": "Hidráulica",
    "102.002.012": "Torno", "102.003.001": "Máquinas Pesadas", "102.004.001": "Colhedoras de Cana",
    "102.005.001": "Caminhões / Veículos", "102.005.003": "Veículos Leves", "102.006.001": "Bombas / Irrigação",
    "102.006.002": "Implementos Agrícolas", "102.006.003": "Reboques",
}
SUB = {"102.001.011": "Ferramentaria"}  # o grupo em azul do PCM
REGRA_TXT = {"24h": "Roda 24 h", "azul": "os 4 em azul rodam 24 h", "1turno": "1 por turno nas 24 h, o resto no horário ADM",
             "2turnos": "Turno A e B (sem noite)", "5x2": "5x2", "nao": "Não precisa simular",
             "sem_nota": "Sem anotação (ficou 5x2)"}


def dados():
    o = oficina.main()
    A = json.loads((AQUI / "analise.json").read_text(encoding="utf-8"))
    cen = {c["id"]: c for c in A["cenarios"]}
    escalas = []
    for cid in oficina.SUGS:
        c = cen[cid]
        g = A["modelo_" + cid]["grade"]
        chk, _ = analise.checar_grade(g)
        escalas.append({"id": cid, "nome": c["nome"], "curto": c["curto"], "turnos": c["turnos"], "trab": c["trab"],
                        "ciclo": c["ciclo"], "h_sem": c["h_sem"], "he_sem": c["he_sem"], "folgas_sem": c["folgas_sem"],
                        "horario": c["horario"], "risco": oficina.RISCO_OF[cid], "nota": NOTA_OF[cid],
                        "grade": g, "rotulos": ROTULOS_GRADE[cid],
                        "dom": max(x["domingos_folga"] for x in chk), "sem": chk[0]["semanas"],
                        "seq": max(x["max_seguidos"] for x in chk)})
    deps, dia = [], []
    for d in o["departamentos"]:
        ps = [p for p in o["pessoas"] if p["dep"] == d["dep"]]
        tipos = d["tipos"]
        base = {"dep": d["dep"], "nome": NOMES.get(d["dep"]) or cargo(d["nome"].replace("MANUTENCAO - ", "")),
                "sub": SUB.get(d["dep"]),
                "regra": d["regra"], "regra_txt": REGRA_TXT[d["regra"]], "nota": " · ".join(d["notas"]), "n": d["n"]}
        if d["postos"] is None:
            base["vai"] = ("Fica como está (não simula)" if d["regra"] == "nao" else "Segunda a sexta, 8 h por dia (40 h)")
            base["lideres"] = tipos.get("Sem registro", 0)
            base["adm_hoje"] = tipos.get("ADM", 0)
            dia.append(base)
            continue
        hoje = {t: sum(1 for p in ps if p.get("turno") == t and (d["regra"] != "azul" or p["azul"])) for t in oficina.TURNOS}
        pres = {t: round(sum(p.get("pres", 0) for p in ps if p.get("turno") == t and (d["regra"] != "azul" or p["azul"])), 2)
                for t in oficina.TURNOS}
        if d["dep"] == "102.001.012":
            base["hoje_txt"] = "2 em 12x36 só de dia"
        base.update(equipe=d["equipe"], adm=d["adm_usavel"], adm_hoje=tipos.get("ADM", 0), hoje=hoje, pres=pres,
                    postos=d["postos"],
                    sugerida=d["sugerida"], lideres=tipos.get("Sem registro", 0), afastados=tipos.get("Afastado", 0),
                    sem_origem=sum(1 for p in ps if p["tipo"] == "Sem escala de origem" and (d["regra"] != "azul" or p["azul"])),
                    resto_5x2=(d["n"] - d["equipe"] - tipos.get("Sem registro", 0) - tipos.get("Afastado", 0)
                               if d["regra"] in ("azul",) else 0))
        deps.append(base)
    # o resto do PCM (fora dos 4 em azul) também vai para 5x2 de dia
    for b in deps:
        if b["regra"] == "azul":
            dia.insert(0, {"dep": b["dep"], "nome": b["nome"], "sub": "fora a Ferramentaria", "regra": "5x2", "parte": True,
                           "regra_txt": "5x2", "nota": b["nota"], "n": b["resto_5x2"],
                           "vai": "Segunda a sexta, 8 h por dia (40 h)", "lideres": 0})
    tot = {"n": len(o["pessoas"]), "lideres": sum(1 for p in o["pessoas"] if p["tipo"] == "Sem registro"),
           "afastados": sum(1 for p in o["pessoas"] if p["tipo"] == "Afastado"),
           "clt_urbano": o["clt_rural"].get("C", 0), "sem_clt": o["clt_rural"].get(None, 0),
           "rural": o["clt_rural"].get("R", 0)}
    return {"area": "Oficina", "base": "20/10/2026", "escalas": escalas, "deps": deps, "dia": dia, "tot": tot}


def gerar():
    js = json.dumps(dados(), ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    logo = (AQUI / "logo_crv_datauri.txt").read_text(encoding="ascii").strip()
    pagina = TEMPLATE.read_text(encoding="utf-8").replace("/*DADOS*/null", js).replace("/*LOGO*/", logo)
    ARTIFACT.parent.mkdir(exist_ok=True)
    ARTIFACT.write_text(pagina, encoding="utf-8")
    corte = pagina.index("</style>") + len("</style>")
    LOCAL.write_text("<!doctype html>\n<html lang=\"pt-BR\"><head><meta charset=\"utf-8\">\n"
                     "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1, viewport-fit=cover\">\n"
                     + pagina[:corte] + "\n</head><body>\n" + pagina[corte:] + "\n</body></html>\n", encoding="utf-8")
    return LOCAL


if __name__ == "__main__":
    print(gerar())
