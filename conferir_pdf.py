"""Confere as contas da impressão das planilhas (PDF) e lista o que não fecha.

    python conferir_pdf.py dados/20.10.26.pdf

Não muda nada: só lê o PDF (programas_pdf.py) e compara cada total com a soma das suas
partes e cada saldo com a conta das colunas. Tolerância de 0,05 ha (arredondamento).
"""
import re
import sys

import programas_pdf as P

TOL = 0.05
achados = []


def confere(onde, oque, impresso, conta, detalhe=""):
    if impresso is None or conta is None:
        return
    if abs(impresso - conta) > TOL:
        achados.append((onde, oque, impresso, round(conta, 2), round(impresso - conta, 2), detalhe))


def v(l, k):
    return l["v"].get(k, 0.0) or 0.0


def semana(path):
    """Página 1: dias × semana, quadro de meta, variações e o total de cada operação."""
    import pymupdf
    s = P.ler_semana(path)
    for a in s["atividades"]:
        nome = a["nome"]
        soma_dias = sum(a["dias"].values())
        if a.get("realSemana") is not None:
            dif = round(a["realSemana"] - soma_dias, 2)
            if abs(dif) > TOL:
                achados.append((f"Semana · {nome}", "realizado da semana × soma dos dias impressos", a["realSemana"],
                                round(soma_dias, 2), dif, "a diferença é o domingo, que a planilha soma e não imprime"
                                if len(s["datas"]) == 6 else ""))
        if a.get("equip") and a.get("metaUnid") is not None and a.get("metaDia") is not None:
            confere(f"Semana · {nome}", "meta do dia = nº de equip./pessoas × meta por unidade", a["metaDia"], a["equip"] * a["metaUnid"])
        if a.get("metaDia") is not None and a.get("metaSemana") is not None:
            confere(f"Semana · {nome}", "meta da semana = meta do dia × 6 dias", a["metaSemana"], a["metaDia"] * 6)
        if a.get("realDia") is not None and a["dias"]:
            ult = a["dias"].get(s["datas"][-1])
            confere(f"Semana · {nome}", "realizado do dia = o do DIA REF", a["realDia"], ult or 0.0)
    # linhas de total de cada operação (as linhas só de números depois das atividades)
    doc = pymupdf.open(path)
    pg = next(p for p in doc if "RESUMO SEMANAL" in P.sem_acento(p.get_text()))
    linhas = P.linhas_da_pagina(pg)
    bloco, frente = [], None
    for _y, ws in linhas:
        nome = " ".join(w[4] for w in ws if w[2] < 160)
        nums = [w for w in ws if w[0] >= 380 and P.NUM.match(w[4])]
        tn = P.sem_acento(nome)
        if nome and not nums and nome.isupper() and len(nome) > 8:
            frente, bloco = nome, []
            continue
        if nome and not tn.startswith(("PESSOAS", "RENDIMENTO", "OBSERVACAO", "ATIVIDADE", "PAGINA")) and nums:
            bloco.append(nome)
            continue
        if (not nome or tn.startswith("OBSERVACAO")) and len(nums) >= 6 and bloco:
            tot = {}
            for w in nums:
                k, x = min(P.DIREITA, key=lambda c: abs(c[1] - w[2]))
                if abs(x - w[2]) <= 14 and not k.startswith("_") and w[4] != "-":
                    tot[k] = P.num(w[4])
            ats = [a for a in s["atividades"] if a["nome"] in bloco]
            ativas = [a for a in ats if (a.get("realSemana") or 0) > 0]
            for k, rot, lista in (("metaSemana", "meta da semana (só atividades com produção)", ativas),
                                  ("realSemana", "realizado da semana", ats), ("metaMes", "meta do mês (atividades com produção)", ativas),
                                  ("realMes", "realizado do mês", ats), ("mesAnt", "realizado do mês anterior", ats)):
                if k in tot:
                    confere(f"Semana · total {frente.title()}", rot, tot[k], sum(a.get(k) or 0 for a in lista))
            bloco = []


def programas(path):
    d = P.ler(path)
    for p in d["programas"]:
        pid = p["id"]
        for sec in p["secoes"]:
            nome_sec = p["nome"] + (" · " + sec["titulo"] if sec["titulo"] else "")
            subs = [l for l in sec["linhas"] if l["t"] == "sub"]
            tots = [l for l in sec["linhas"] if l["t"] == "tot"]
            for l in sec["linhas"]:
                rot = l.get("cod") and f"faz. {l['cod']} {l.get('prop', '')}".strip() or l.get("rot", "")
                onde = f"{nome_sec} · {rot}"
                if pid == "pre":
                    confere(onde, "a realizar = colhida − realizada − reforma", v(l, "aRealizar"),
                            v(l, "colhida") - v(l, "realizada") - v(l, "reforma"))
                elif pid == "bord":
                    confere(onde, "a realizar = aberta − realizada", v(l, "aRealizar"), v(l, "aberta") - v(l, "realizada"))
                elif pid == "cat1":
                    confere(onde, "realizado total = realizada MEC + realizada MAN", v(l, "realTotal"),
                            v(l, "realMec") + v(l, "realMan"))
                    confere(onde, "a realizar (>20 + <20) = aberta − reforma − realizado total",
                            v(l, "maior20") + v(l, "menor20"), v(l, "aberta") - v(l, "reforma") - v(l, "realTotal"))
                elif pid == "adub":
                    confere(onde, "a realizar = a tratar − cana muda − reforma − localizada − cocho", v(l, "aRealizar"),
                            v(l, "aTratar") - v(l, "canaMuda") - v(l, "reforma") - v(l, "localizada") - v(l, "cocho"))
            # total da seção = soma dos subtotais (pré: por região, com todas as regiões impressas)
            if pid == "pre" and subs and tots:
                for k in ("colhida", "aColher", "reforma", "realizada", "aRealizar", "fornecedor"):
                    confere(f"{nome_sec} · Total", f"total = soma dos subtotais por região ({k})", v(tots[-1], k),
                            sum(v(x, k) for x in subs))
            if pid == "adub" and subs and tots:
                # a impressão esconde os blocos sem saldo: o saldo a realizar tem de fechar
                for t in tots:
                    if "ALENCAR" in t["rot"] and "C/" not in t["rot"]:
                        continue
                    confere(f"{nome_sec} · {t['rot']}", "saldo a realizar = soma dos subtotais impressos", v(t, "aRealizar"),
                            sum(v(x, "aRealizar") for x in subs))
                    break
        if pid == "pre":
            ex, tot = p["extras"], [l for s_ in p["secoes"] for l in s_["linhas"] if l["t"] == "tot"][-1]
            if ex.get("Geral") is not None:
                confere("Pré-emergência · ordens", "geral = encerradas + abertas", ex["Geral"],
                        ex.get("Total Encerrado", 0) + ex.get("Total Aberto", 0))
                confere("Pré-emergência · ordens", "abertas = área a colher", ex.get("Total Aberto"), v(tot, "aColher"))
                confere("Pré-emergência · ordens", "encerradas = área colhida + fornecedor", ex.get("Total Encerrado"),
                        v(tot, "colhida") + v(tot, "fornecedor"))
        if pid == "cat1" and p["extras"].get("regioes"):
            reg = p["extras"]["regioes"]
            tot = [l for s_ in p["secoes"] for l in s_["linhas"] if l["t"] == "tot"][-1]
            for k, kt in (("colhida", "aberta"), ("realizada", "realTotal"), ("maior20", "maior20"), ("menor20", "menor20")):
                confere("1ª catação · gráfico por região", f"soma das regiões = total ({k})", sum(r[k] for r in reg.values()), v(tot, kt))
            for n, r in reg.items():
                confere(f"1ª catação · região {n}", "falta aplicar = a realizar <20 + >20", r["faltaAplicar"], r["menor20"] + r["maior20"])
        if pid == "adub":
            tots = {l["rot"]: l for s_ in p["secoes"] for l in s_["linhas"] if l["t"] == "tot"}
            al = next((l for r, l in tots.items() if "ALENCAR" in r and "C/" not in r), None)
            al_secs = [s_ for s_ in p["secoes"] if "Alencar" in s_["titulo"]]
            if al and len(al_secs) >= 2:
                r1 = [l for l in al_secs[0]["linhas"] if l["t"] == "tot"]
                r2s = [l for l in al_secs[1]["linhas"] if l["t"] == "sub"]
                for k in ("aTratar", "cocho", "aRealizar"):
                    confere("Adubação · cocho Alencar", f"total Alencar = total região 1 + subtotais da região 2 ({k})",
                            v(al, k), (v(r1[-1], k) if r1 else 0) + sum(v(x, k) for x in r2s))
            rs = p["extras"].get("resumo")
            loc = next((l for r, l in tots.items() if r.endswith("LOCALIZADA")), None)
            fb = [l for s_ in p["secoes"] if "Fábio" in s_["titulo"] or "Fabio" in s_["titulo"] for l in s_["linhas"] if l["t"] == "tot"]
            if rs and loc and al and fb:
                for k in ("aTratar", "reforma", "localizada", "cocho", "aRealizar"):
                    confere("Adubação · resumo final (geral)", f"geral = localizada + Alencar + Fábio do Prado ({k})", rs.get(k, 0),
                            v(loc, k) + v(al, k) + v(fb[-1], k))


if __name__ == "__main__":
    pdf = sys.argv[1]
    semana(pdf)
    programas(pdf)
    if not achados:
        print("nenhuma conta fora")
    for onde, oque, imp, conta, dif, det in achados:
        print(f"- {onde}: {oque}: impresso {imp:,.2f} · conta {conta:,.2f} · diferença {dif:+,.2f}" + (f" ({det})" if det else ""))
