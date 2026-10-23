"""Faz os gráficos entrarem por categoria (ou por série) na apresentação.

O modelo de objetos do PowerPoint recusa esse tipo de animação via automação ("Invalid
request"), então ela é gravada direto no XML do slide, no formato que o próprio PowerPoint
usa: um passo para eixos e legenda (bldStep="gridLegend") e um passo por categoria ou série,
cada um "após o anterior", e <a:bldChart> na lista de builds. Depois deste script, o
animar_e_exportar.ps1 reabre o arquivo no PowerPoint para conferir que ele leu cada passo.

Uso:  python graficos_por_etapa.py <entrada.pptx> <saida.pptx> <animacoes.json>
"""

import copy
import json
import sys
import zipfile

from lxml import etree

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "{%s}" % NS_P
A = "{%s}" % NS_A
NS = {"p": NS_P, "a": NS_A}


def alvo_grafico(par, passo, serie, categoria):
    for sp in par.iter(P + "spTgt"):
        for velho in list(sp):
            sp.remove(velho)
        ge = etree.SubElement(sp, P + "graphicEl")
        ch = etree.SubElement(ge, A + "chart")
        ch.set("seriesIdx", str(serie))
        ch.set("categoryIdx", str(categoria))
        ch.set("bldStep", passo)


def define_duracao(par, ms):
    for efeito in par.iter(P + "animEffect"):
        efeito.find(P + "cBhvr").find(P + "cTn").set("dur", str(ms))


def expandir(raiz, nome, modo, etapas, ms):
    spid = raiz.xpath(f'//p:cNvPr[@name="{nome}"]', namespaces=NS)[0].get("id")
    efeito = next(c for c in raiz.xpath('//p:timing//p:cTn[@presetClass="entr"]', namespaces=NS)
                  if c.xpath(f'.//p:spTgt[@spid="{spid}"]', namespaces=NS))
    par_efeito = efeito.getparent()
    par_grupo = par_efeito.getparent().getparent().getparent()  # par "após o anterior" que contém o efeito
    lista = par_grupo.getparent()                                 # childTnLst do clique
    modelo = copy.deepcopy(par_efeito)

    # passo 0: eixos, rótulos de categoria e legenda, no lugar do efeito original
    alvo_grafico(par_efeito, "gridLegend", -3, -3)
    define_duracao(par_efeito, min(ms, 300))

    pos = lista.index(par_grupo)
    for i in range(etapas):
        passo = copy.deepcopy(modelo)
        ctn = passo.find(P + "cTn")
        ctn.set("nodeType", "afterEffect")
        ctn.find(P + "stCondLst").find(P + "cond").set("delay", "0")
        if modo == "categoria":
            alvo_grafico(passo, "category", -4, i)
        else:
            alvo_grafico(passo, "series", i, -4)
        define_duracao(passo, ms)
        grupo = etree.Element(P + "par")
        gc = etree.SubElement(grupo, P + "cTn", {"id": "0", "fill": "hold"})
        etree.SubElement(etree.SubElement(gc, P + "stCondLst"), P + "cond", {"delay": "0"})
        etree.SubElement(gc, P + "childTnLst").append(passo)
        lista.insert(pos + 1 + i, grupo)

    bld = raiz.xpath(f'//p:bldLst/p:bldGraphic[@spid="{spid}"]', namespaces=NS)[0]
    for velho in list(bld):
        bld.remove(velho)
    sub = etree.SubElement(bld, P + "bldSub")
    etree.SubElement(sub, A + "bldChart", {"bld": "category" if modo == "categoria" else "series"})


def recalcula_atrasos(raiz):
    """Cada bloco "após o anterior" começa quando o anterior termina (como o PowerPoint grava)."""
    for seq in raiz.xpath('//p:cTn[@nodeType="mainSeq"]', namespaces=NS):
        for clique in seq.find(P + "childTnLst").findall(P + "par"):
            t = 0
            for bloco in clique.find(P + "cTn").find(P + "childTnLst").findall(P + "par"):
                bc = bloco.find(P + "cTn")
                bc.find(P + "stCondLst").find(P + "cond").set("delay", str(t))
                compr = 0
                for ef in bc.find(P + "childTnLst").findall(P + "par"):
                    ec = ef.find(P + "cTn")
                    atraso = int(ec.find(P + "stCondLst").find(P + "cond").get("delay", "0"))
                    durs = [int(x.get("dur")) for x in ec.iter(P + "cTn") if (x.get("dur") or "").isdigit()]
                    compr = max(compr, atraso + max(durs, default=0))
                t += compr


def renumera(raiz):
    timing = raiz.find(".//" + P + "timing")
    mapa = {}
    for i, c in enumerate(timing.iter(P + "cTn"), start=1):
        if c.get("id") != "0":
            mapa[c.get("id")] = str(i)
        c.set("id", str(i))
    for tn in timing.iter(P + "tn"):
        tn.set("val", mapa.get(tn.get("val"), tn.get("val")))


def main():
    entrada, saida, roteiro = sys.argv[1:4]
    spec = json.load(open(roteiro, encoding="utf-8"))
    por_slide = {}
    for e in spec["efeitos"]:
        if e.get("grafico"):
            por_slide.setdefault(e["slide"], []).append(e)
    with zipfile.ZipFile(entrada) as zin, zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            dados = zin.read(item.filename)
            nome = item.filename
            if nome.startswith("ppt/slides/slide") and nome.endswith(".xml"):
                n = int(nome[len("ppt/slides/slide"):-4])
                if n in por_slide:
                    raiz = etree.fromstring(dados)
                    for e in por_slide[n]:
                        expandir(raiz, e["forma"], e["grafico"], int(e["etapas"]), int(round(e["duracao"] * 1000)))
                        print(f"slide {n}: {e['forma']} por {e['grafico']}, {e['etapas']} passos")
                    recalcula_atrasos(raiz)
                    renumera(raiz)
                    dados = etree.tostring(raiz, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, dados)


if __name__ == "__main__":
    main()
