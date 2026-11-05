"""Motoristas e operadores ativos (planilha "Quadro motorista operador", out/2026): contagens por papel e por atividade.

python pessoal_frota.py        ->  pessoal_frota.json
python pessoal_frota.py --pdf  ->  só os ativos por papel, do resumo por função do relatório do Transporte
                                   (dados/frota/ativos_transporte.pdf, "Escala · base de colaboradores"); as atividades
                                   ficam as da planilha  (só contagens; a planilha, com nome, matrícula e CPF, fica em dados/frota/ e
não sobe para o git)

O papel vem da função: "MOTORISTA ..." é motorista, "OP. DE MAQUINAS AGRICOLAS ..." é operador. A atividade é o departamento
(ex.: "MOTORISTA - CAMINHAO CANAVIEIRO"). Quem tem data de rescisão anotada fica fora dos ativos.
"""
import collections
import json
import re
from pathlib import Path

import openpyxl

AQUI = Path(__file__).parent
ARQ = AQUI / "dados" / "frota" / "quadro_motorista_operador.xlsx"
PDF = AQUI / "dados" / "frota" / "ativos_transporte.pdf"
SAIDA = AQUI / "pessoal_frota.json"

PALAVRAS = {  # acentos que a planilha não traz
    "CAMINHAO": "caminhão", "CAMINH": "caminhão", "ONIBUS": "ônibus", "VEICULOS": "veículos", "VEICULO": "veículo",
    "CACAMBA": "caçamba", "MAQUINAS": "máquinas", "PATIO": "pátio", "IRRIG": "irrigação", "IRRIGACAO": "irrigação",
    "ADUBACAO": "adubação", "ELETROIMA": "eletroímã", "INDUSTRIA": "indústria", "AMBULANCIA": "ambulância",
    "MANUTENCAO": "manutenção", "PA": "pá", "CARREG.": "carregadeira", "APLIC": "aplicação", "VINH/AGUA": "vinhaça/água",
    "TRANS": "transporte", "MEC.": "mecanizada", "COMBOIO": "comboio", "LOCALIZADA": "localizada", "CCT": "CCT",
}


def atividade(desc):
    """'MOTORISTA - CAMINHAO CANAVIEIRO' -> 'Caminhão canavieiro'."""
    txt = re.sub(r"^(MOTORISTA|OPERADOR|AUXILIARES( TRAINEE)?|MANUTENCAO|FISCAIS)\s*-\s*", "", desc.strip())
    txt = txt.replace("(", " (").replace("  ", " ")
    palavras = [PALAVRAS.get(p, p.lower()) for p in txt.split()]
    s = " ".join(palavras).replace("( ", "(")
    return s[:1].upper() + s[1:]


def main():
    linhas = list(openpyxl.load_workbook(ARQ, read_only=True, data_only=True).worksheets[0].iter_rows(values_only=True))
    col = {str(c).strip(): i for i, c in enumerate(linhas[0]) if c}
    ativos, fora = collections.Counter(), 0
    por_ativ = collections.Counter()
    for r in linhas[1:]:
        if not r[col["Matrícula"]]:
            continue
        resc = str(r[col["Data da Rescisão"]] or "").replace("/", "").strip()
        if str(r[col["Situação"]]).strip() != "A" or resc:
            fora += 1
            continue
        papel = "motorista" if "MOTORISTA" in str(r[col["Nome da Função"]]).upper() else "operador"
        ativos[papel] += 1
        por_ativ[(papel, atividade(str(r[col["Descrição Departamento"]])))] += 1
    lista = [{"papel": p, "nome": n, "n": v} for (p, n), v in sorted(por_ativ.items(), key=lambda x: (x[0][0], -x[1]))]
    saida = {"fonte": "Quadro motorista operador.xlsx (26/10/2026)", "ativos": dict(ativos), "fora": fora, "atividades": lista}
    SAIDA.write_text(json.dumps(saida, ensure_ascii=False, indent=1), encoding="utf-8")
    return saida


def do_pdf():
    """Ativos por papel pelo "Resumo por função" do relatório do Transporte (última página)."""
    import pdfplumber
    with pdfplumber.open(PDF) as pdf:
        txt = pdf.pages[-1].extract_text() or ""
        emitido = re.search(r"Gerado em (\d{2}/\d{2}/\d{4})", txt)
    ativos = collections.Counter()
    for l in txt.splitlines():
        m = re.match(r"^(MOTORISTA|OP\. DE MAQUINAS)\D*?((?:\d+ ){5}\d+)$", l.strip())
        if m:
            ativos["motorista" if m.group(1) == "MOTORISTA" else "operador"] += int(m.group(2).split()[0])
    saida = json.loads(SAIDA.read_text(encoding="utf-8")) if SAIDA.exists() else {"fora": 0, "atividades": []}
    saida.update(fonte=f"Transporte, escala · base de colaboradores ({emitido.group(1) if emitido else '?'})", ativos=dict(ativos))
    SAIDA.write_text(json.dumps(saida, ensure_ascii=False, indent=1), encoding="utf-8")
    return saida


if __name__ == "__main__":
    import sys
    if "--pdf" in sys.argv:
        print("ativos:", do_pdf()["ativos"])
        sys.exit()
    s = main()
    print("ativos:", s["ativos"], "| fora (rescisão):", s["fora"])
    for a in s["atividades"]:
        print(f'  {a["papel"]:9s} {a["n"]:4d}  {a["nome"]}')
