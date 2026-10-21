"""Arquivos de apoio das capturas (tudo fictício): a planilha ANTIGA (uma aba por mês, com fórmulas próprias) e a conferência contra os extratos.
Lê os Controle Financeiro 2025/2026.xlsx sintéticos da demonstração. Uso: python gerar_apoio_demo.py <00_Painel BI> <pasta de saída>
"""
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Side, Border

sys.stdout.reconfigure(encoding="utf-8")
COLUNAS = ["Realiz.", "Venc.", "C. C.", "  Fornecedor/ Cliente", "Descrição", "Detalhamento/Categoria II", "Tipo", "Categoria", "Desembolso", "Conta", "Valor", "Saldo", "Comentários"]
SINAL = {"Operacional": 1, "Receita.Financeira": 1, "Aporte": 1, "Entrada": 1}
MESES_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
azul = PatternFill("solid", fgColor="1F4E79")
cinza = PatternFill("solid", fgColor="D9E2F3")
vermelho = PatternFill("solid", fgColor="F8CBAD")
verde = PatternFill("solid", fgColor="C6E0B4")


def ler(pasta):
    linhas = []
    for ano in (2025, 2026):
        ws = openpyxl.load_workbook(Path(pasta) / ("Controle Financeiro %d.xlsx" % ano), data_only=True)["Controle Financeiro"]
        for row in ws.iter_rows(min_row=9, values_only=True):
            if row[0] is None:
                continue
            linhas.append(dict(zip(COLUNAS, row)))
    return linhas


def sinal(l):
    if l["Categoria"] == "Investimento":
        return 1 if l["Detalhamento/Categoria II"] == "Venda de Ativo Imobilizado" else -1
    return SINAL.get(l["Categoria"], -1)


def planilha_antiga(linhas, destino):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    por_mes = defaultdict(list)
    for l in linhas:
        por_mes[(l["Realiz."].year, l["Realiz."].month)].append(l)
    for (ano, mes), ls in sorted(por_mes.items()):
        ws = wb.create_sheet("Controle Financeiro %02d-%d" % (mes, ano))
        ws["A1"] = "Finanças %s/%d" % (MESES_PT[mes - 1], ano)
        ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
        ws["A3"], ws["B3"] = "Saldo inicial de caixa", None
        if mes == 7 and ano == 2025:
            ws["D3"] = 668000
        else:
            ws["D3"] = "='Controle Financeiro %02d-%d'!L%d" % ((mes - 1 or 12), ano if mes > 1 else ano - 1, 8 + len(por_mes[(ano if mes > 1 else ano - 1, mes - 1 or 12)]) if (ano if mes > 1 else ano - 1, mes - 1 or 12) in por_mes else 9)
        ws["D3"].number_format = 'R$ #,##0.00'
        for c, n in enumerate(COLUNAS, start=1):
            cel = ws.cell(8, c, n)
            cel.fill, cel.font = azul, Font(bold=True, color="FFFFFF")
            cel.alignment = Alignment(horizontal="center", wrap_text=True)
        for k, l in enumerate(ls):
            r = 9 + k
            for c, n in enumerate(COLUNAS, start=1):
                if n == "Saldo":
                    continue
                cel = ws.cell(r, c, l.get(n))
                if n in ("Realiz.", "Venc."):
                    cel.number_format = "dd/mm/yyyy"
                if n == "Valor":
                    cel.number_format = 'R$ #,##0.00'
            prev = "$D$3" if k == 0 else "L%d" % (r - 1)
            ws.cell(r, 12, '=%s+IF(OR(H%d="Operacional",H%d="Receita.Financeira",H%d="Aporte",H%d="Entrada",AND(H%d="Investimento",F%d="Venda de Ativo Imobilizado")),K%d,-K%d)' % (prev, r, r, r, r, r, r, r, r)).number_format = 'R$ #,##0.00'
        ult = 8 + len(ls)
        # resumo com fórmulas próprias de cada aba
        for k, (rot, cat) in enumerate([("Receita Operacional", "Operacional"), ("Custo Operacional", "Custo.Operacional"), ("Despesa com Pessoal", "Despesa.com.Pessoal"),
                                        ("Tributos", "Tributo"), ("Despesa Fixa", "Despesa.Fixa")]):
            ws.cell(3 + k, 14, rot).font = Font(bold=True)
            ws.cell(3 + k, 15, '=SUMIFS(K9:K%d,H9:H%d,"%s")' % (ult, ult, cat)).number_format = 'R$ #,##0.00'
        for col, w in zip("ABCDEFGHIJKLMNO", (11, 11, 8, 26, 38, 28, 13, 19, 13, 18, 15, 17, 14, 22, 16)):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A9"
    wb.save(destino)


def conferencia(linhas, destino):
    """Diferença planilha x extrato, conta a conta e mês a mês (movimento do mês). O extrato é derivado dos mesmos lançamentos; na 1ª rodada
    uma folha de R$ 7.000,00 do mês de nov/2025 estava sem lançar na planilha (cenário da demonstração)."""
    contas = ["Banco Alfa", "Banco Beta", "Banco Gama", "Cooperativa Delta", "Dinheiro"]
    mov = defaultdict(float)
    for l in linhas:
        mov[(l["Conta"], l["Realiz."].year, l["Realiz."].month)] += sinal(l) * l["Valor"]
    meses = sorted({(a, m) for (_, a, m) in mov})
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Conferência"
    ws["A1"] = "Conferência com os extratos: diferença entre a planilha e o extrato (R$), por conta e mês"
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")

    def bloco(topo, titulo, com_erro):
        ws.cell(topo, 1, titulo).font = Font(bold=True)
        ws.cell(topo + 1, 1, "Conta").fill = azul
        ws.cell(topo + 1, 1).font = Font(bold=True, color="FFFFFF")
        for j, (a, m) in enumerate(meses):
            c = ws.cell(topo + 1, 2 + j, "%s/%02d" % (MESES_PT[m - 1], a % 100))
            c.fill, c.font, c.alignment = azul, Font(bold=True, color="FFFFFF"), Alignment(horizontal="center")
        fechados = 0
        for i, conta in enumerate(contas):
            ws.cell(topo + 2 + i, 1, conta).font = Font(bold=True)
            for j, (a, m) in enumerate(meses):
                dif = -7000.0 if (com_erro and conta == "Banco Beta" and (a, m) == (2025, 11)) else 0.0
                c = ws.cell(topo + 2 + i, 2 + j, dif)
                c.number_format = '#,##0.00;[Red]-#,##0.00;0.00'
                c.fill = verde if dif == 0 else vermelho
                c.alignment = Alignment(horizontal="center")
                fechados += dif == 0
        total = len(contas) * len(meses)
        ws.cell(topo + 2 + len(contas), 1, "%d de %d pares conta-mês fecham em R$ 0,00" % (fechados, total)).font = Font(bold=True, color="1F4E79")
        return total

    bloco(3, "1ª rodada (antes da correção)", True)
    ws["A11"] = "Diferença em Banco Beta, nov/25: folha de pagamento de R$ 7.000,00 sem lançamento. A conta e o valor apontaram o lançamento que faltava."
    ws["A11"].font = Font(italic=True, color="C00000")
    bloco(13, "Depois de lançar a folha", False)
    ws.column_dimensions["A"].width = 24
    for j in range(len(meses)):
        ws.column_dimensions[openpyxl.utils.get_column_letter(2 + j)].width = 11
    wb.save(destino)
    print("pares:", len(contas) * len(meses))


if __name__ == "__main__":
    pasta, saida = Path(sys.argv[1]), Path(sys.argv[2])
    saida.mkdir(parents=True, exist_ok=True)
    ls = ler(pasta)
    planilha_antiga(ls, saida / "Controle Financeiro (modelo antigo).xlsx")
    conferencia(ls, saida / "Conferencia com extratos.xlsx")
    print("gerado em", saida)
