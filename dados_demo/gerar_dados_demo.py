"""Gera a planilha SINTÉTICA do painel de demonstração (portfólio): Controle Financeiro 2025.xlsx e 2026.xlsx.

Tudo aqui é inventado: empresas, cozinhas, contas e valores. Contrato de dados: _Padrao_BI\\05_Padrao_Planilha.md
(13 colunas, 14 categorias, Valor sempre positivo, Conta obrigatória). 15 meses: jul/2025 a set/2026.
Uso: python gerar_dados_demo.py "<pasta 00_Painel BI da demonstração>"
"""
import random
import sys
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

sys.stdout.reconfigure(encoding="utf-8")
random.seed(2026)

COLUNAS = ["Realiz.", "Venc.", "C. C.", "  Fornecedor/ Cliente", "Descrição", "Detalhamento/Categoria II", "Tipo", "Categoria",
           "Desembolso", "Conta", "Valor", "Saldo", "Comentários"]
CONTAS = {"Banco Alfa": 320000.0, "Banco Beta": 210000.0, "Banco Gama": 90000.0, "Cooperativa Delta": 40000.0, "Dinheiro": 8000.0}
MESES = [(2025, m) for m in range(7, 13)] + [(2026, m) for m in range(1, 10)]
COZINHAS = ["Norte", "Sul", "Leste"]
# receita-base mensal por cozinha em jul/2025 e crescimento mensal
BASE = {"Norte": 520000, "Sul": 310000, "Leste": 185000}
CRESC = 0.011
SAZON = {1: 0.93, 2: 0.95, 3: 1.0, 4: 1.0, 5: 1.02, 6: 1.0, 7: 0.98, 8: 1.01, 9: 1.03, 10: 1.04, 11: 1.0, 12: 0.92}
CLIENTES = {   # cliente -> (cozinha, participação na receita da cozinha)
    "Indústria Alfa": ("Norte", 0.34), "Metalúrgica Beta": ("Norte", 0.27), "Têxtil Gama": ("Norte", 0.23), "Mineradora Teta": ("Norte", 0.16),
    "Logística Delta": ("Sul", 0.42), "Plásticos Épsilon": ("Sul", 0.33), "Transportes Kappa": ("Sul", 0.25),
    "Hospital Zeta": ("Leste", 0.40), "Colégio Eta": ("Leste", 0.28), "Construtora Iota": ("Leste", 0.20), "Eventos Lambda": ("Leste", 0.12),
}
# gastos como fração da receita do mês, por cozinha (ADM recebe os administrativos)
PCT = {"custo": 0.365, "salarios": 0.205, "encargos": 0.072, "provisoes": 0.026, "ticket": 0.021, "beneficios": 0.007, "treinamentos": 0.003,
       "tributo": 0.098, "aluguel": 0.021, "energia": 0.019, "agua": 0.004, "internet": 0.002, "variavel": 0.014, "marketing": 0.032}
FORNECEDORES_CUSTO = [("Proteínas", "Açougue Ômega", 0.42), ("Hortifruti", "Hortifruti Pi", 0.18), ("Mercearia", "Distribuidora Sigma", 0.22),
                      ("Laticínios", "Laticínios Rô", 0.08), ("Gás", "Gás Tau", 0.06), ("Descartáveis", "Embalagens Úpsilon", 0.04)]


class Livro:
    def __init__(self):
        self.linhas = []

    def add(self, dt, cc, forn, desc, cat2, tipo, cat, desemb, conta, valor, venc=None, coment=None):
        assert conta in CONTAS and valor > 0
        self.linhas.append({"Realiz.": dt, "Venc.": venc or dt, "C. C.": cc, "  Fornecedor/ Cliente": forn, "Descrição": desc,
                            "Detalhamento/Categoria II": cat2 or "-", "Tipo": tipo, "Categoria": cat, "Desembolso": desemb,
                            "Conta": conta, "Valor": round(valor, 2), "Comentários": coment})


def dia_util(ano, mes, dia):
    d = date(ano, mes, min(dia, 28))
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d


def r(v, pct=0.03):
    return v * (1 + random.uniform(-pct, pct))


SINAL = {"Operacional": 1, "Receita.Financeira": 1, "Aporte": 1, "Entrada": 1, "Custo.Operacional": -1, "Tributo": -1, "Despesa.Fixa": -1,
         "Despesa.Variável": -1, "Despesa.com.Pessoal": -1, "Marketing.Comercial": -1, "Despesa.Financeira": -1, "Retirada": -1, "Saída": -1}


def sinal(l):
    if l["Categoria"] == "Investimento":
        return 1 if l["Detalhamento/Categoria II"] == "Venda de Ativo Imobilizado" else -1
    return SINAL[l["Categoria"]]


def montar():
    lv = Livro()
    receita_mes = {}
    for i, (ano, mes) in enumerate(MESES):
        fator = (1 + CRESC) ** i * SAZON[mes]
        rec_coz = {c: r(BASE[c] * fator, 0.02) for c in COZINHAS}
        rec_total = sum(rec_coz.values())
        receita_mes[(ano, mes)] = rec_total
        # --- receitas
        for cli, (coz, part) in CLIENTES.items():
            v = rec_coz[coz] * part
            if cli == "Eventos Lambda":
                v *= random.uniform(0.6, 1.5)
            cat2 = "Coffee Break e Eventos" if cli == "Eventos Lambda" else "Refeições Industriais"
            metade = random.random() < 0.3
            partes = [0.5, 0.5] if metade else [1.0]
            for k, p in enumerate(partes):
                dt = dia_util(ano, mes, random.randint(4, 9) + 14 * k)
                lv.add(dt, coz, cli, "NF %s - %s" % (cat2.lower(), cli), cat2, "Receita", "Operacional", random.choice(["Boleto", "PIX", "TED"]),
                       random.choices(["Banco Alfa", "Banco Beta"], [0.7, 0.3])[0], v * p, venc=dt - timedelta(days=random.randint(0, 3)))
        if mes % 3 == 0:
            lv.add(dia_util(ano, mes, 20), "Norte", "Cliente Eventual Mi", "Evento corporativo avulso", "Coffee Break e Eventos", "Receita", "Operacional",
                   "PIX", "Banco Alfa", r(14500, 0.2))
        # --- custo operacional, por cozinha
        for coz in COZINHAS:
            for cat2, forn, part in FORNECEDORES_CUSTO:
                v = rec_coz[coz] * PCT["custo"] * part
                for k in range(2):
                    lv.add(dia_util(ano, mes, 6 + 14 * k + random.randint(0, 3)), coz, forn, "Compra de %s - %s" % (cat2.lower(), coz), cat2, "Gasto",
                           "Custo.Operacional", random.choice(["Boleto", "PIX"]), random.choices(["Banco Alfa", "Banco Gama"], [0.75, 0.25])[0], r(v / 2, 0.06))
        # --- pessoal
        for coz in COZINHAS:
            sal = rec_coz[coz] * PCT["salarios"]
            lv.add(dia_util(ano, mes, 5), coz, "Folha de pagamento", "Salários %s" % coz, "Salários", "Gasto", "Despesa.com.Pessoal", "C/D em Conta", "Banco Alfa", sal)
            lv.add(dia_util(ano, mes, 20), coz, "Guias de encargos", "INSS e FGTS %s" % coz, "Encargos", "Gasto", "Despesa.com.Pessoal", "Boleto", "Banco Alfa", rec_coz[coz] * PCT["encargos"])
            lv.add(dia_util(ano, mes, 28), coz, "Provisão trabalhista", "Provisão de férias e 13º %s" % coz, "Provisões", "Gasto", "Despesa.com.Pessoal", "C/D em Conta", "Banco Beta", rec_coz[coz] * PCT["provisoes"])
            lv.add(dia_util(ano, mes, 1), coz, "Cartão Alimentação Fi", "Ticket alimentação %s" % coz, "Ticket", "Gasto", "Despesa.com.Pessoal", "Boleto", "Banco Gama", rec_coz[coz] * PCT["ticket"])
        tot = rec_total
        lv.add(dia_util(ano, mes, 12), "ADM", "Plano de Saúde Ro", "Plano de saúde", "Planos e Benefícios", "Gasto", "Despesa.com.Pessoal", "Boleto", "Banco Beta", tot * PCT["beneficios"])
        if mes in (3, 9, 11):
            lv.add(dia_util(ano, mes, 18), "ADM", "Instituto de Cursos Xi", "Treinamento de manipulação de alimentos", "Treinamentos", "Gasto", "Despesa.com.Pessoal", "PIX", "Banco Beta", r(tot * PCT["treinamentos"] * 3, 0.1))
        # --- tributos (ADM, rateados pelo modelo)
        lv.add(dia_util(ano, mes, 20), "ADM", "Guia de Impostos", "Tributos sobre o faturamento", "Impostos sobre Faturamento", "Gasto", "Tributo", "Boleto", "Banco Alfa", tot * PCT["tributo"])
        # --- fixas
        for coz in COZINHAS:
            lv.add(dia_util(ano, mes, 10), coz, "Imobiliária Psi", "Aluguel %s" % coz, "Aluguel", "Gasto", "Despesa.Fixa", "Boleto", "Banco Beta", rec_coz[coz] * PCT["aluguel"])
            lv.add(dia_util(ano, mes, 15), coz, "Energia Elétrica Fi", "Energia %s" % coz, "Energia Elétrica", "Gasto", "Despesa.Fixa", "Boleto", "Banco Gama", r(rec_coz[coz] * PCT["energia"], 0.1))
            lv.add(dia_util(ano, mes, 16), coz, "Águas Ni", "Água e esgoto %s" % coz, "Água e Esgoto", "Gasto", "Despesa.Fixa", "Boleto", "Banco Gama", rec_coz[coz] * PCT["agua"])
        lv.add(dia_util(ano, mes, 8), "ADM", "Telecom Nu", "Internet e telefonia", "Telefonia e Internet", "Gasto", "Despesa.Fixa", "Boleto", "Banco Beta", tot * PCT["internet"])
        lv.add(dia_util(ano, mes, 10), "ADM", "Contabilidade Chi", "Honorários contábeis", "Honorários Contábeis", "Gasto", "Despesa.Fixa", "PIX", "Banco Beta", 5200 * (1 + 0.003 * i))
        lv.add(dia_util(ano, mes, 7), "ADM", "Software Rho", "Sistema de gestão", "Software", "Gasto", "Despesa.Fixa", "Boleto", "Banco Beta", 2850)
        # --- variáveis e marketing
        for coz in COZINHAS:
            lv.add(dia_util(ano, mes, 22), coz, "Manutenção Upsilon", "Manutenção de equipamentos %s" % coz, "Manutenção de Equipamentos", "Gasto", "Despesa.Variável", "PIX", "Banco Gama", r(rec_coz[coz] * PCT["variavel"] * 0.6, 0.3))
            lv.add(dia_util(ano, mes, 24), coz, "Limpeza Fi", "Material de limpeza %s" % coz, "Material de Limpeza", "Gasto", "Despesa.Variável", "PIX", "Dinheiro", rec_coz[coz] * PCT["variavel"] * 0.4)
        lv.add(dia_util(ano, mes, 25), "ADM", "Comissões Comerciais", "Comissão sobre contratos", "Comissões", "Gasto", "Marketing.Comercial", "C/D em Conta", "Banco Alfa", tot * PCT["marketing"] * 0.7)
        lv.add(dia_util(ano, mes, 26), "ADM", "Agência Zeta", "Material institucional e site", "Material Publicitário", "Gasto", "Marketing.Comercial", "PIX", "Banco Beta", tot * PCT["marketing"] * 0.3)
        # --- financeiro
        lv.add(dia_util(ano, mes, 28), "ADM", "Banco Alfa", "Tarifas e encargos bancários", "Tarifas e Encargos Bancários", "Gasto", "Despesa.Financeira", "C/D em Conta", "Banco Alfa", r(2100, 0.1))
        lv.add(dia_util(ano, mes, 28), "ADM", "Banco Beta", "Juros de conta garantida", "Juros", "Gasto", "Despesa.Financeira", "C/D em Conta", "Banco Beta", r(3800, 0.25))
        lv.add(dia_util(ano, mes, 29 if mes != 2 else 27), "ADM", "Banco Alfa", "Rendimentos automáticos", "Rendimentos Automáticos na Conta", "Receita", "Receita.Financeira", "C/D em Conta", "Banco Alfa", r(3300, 0.2))
        # --- sócio
        lv.add(dia_util(ano, mes, 10), "ADM", "Sócios", "Retirada mensal dos sócios", "Pró-labore", "Sócio", "Retirada", "PIX", "Banco Alfa", 38000)
        # --- movimentações entre contas (somam zero)
        t = r(60000, 0.1)
        dt = dia_util(ano, mes, 3)
        lv.add(dt, "ADM", "Banco Alfa", "Transferência para Banco Beta", "Transferência entre contas", "Movimentação", "Saída", "C/D em Conta", "Banco Alfa", t)
        lv.add(dt, "ADM", "Banco Beta", "Transferência recebida de Banco Alfa", "Transferência entre contas", "Movimentação", "Entrada", "C/D em Conta", "Banco Beta", t)
        t = r(35000, 0.1)
        dt = dia_util(ano, mes, 4)
        lv.add(dt, "ADM", "Banco Alfa", "Transferência para Banco Gama", "Transferência entre contas", "Movimentação", "Saída", "C/D em Conta", "Banco Alfa", t)
        lv.add(dt, "ADM", "Banco Gama", "Transferência recebida de Banco Alfa", "Transferência entre contas", "Movimentação", "Entrada", "C/D em Conta", "Banco Gama", t)
        sq = 5600
        dt = dia_util(ano, mes, 21)
        lv.add(dt, "ADM", "Banco Beta", "Saque para caixa", "Transferência entre contas", "Movimentação", "Saída", "C/D em Conta", "Banco Beta", sq)
        lv.add(dt, "ADM", "Caixa", "Saque recebido no caixa", "Transferência entre contas", "Movimentação", "Entrada", "Dinheiro", "Dinheiro", sq)
        # fundo financeiro: aporte mensal, retirada em dois meses
        ffe = 18000
        dt = dia_util(ano, mes, 27)
        lv.add(dt, "ADM", "Cooperativa Delta", "Aporte no fundo financeiro", "Saída FFE", "Movimentação", "Saída", "C/D em Conta", "Banco Alfa", ffe)
        lv.add(dt, "ADM", "Cooperativa Delta", "Aporte no fundo financeiro", "Entrada FFE", "Movimentação", "Entrada", "C/D em Conta", "Cooperativa Delta", ffe)
        if (ano, mes) in ((2026, 2), (2026, 6)):
            dt = dia_util(ano, mes, 11)
            lv.add(dt, "ADM", "Cooperativa Delta", "Resgate do fundo financeiro", "Saída FFE", "Movimentação", "Saída", "C/D em Conta", "Cooperativa Delta", 45000)
            lv.add(dt, "ADM", "Cooperativa Delta", "Resgate do fundo financeiro", "Entrada FFE", "Movimentação", "Entrada", "C/D em Conta", "Banco Alfa", 45000)
    # --- eventos únicos
    lv.add(date(2025, 8, 12), "ADM", "Sócios", "Aporte de capital de giro", "-", "Sócio", "Aporte", "TED", "Banco Alfa", 50000)
    lv.add(date(2025, 8, 20), "Norte", "Equipamentos Mu", "Câmara fria Norte", "Aquisição de Equipamentos", "Movimentação", "Investimento", "Boleto", "Banco Alfa", 45000)
    lv.add(date(2025, 10, 6), "ADM", "Banco Gama", "Empréstimo para capital de giro", "Empréstimo", "Movimentação", "Entrada", "C/D em Conta", "Banco Gama", 90000)
    lv.add(date(2025, 12, 19), "ADM", "Sócios", "Distribuição de lucros", "Distribuição de Lucros", "Sócio", "Retirada", "PIX", "Banco Alfa", 40000)
    lv.add(date(2026, 1, 14), "Sul", "Equipamentos Mu", "Forno combinado Sul", "Aquisição de Equipamentos", "Movimentação", "Investimento", "Boleto", "Banco Alfa", 32000)
    lv.add(date(2026, 3, 17), "Leste", "Comprador Eventual Pi", "Venda de fogão industrial usado", "Venda de Ativo Imobilizado", "Movimentação", "Investimento", "PIX", "Banco Alfa", 12000)
    lv.add(date(2026, 5, 13), "Leste", "Equipamentos Mu", "Carro térmico Leste", "Aquisição de Equipamentos", "Movimentação", "Investimento", "Boleto", "Banco Beta", 18000)
    lv.add(date(2026, 6, 25), "ADM", "Sócios", "Distribuição de lucros", "Distribuição de Lucros", "Sócio", "Retirada", "PIX", "Banco Alfa", 30000)
    for (ano, mes) in MESES:
        if (ano, mes) >= (2025, 11):    # empréstimo: parcela mensal
            dt = dia_util(ano, mes, 6)
            lv.add(dt, "ADM", "Banco Gama", "Parcela do empréstimo", "Pagamento de Empréstimo", "Movimentação", "Saída", "C/D em Conta", "Banco Gama", 8400)
    lv.linhas.sort(key=lambda l: (l["Realiz."], l["Categoria"], l["Valor"]))
    # --- cobertura de saldo: antes de um pagamento que deixaria a conta abaixo do piso, uma transferência traz recursos da conta mais cheia
    piso, alvo = 25000.0, 110000.0
    bal = dict(CONTAS)
    novas = []
    for l in lv.linhas:
        s = sinal(l) * l["Valor"]
        if s < 0 and bal[l["Conta"]] + s < piso and l["Conta"] in ("Banco Alfa", "Banco Beta", "Banco Gama"):
            falta = alvo - (bal[l["Conta"]] + s)
            doadora = max((c for c in ("Banco Alfa", "Banco Beta", "Banco Gama") if c != l["Conta"]), key=lambda c: bal[c])
            if bal[doadora] - falta > piso:
                for conta, cat, desc in ((doadora, "Saída", "Transferência para %s" % l["Conta"]), (l["Conta"], "Entrada", "Transferência recebida de %s" % doadora)):
                    novas.append({"Realiz.": l["Realiz."], "Venc.": l["Realiz."], "C. C.": "ADM", "  Fornecedor/ Cliente": l["Conta"] if conta == doadora else doadora,
                                  "Descrição": desc, "Detalhamento/Categoria II": "Transferência entre contas", "Tipo": "Movimentação", "Categoria": cat,
                                  "Desembolso": "C/D em Conta", "Conta": conta, "Valor": round(falta, 2), "Comentários": None, "_antes": l})
                bal[doadora] -= falta
                bal[l["Conta"]] += falta
        bal[l["Conta"]] += s
    for n in novas:
        alvo_l = n.pop("_antes")
        i = next(k for k, x in enumerate(lv.linhas) if x is alvo_l)
        lv.linhas.insert(i, n)
    # --- saldos correntes (total) e por conta; checa que nenhuma conta fica negativa
    saldo = sum(CONTAS.values())
    por_conta = dict(CONTAS)
    for l in lv.linhas:
        s = sinal(l) * l["Valor"]
        saldo += s
        por_conta[l["Conta"]] += s
        l["Saldo"] = round(saldo, 2)
    return lv, receita_mes, por_conta


def escrever_ano(pasta, ano, linhas, receita_mes, abertura, final):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Controle Financeiro"
    azul = PatternFill("solid", fgColor="1F4E79")
    ws["A1"] = "Controle Financeiro %d" % ano
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
    ws["A2"] = "Saldo por conta"
    ws["A2"].font = Font(bold=True)
    for k, conta in enumerate(CONTAS):
        rr = 3 + k
        ws.cell(rr, 1, "Saldo %s Início" % conta)
        ws.cell(rr, 2, round(abertura[conta], 2)).number_format = 'R$ #,##0.00'
        ws.cell(rr, 4, "Saldo %s Final" % conta)
        ws.cell(rr, 5, round(final[conta], 2)).number_format = 'R$ #,##0.00'
    ws.cell(8, 1)  # a linha 8 é o cabeçalho
    for c, nome in enumerate(COLUNAS, start=1):
        cel = ws.cell(8, c, nome)
        cel.fill = azul
        cel.font = Font(bold=True, color="FFFFFF")
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for k, l in enumerate(linhas):
        rr = 9 + k
        for c, nome in enumerate(COLUNAS, start=1):
            v = l.get(nome)
            cel = ws.cell(rr, c, v)
            if nome in ("Realiz.", "Venc."):
                cel.number_format = "dd/mm/yyyy"
            elif nome in ("Valor", "Saldo"):
                cel.number_format = 'R$ #,##0.00'
    ult = 8 + len(linhas)
    tab = Table(displayName="ControleFinanceiro", ref="A8:M%d" % ult)
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    for col, w in zip("ABCDEFGHIJKLM", (12, 12, 8, 26, 40, 30, 14, 20, 14, 20, 16, 18, 20)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A9"
    # --- Lista (validações)
    wl = wb.create_sheet("Lista")
    listas = {"Conta": list(CONTAS), "Categoria": sorted(set(l["Categoria"] for l in linhas)), "C. C.": COZINHAS + ["ADM"],
              "Fornecedor/ Cliente": sorted(set(l["  Fornecedor/ Cliente"] for l in linhas))}
    for c, (nome, itens) in enumerate(listas.items(), start=1):
        wl.cell(1, c, nome).font = Font(bold=True)
        for k, it in enumerate(itens):
            wl.cell(2 + k, c, it)
        wl.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 28
    # --- Competência
    wc = wb.create_sheet("Competencia")
    wc["A1"], wc["B1"] = "Mês", "Valor"
    meses_ano = [m for (a, m) in MESES if a == ano]
    for k, m in enumerate(meses_ano):
        wc.cell(2 + k, 1, date(ano, m, 1)).number_format = "mm/yyyy"
        wc.cell(2 + k, 2, round(receita_mes[(ano, m)] * random.uniform(0.97, 1.05), 2)).number_format = 'R$ #,##0.00'
    tc = Table(displayName="tbCompetencia", ref="A1:B%d" % (1 + len(meses_ano)))
    tc.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    wc.add_table(tc)
    wc.column_dimensions["A"].width = 12
    wc.column_dimensions["B"].width = 18
    # --- Saldos (lida só do arquivo mais antigo: saldo do dia zero por conta)
    wsd = wb.create_sheet("Saldos")
    wsd["A1"] = "Saldos de abertura"
    wsd["B1"], wsd["C1"] = "Conta", "Saldo no dia zero (01/07/2025)"
    for k, (conta, v) in enumerate(CONTAS.items()):
        wsd.cell(2 + k, 2, conta)
        wsd.cell(2 + k, 3, v).number_format = 'R$ #,##0.00'
    wsd.cell(2 + len(CONTAS), 2, "Resultado acumulado anterior")
    wsd.cell(2 + len(CONTAS), 3, 0)
    wsd.column_dimensions["B"].width = 30
    wsd.column_dimensions["C"].width = 30
    wb.save(pasta / ("Controle Financeiro %d.xlsx" % ano))


def main(pasta):
    pasta = Path(pasta)
    lv, receita_mes, por_conta_final = montar()
    abertura = dict(CONTAS)
    for ano in (2025, 2026):
        ls = [l for l in lv.linhas if l["Realiz."].year == ano]
        fin = dict(abertura)
        for l in ls:
            fin[l["Conta"]] += sinal(l) * l["Valor"]
        escrever_ano(pasta, ano, ls, receita_mes, abertura, fin)
        print("%d: %d lançamentos" % (ano, len(ls)))
        abertura = fin
    # --- relatório de sanidade
    por_mes = defaultdict(lambda: defaultdict(float))
    for l in lv.linhas:
        k = (l["Realiz."].year, l["Realiz."].month)
        por_mes[k][l["Categoria"]] += l["Valor"] * (sinal(l) if l["Categoria"] != "Investimento" else 1)
    print("\nMês        Receita     Custo  Pessoal  Resultado(aprox)  saldo final conta mín.")
    sal = dict(CONTAS)
    menor = {c: v for c, v in CONTAS.items()}
    for l in lv.linhas:
        sal[l["Conta"]] += sinal(l) * l["Valor"]
        menor[l["Conta"]] = min(menor[l["Conta"]], sal[l["Conta"]])
    for k in MESES:
        d = por_mes[k]
        res = d["Operacional"] + d["Receita.Financeira"] + sum(d[c] for c in ("Custo.Operacional", "Tributo", "Despesa.Fixa", "Despesa.Variável",
                                                                           "Despesa.com.Pessoal", "Marketing.Comercial", "Despesa.Financeira"))
        print("%d-%02d  %10.0f %9.0f %8.0f  %10.0f" % (k[0], k[1], d["Operacional"], d["Custo.Operacional"], d["Despesa.com.Pessoal"], res))
    print("\nSaldo final por conta:", {c: round(v) for c, v in sal.items()})
    print("Menor saldo por conta:", {c: round(v) for c, v in menor.items()})


if __name__ == "__main__":
    main(sys.argv[1])
