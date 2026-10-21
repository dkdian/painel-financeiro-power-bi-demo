"""Escreve as páginas do painel do AnáliseFinanceira.pbix do cliente direto no relatório PBIR de dentro do .pbix.
Padrão BI: base paginas_cliente_a.py (PLANO G, 09/10/2026): faixa azul com abas, gráficos na frente das tabelas, aba Lançamentos.
Módulos do cliente (C.C., competência, grupos de Pessoal) vêm de ../config_cliente.py; páginas e abas se ajustam.

Método do cliente B (paginas_cliente_b.py): Power BI Desktop FECHADO; copia todas as entradas do .pbix byte a byte (o DataModel não muda),
troca só Report/definition/pages/**, remove SecurityBindings (o Desktop recusa .pbix editado por fora com ela) e valida antes de gravar.
Moldes (slicer, forma, caixas de texto, colunas, linha, pizza) e as 4 matrizes (DRE, DRE por cozinha, DRE Alternativa, Fluxo de Caixa
Mensal) vêm do backup `AnáliseFinanceira_20261009_antes_planoG.pbix`: as matrizes são copiadas como estão e só reposicionadas.

Uso:  python paginas_painel.py [--gravar]     (sem --gravar só valida e mostra o resumo; o arquivo vai para %TEMP%, nunca sobre o original)
"""
import argparse
import copy
import datetime
import hashlib
import json
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8")
AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(AQUI))        # config_cliente.py fica em 02_Automacao
import config_cliente as cfg
BASE = os.path.dirname(os.path.dirname(AQUI))                 # ...\00_Painel BI
PBIX = os.path.join(BASE, cfg.PBIX + ".pbix")
MODELOS_PBIX = os.path.join(AQUI, cfg.PAINEL_MOLDES)   # zip só com Report/definition/pages/** (slicer, forma, caixas, colunas, linha, pizza, 4 matrizes); sem dados
TMDL = os.path.join(BASE, "03_Documentacao", "gerado")
SCHEMA_VISUAL = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.12.0/schema.json"
SCHEMA_PAGE = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json"
SCHEMA_PAGES = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json"
CF = "ControleFinanceiro"
AZUL, AZUL_ESCURO, TEXTO = "#094E8C", "#16263A", "#16263A"
VERDE, VERMELHO = "#1E8E3E", "#C62828"
AZUIS_PIZZA = {"(-) Custo Operacional": "#094E8C", "(-) Despesa com Pessoal": "#2F78B7", "(-) Despesa Financeira": "#5E9BCF",
               "(-) Despesa Variável": "#8DB9E1", "(-) Despesa Fixa": "#B1D0EB", "(-) Marketing/Comercial": "#D0E3F4",
               "(-) Tributo": "#4F6D94", "Resultado": "#16263A"}
Y0, ALTURA_FAIXA = 166, 150                                     # topo do conteúdo e altura da faixa azul do cabeçalho (título, filtros, abas)
NOME_CC = cfg.ROTULO_CC                                          # "Cozinha", "Loja", "Obra"...
ABA_CLIENTES = "Clientes e " + cfg.ROTULO_CC_PLURAL if cfg.TEM_CC else "Clientes"
ABA_DRE_CC = "DRE por " + cfg.ROTULO_CC.lower()
ABA_EFICIENCIA = "Eficiência e pessoal"
# largura de cada aba na faixa (px, ajustada ao texto); só entram as abas dos módulos ligados
LARGURA_ABA = {"Receita": 90, "Resultado e gastos": 150, "DRE": 70, "DFC": 70, "Lançamentos": 120, ABA_CLIENTES: 160,
               ABA_EFICIENCIA: 160, ABA_DRE_CC: 130}
_ordem = ["Receita", ABA_CLIENTES, "Resultado e gastos"] + ([ABA_EFICIENCIA] if cfg.TEM_COMPETENCIA or cfg.TEM_GRUPOS_PESSOAL else [])     + ["DRE"] + ([ABA_DRE_CC] if cfg.TEM_CC else []) + ["DFC", "Lançamentos"]
ABAS = [(n, LARGURA_ABA[n]) for n in _ordem]                    # navegação do cabeçalho


def lit(v):
    return {"expr": {"Literal": {"Value": v}}}


def col(entidade, prop):
    return {"Column": {"Expression": {"SourceRef": {"Entity": entidade}}, "Property": prop}}


def med(entidade, prop):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": entidade}}, "Property": prop}}


def proj(campo, nome_exibicao=None, ativo=None):
    tipo = "Measure" if "Measure" in campo else "Column"
    ent = campo[tipo]["Expression"]["SourceRef"]["Entity"]
    prop = campo[tipo]["Property"]
    p = {"field": campo, "queryRef": "%s.%s" % (ent, prop), "nativeQueryRef": prop}
    if ativo is not None:
        p["active"] = ativo
    if nome_exibicao:
        p["displayName"] = nome_exibicao
    return p


# ---------------------------------------------------------------- modelo (para validar campos)
def catalogo_modelo(pasta_tmdl):
    """{(tabela, nome)} de colunas e medidas, lido do export TMDL do modelo (03_Documentacao\\gerado\\Modelo_TMDL_*, o mais recente)."""
    pasta = sorted(d for d in os.listdir(pasta_tmdl) if d.startswith("Modelo_TMDL_"))[-1]
    tabelas = os.path.join(pasta_tmdl, pasta, "tables")
    out = set()
    for f in os.listdir(tabelas):
        txt = open(os.path.join(tabelas, f), encoding="utf-8").read()
        m = re.search(r"^table (?:'([^']+)'|(\S+))", txt, flags=re.M)
        tab = m.group(1) or m.group(2)
        for k in re.finditer(r"^\t(?:column|measure|calculationItem) (?:'((?:[^']|'')+)'|([^\s=]+))", txt, flags=re.M):
            out.add((tab, (k.group(1) or k.group(2)).replace("''", "'")))
    return out, pasta


def campos_usados(obj, acc):
    if isinstance(obj, dict):
        for tipo in ("Column", "Measure"):
            if tipo in obj and isinstance(obj[tipo], dict) and "Property" in obj[tipo]:
                ent = obj[tipo].get("Expression", {}).get("SourceRef", {}).get("Entity")
                if ent:
                    acc.add((ent, obj[tipo]["Property"]))
        for v in obj.values():
            campos_usados(v, acc)
    elif isinstance(obj, list):
        for v in obj:
            campos_usados(v, acc)


def titulo_visual(d):
    try:
        return d["visual"]["visualContainerObjects"]["title"][0]["properties"]["text"]["expr"]["Literal"]["Value"].strip("'")
    except (KeyError, IndexError):
        return ""


# ---------------------------------------------------------------- blocos reaproveitados do relatório atual
class Modelos:
    """Visuais do .pbix antigo usados como molde (copiados e ajustados). Nada é inventado: só troca de campo, texto e posição."""
    def __init__(self, z):
        self.v = {}
        base = "Report/definition/pages/"
        for n in z.namelist():
            if n.startswith(base) and n.endswith("visual.json"):
                d = json.loads(z.read(n).decode("utf-8"))
                self.v.setdefault(d["visual"].get("visualType"), []).append(d)
        js = lambda d: json.dumps(d, ensure_ascii=False)
        self.slicer_periodo = next(d for d in self.v["slicer"] if "dCalendario" in js(d))
        self.slicer_texto = next(d for d in self.v["slicer"] if "Fornecedor" in js(d))
        self.forma = self.v["shape"][0]
        caixas = self.v["textbox"]
        self.caixa_titulo = next(d for d in caixas if 'fontSize": "18pt' in js(d))
        self.caixa_sub = next(d for d in caixas if "10pt" in js(d))
        self.colunas = next(d for d in self.v["clusteredColumnChart"] if "Receita_Operacional" in js(d))
        self.linha = next(d for d in self.v["lineChart"] if "Plot_Resultado_Acumulado" in js(d))
        self.pizza = next(d for d in self.v["pieChart"] if "Categoria_Graf" in js(d))
        self.matrizes = {titulo_visual(d): d for d in self.v["pivotTable"]}


def novo_id(*partes):
    return hashlib.md5("|".join(partes).encode("utf-8")).hexdigest()[:20]


def posicionar(d, x, y, w, h, z, tab):
    d["position"] = {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": tab}


def container_padrao(titulo=None, fundo=True):
    """visualContainerObjects de um cartão/gráfico: fundo branco, borda #D9E2EC, sombra leve e (opcional) título."""
    o = {"background": [{"properties": {"show": lit("true" if fundo else "false"), "color": {"solid": {"color": lit("'#FFFFFF'")}},
                                        "transparency": lit("0D")}}],
         "border": [{"properties": {"show": lit("true"), "color": {"solid": {"color": lit("'#D9E2EC'")}}, "radius": lit("10D"),
                                    "width": lit("1D")}}],
         "dropShadow": [{"properties": {"show": lit("true"), "color": {"solid": {"color": lit("'#16263A'")}}, "position": lit("'Outer'"),
                                        "preset": lit("'Bottom'"), "transparency": lit("88D"), "shadowBlur": lit("10D"),
                                        "shadowDistance": lit("2D"), "shadowSpread": lit("0D")}}],
         "padding": [{"properties": {"top": lit("8D"), "bottom": lit("8D"), "left": lit("12D"), "right": lit("12D")}}]}
    if titulo:
        o["title"] = [{"properties": {"show": lit("true"), "fontColor": {"solid": {"color": lit("'#16263A'")}}, "fontSize": lit("12D"),
                                      "bold": lit("true"), "fontFamily": lit("'Segoe UI Semibold'"), "text": lit("'%s'" % titulo)}}]
    return o


class Pagina:
    def __init__(self, nome_exib, mod, titulo, subtitulo=cfg.PAINEL_SUBTITULO):
        self.nome = novo_id("pagina", nome_exib)
        self.exib = nome_exib
        self.mod = mod
        self.visuais = {}
        self.z = 1000
        self._cabecalho(titulo, subtitulo)

    def _add(self, d, chave):
        d = copy.deepcopy(d)
        d["$schema"] = SCHEMA_VISUAL
        d["name"] = novo_id(self.nome, chave)
        self.visuais[d["name"]] = d
        return d

    def _cabecalho(self, titulo, subtitulo):
        f = self._add(self.mod.forma, "faixa")
        posicionar(f, 0, 0, 1280, ALTURA_FAIXA, 0, 1000)
        t = self._add(self.mod.caixa_titulo, "titulo")
        t["visual"]["objects"]["general"][0]["properties"]["paragraphs"][0]["textRuns"][0]["value"] = titulo
        posicionar(t, 24, 14, 384, 44, 2000, 3000)
        s = self._add(self.mod.caixa_sub, "subtitulo")
        s["visual"]["objects"]["general"][0]["properties"]["paragraphs"][0]["textRuns"][0]["value"] = subtitulo
        posicionar(s, 24, 58, 384, 26, 1000, 2000)
        self._carimbo(24, 82, 190, 24)
        self._limpar(1118, 32, 146, 30)
        x = 24
        for i, (nome, larg) in enumerate(ABAS):
            self._aba(i, nome, x, 110, larg, 36)
            x += larg + 8

    def _carimbo(self, x, y, w, h):
        """Data do último lançamento da planilha (medida Txt_Dados_Ate, sem filtros), para a gestão saber até quando vão os dados."""
        d = {"$schema": SCHEMA_VISUAL, "name": novo_id(self.nome, "carimbo"),
             "visual": {"visualType": "card",
                        "query": {"queryState": {"Values": {"projections": [proj(med(CF, "Txt_Dados_Ate"))]}}},
                        "objects": {"labels": [{"properties": {"fontSize": lit("10D"), "color": {"solid": {"color": lit("'#BFD4EA'")}},
                                                               "fontFamily": lit("'Segoe UI'")}}],
                                    "categoryLabels": [{"properties": {"show": lit("false")}}]},
                        "visualContainerObjects": {**{k: [{"properties": {"show": lit("false")}}]
                                                      for k in ("background", "border", "dropShadow", "title")},
                                                   "padding": [{"properties": {"top": lit("0D"), "bottom": lit("0D"), "left": lit("0D"),
                                                                               "right": lit("0D")}}]},
                        "drillFilterOtherVisuals": True}}
        posicionar(d, x, y, w, h, 5000, 5000)
        self.visuais[d["name"]] = d

    def _botao(self, chave, texto, x, y, w, h, z, fundo, transp, cor_texto, contorno=None, negrito=True, tamanho=10, link=None, dica=None):
        """Pastilha clicável em 3 camadas (a forma não mostra texto neste Desktop e o botão de ação não desenha preenchimento):
        forma arredondada com o preenchimento, caixa de texto centralizada por cima e, no topo, uma forma 100% transparente com
        a ação de clique (sem `link` a pastilha é só um rótulo, como a aba da página atual)."""
        b = self._add(self.mod.forma, chave)
        o = b["visual"]["objects"]
        o["shape"] = [{"properties": {"tileShape": lit("'rectangle'")}}]
        o["fill"] = [{"properties": {"show": lit("true"), "fillColor": {"solid": {"color": lit("'%s'" % fundo)}}, "transparency": lit("%dD" % transp)},
                      "selector": {"id": "default"}}]
        o["outline"] = [{"properties": {"show": lit("false")}}] if not contorno else \
            [{"properties": {"show": lit("true"), "lineColor": {"solid": {"color": lit("'%s'" % contorno)}}, "weight": lit("1D")}}]
        posicionar(b, x, y, w, h, z, z)
        t = self._add(self.mod.caixa_sub, chave + "_texto")
        par = t["visual"]["objects"]["general"][0]["properties"]["paragraphs"][0]
        par["horizontalTextAlignment"] = "center"
        par["textRuns"][0]["value"] = texto
        par["textRuns"][0]["textStyle"] = {"fontFamily": "Segoe UI Semibold" if negrito else "Segoe UI", "fontSize": "%dpt" % tamanho,
                                           "color": cor_texto, **({"fontWeight": "bold"} if negrito else {})}
        posicionar(t, x, y + (h - 16) // 2 - 2, w, h - (h - 16) // 2, z + 100, z + 100)
        if link:
            c = self._add(self.mod.forma, chave + "_clique")
            c["visual"]["objects"]["fill"] = [{"properties": {"show": lit("true"), "fillColor": {"solid": {"color": lit("'#FFFFFF'")}},
                                                              "transparency": lit("100D")}, "selector": {"id": "default"}}]
            c["visual"]["objects"]["outline"] = [{"properties": {"show": lit("false")}}]
            c["visual"]["visualContainerObjects"] = {"visualLink": [{"properties": dict(show=lit("true"), tooltip=lit("'%s'" % dica), **link)}]}
            posicionar(c, x, y, w, h, z + 200, z + 200)
        return b

    def _limpar(self, x, y, w, h):
        """Pastilha de limpar filtros no canto superior direito da faixa: ✕ + texto, contorno branco sobre o azul."""
        self._botao("limpar", "✕  Limpar filtros", x, y, w, h, 8000, "#FFFFFF", 90, "#FFFFFF", contorno="#FFFFFF",
                    link={"type": lit("'ClearAllSlicers'")}, dica="Limpa os filtros do painel")

    def _aba(self, i, nome, x, y, w, h):
        """Aba de navegação entre as páginas: a da página atual fica branca (selecionada); as outras, translúcidas."""
        atual = nome == self.exib
        self._botao("aba_%d" % i, nome, x, y, w, h, 7500 + i, "#FFFFFF", 0 if atual else 85, "#094E8C" if atual else "#FFFFFF",
                    negrito=atual, link={"type": lit("'PageNavigation'"), "navigationSection": lit("'%s'" % novo_id("pagina", nome))},
                    dica="Ir para " + nome)

    def slicer(self, chave, entidade, prop, rotulo, x, grupo, modelo="texto", largura=212, ordem=7000, y=14, h=66, corpo=False):
        """Segmentador em lista suspensa. corpo=True: sobre o fundo claro da página (cabeçalho escuro, caixa branca com borda)."""
        base = self.mod.slicer_periodo if modelo == "periodo" else self.mod.slicer_texto
        d = self._add(base, "slicer_" + chave)
        d["visual"]["query"]["queryState"]["Values"]["projections"] = [proj(col(entidade, prop), rotulo, True)]
        d["visual"]["syncGroup"] = {"groupName": grupo, "fieldChanges": True, "filterChanges": True}
        if corpo:
            d["visual"]["objects"]["header"][0]["properties"]["fontColor"] = {"solid": {"color": lit("'#4A5B70'")}}
            vc = d["visual"]["visualContainerObjects"]
            vc["background"] = [{"properties": {"show": lit("true"), "color": {"solid": {"color": lit("'#FFFFFF'")}}, "transparency": lit("0D")}}]
            vc["border"] = [{"properties": {"show": lit("true"), "color": {"solid": {"color": lit("'#D9E2EC'")}}, "radius": lit("8D"),
                                            "width": lit("1D")}}]
        posicionar(d, x, y, largura, h, ordem, ordem)
        return d

    def filtros_padrao(self):
        """Faixa azul: os filtros de sempre (Período, C.C. se o cliente usa, Cliente / Fornecedor), sincronizados entre as páginas."""
        self.slicer("periodo", "dCalendario", "Ano_Mês", "Período", 424, "Período", modelo="periodo", largura=216)
        if cfg.TEM_CC:
            self.slicer("cozinha", CF, cfg.COL_CC, NOME_CC, 652, cfg.COL_CC, largura=216)
        self.slicer("cliente", CF, cfg.COL_CLIENTE, "Cliente / Fornecedor", 880, cfg.COL_CLIENTE, largura=224)

    def cartao(self, chave, titulo, medida, x, y, w, h, medida_cor=None, subtitulo=None):
        d = {"$schema": SCHEMA_VISUAL, "name": novo_id(self.nome, "cartao_" + chave),
             "visual": {"visualType": "card",
                        "query": {"queryState": {"Values": {"projections": [proj(med(CF, medida))]}}},
                        "objects": {"labels": [{"properties": {"fontSize": lit("20D"), "color": {"solid": {"color": lit("'#094E8C'")}},
                                                               "fontFamily": lit("'Segoe UI Semibold'"), "labelDisplayUnits": lit("1D")}}],
                                    "categoryLabels": [{"properties": {"show": lit("false")}}]},
                        "visualContainerObjects": container_padrao(titulo),
                        "drillFilterOtherVisuals": True}}
        if medida_cor:
            d["visual"]["objects"]["labels"][0]["properties"]["color"] = {"solid": {"color": {"expr": med(CF, medida_cor)}}}
        d["visual"]["visualContainerObjects"]["padding"] = [{"properties": {"top": lit("4D"), "bottom": lit("4D"), "left": lit("12D"), "right": lit("12D")}}]
        if subtitulo:      # medida de texto: diz, sob o título, a que período o número se refere
            self.subtitulo_med(d, subtitulo)
        posicionar(d, x, y, w, h, 3000 + len(self.visuais), 3000 + len(self.visuais))
        self.visuais[d["name"]] = d
        return d

    @staticmethod
    def subtitulo(d, texto):
        """Linha cinza sob o título do visual dizendo o que o gráfico mostra."""
        d["visual"]["visualContainerObjects"]["subTitle"] = [{"properties": {
            "show": lit("true"), "text": lit("'%s'" % texto), "fontColor": {"solid": {"color": lit("'#5B6B7F'")}},
            "fontSize": lit("9D"), "fontFamily": lit("'Segoe UI'")}}]

    @staticmethod
    def subtitulo_med(d, medida):
        """Subtítulo dinâmico: medida de texto (ex.: 'Mês de referência: Set/26')."""
        d["visual"]["visualContainerObjects"]["subTitle"] = [{"properties": {
            "show": lit("true"), "text": {"expr": med(CF, medida)}, "fontColor": {"solid": {"color": lit("'#5B6B7F'")}},
            "fontSize": lit("9D"), "fontFamily": lit("'Segoe UI'")}}]

    def _z(self):
        return 4000 + len(self.visuais)

    def colunas(self, chave, titulo, medida, x, y, w, h, tooltip=None, unidade="1000000D", precisao="2L", cor=AZUL, medida_cor=None,
                categoria=("dEixoMes", "Eixo_Mes"), entidade=CF, medida_cor_campo=None):
        d = self._add(self.mod.colunas, "col_" + chave)
        q = d["visual"]["query"]["queryState"]
        q["Category"]["projections"] = [proj(col(*categoria), None, True)]
        q["Y"]["projections"] = [proj(med(entidade, medida))]
        if tooltip:
            q["Tooltips"] = {"projections": [proj(med(CF, t)) for t in tooltip]}
        d["visual"]["query"]["sortDefinition"] = {"sort": [{"field": col(*categoria), "direction": "Ascending"}]}
        ob = d["visual"]["objects"]
        ob["labels"][0]["properties"]["labelDisplayUnits"] = lit(unidade)
        ob["labels"][0]["properties"]["labelPrecision"] = lit(precisao)
        ob["labels"][0]["properties"]["color"] = {"solid": {"color": lit("'#000000'")}}
        # cor por sinal do próprio valor da barra (regra Conditional, igual à fonte vermelha das matrizes)
        corv = {"expr": {"Conditional": {"Cases": [{"Condition": {"Comparison": {"ComparisonKind": 3, "Left": med(entidade, medida),
                                                                                 "Right": {"Literal": {"Value": "0D"}}}},
                                                    "Value": {"Literal": {"Value": "'%s'" % VERMELHO}}}],
                                         "DefaultValue": {"Literal": {"Value": "'%s'" % VERDE}}}}} if medida_cor else lit("'%s'" % cor)
        dp = {"properties": {"fill": {"solid": {"color": corv}}}}
        if medida_cor:
            dp["selector"] = {"data": [{"dataViewWildcard": {"matchingOption": 1}}], "metadata": "%s.%s" % (entidade, medida)}
        if medida_cor_campo:     # a cor de cada barra vem de uma medida de texto (#hex), como a cor dos números dos cartões
            dp = {"properties": {"fill": {"solid": {"color": {"expr": med(CF, medida_cor_campo)}}}},
                  "selector": {"metadata": "%s.%s" % (entidade, medida)}}
        ob["dataPoint"] = [dp]
        ob["legend"] = [{"properties": {"show": lit("false")}}]
        ob["valueAxis"][0]["properties"]["labelDisplayUnits"] = lit(unidade)
        ob["valueAxis"][0]["properties"]["showAxisTitle"] = lit("false")
        ob["valueAxis"][0]["properties"]["labelPrecision"] = lit("0L" if unidade in ("1000D", "1D") else "1L")
        ob["categoryAxis"] = [{"properties": {"showAxisTitle": lit("false")}}]
        d["visual"]["visualContainerObjects"]["title"][0]["properties"]["text"] = lit("'%s'" % titulo)
        posicionar(d, x, y, w, h, self._z(), self._z())
        return d

    def colunas_duas(self, chave, titulo, m1, nome1, m2, nome2, x, y, w, h, unidade="1000000D", precisao="2L", cores=(AZUL, "#8DB9E1"), rotulos=False):
        """Colunas agrupadas com duas medidas (Receita Caixa × Receita Competência), legenda no topo."""
        d = self.colunas(chave, titulo, m1, x, y, w, h, unidade=unidade, precisao=precisao)
        d["visual"]["query"]["queryState"]["Y"]["projections"] = [proj(med(CF, m1), nome1), proj(med(CF, m2), nome2)]
        ob = d["visual"]["objects"]
        ob["dataPoint"] = [{"properties": {"fill": {"solid": {"color": lit("'%s'" % c)}}}, "selector": {"metadata": "%s.%s" % (CF, m)}}
                           for m, c in ((m1, cores[0]), (m2, cores[1]))]
        ob["legend"] = [{"properties": {"show": lit("true"), "position": lit("'Top'"), "showTitle": lit("false"), "fontSize": lit("10D")}}]
        if rotulos:
            ob["labels"][0]["properties"]["fontSize"] = lit("8D")
        else:
            ob["labels"] = [{"properties": {"show": lit("false")}}]      # 24 rótulos juntos poluem: o valor aparece ao passar o mouse
        return d

    def pareto(self, chave, titulo, x, y, w, h):
        """Pareto de clientes: colunas de receita (do maior para o menor) + linha do % acumulado no eixo secundário. Vale para o Período
        selecionado; sem seleção, para o mês de referência."""
        d = self.colunas(chave, titulo, "Receita_Operacional_Mes", x, y, w, h, categoria=(CF, cfg.COL_CLIENTE), unidade="0D", precisao="0L")
        d["visual"]["visualType"] = "lineClusteredColumnComboChart"
        q = d["visual"]["query"]["queryState"]
        q["Y"]["projections"] = [proj(med(CF, "Receita_Operacional_Mes"), "Receita")]
        q["Y2"] = {"projections": [proj(med(CF, "Painel_Pareto_Acum"), "% acumulado")]}
        d["visual"]["query"]["sortDefinition"] = {"sort": [{"field": med(CF, "Receita_Operacional_Mes"), "direction": "Descending"}]}
        ob = d["visual"]["objects"]
        ob["dataPoint"] = [{"properties": {"fill": {"solid": {"color": lit("'%s'" % AZUL)}}}, "selector": {"metadata": "%s.Receita_Operacional_Mes" % CF}},
                           {"properties": {"fill": {"solid": {"color": lit("'#ED7D31'")}}}, "selector": {"metadata": "%s.Painel_Pareto_Acum" % CF}}]
        ob["labels"][0]["properties"]["fontSize"] = lit("8D")
        ob["labels"][0]["properties"].pop("labelPrecision", None)
        ob["valueAxis"][0]["properties"]["labelDisplayUnits"] = lit("0D")
        ob["valueAxis"][0]["properties"].pop("labelPrecision", None)
        ob["valueAxis"][0]["properties"]["secShow"] = lit("true")
        ob["legend"] = [{"properties": {"show": lit("true"), "position": lit("'Top'"), "showTitle": lit("false"), "fontSize": lit("10D")}}]
        return d

    def linha(self, chave, titulo, medida, x, y, w, h, unidade="1D", precisao="1L", cor=AZUL, eixo=None, categoria=("dEixoMes", "Eixo_Mes")):
        """Linha de uma medida. eixo=(início, fim) fixa a escala (ex.: (0, 1) para 0–100%); None deixa automático."""
        d = self._add(self.mod.linha, "lin_" + chave)
        q = d["visual"]["query"]["queryState"]
        q["Category"]["projections"] = [proj(col(*categoria), None, True)]
        q["Y"]["projections"] = [proj(med(CF, medida))]
        d["visual"]["query"]["sortDefinition"] = {"sort": [{"field": col(*categoria), "direction": "Ascending"}]}
        ob = d["visual"]["objects"]
        ob["labels"][0]["properties"]["color"] = {"solid": {"color": lit("'#000000'")}}
        ob["labels"][0]["properties"]["labelDisplayUnits"] = lit(unidade)
        ob["labels"][0]["properties"]["labelPrecision"] = lit(precisao)
        ob["dataPoint"] = [{"properties": {"fill": {"solid": {"color": lit("'%s'" % cor)}}},
                            "selector": {"metadata": "%s.%s" % (CF, medida)}}]
        ob["legend"] = [{"properties": {"show": lit("false")}}]
        va = ob["valueAxis"][0]["properties"]
        va["labelDisplayUnits"] = lit(unidade)
        va["showAxisTitle"] = lit("false")
        ob["categoryAxis"] = [{"properties": {"showAxisTitle": lit("false")}}]
        if eixo:
            va["start"] = lit("%sD" % eixo[0])
            va["end"] = lit("%sD" % eixo[1])
        d["visual"]["visualContainerObjects"]["title"][0]["properties"]["text"] = lit("'%s'" % titulo)
        posicionar(d, x, y, w, h, self._z(), self._z())
        return d

    def linhas(self, chave, titulo, medidas, x, y, w, h, unidade="1000D", precisao="0L"):
        """Várias linhas no mesmo gráfico: medidas = [(medida, nome na legenda, cor)]."""
        d = self.linha(chave, titulo, medidas[0][0], x, y, w, h, unidade=unidade, precisao=precisao)
        d["visual"]["query"]["queryState"]["Y"]["projections"] = [proj(med(CF, m), nome) for m, nome, _ in medidas]
        ob = d["visual"]["objects"]
        ob["dataPoint"] = [{"properties": {"fill": {"solid": {"color": lit("'%s'" % c)}}}, "selector": {"metadata": "%s.%s" % (CF, m)}}
                           for m, _, c in medidas]
        ob["legend"] = [{"properties": {"show": lit("true"), "position": lit("'Top'"), "showTitle": lit("false"), "fontSize": lit("10D")}}]
        ob["labels"][0]["properties"]["fontSize"] = lit("8D")
        return d

    def pizza(self, chave, titulo, categoria, medida, x, y, w, h, cores=None, rotulo="Percent of total", ordenar_por_valor=True):
        """Pizza: categoria = (entidade, coluna); medida = (entidade, medida). cores = {categoria: cor} (senão, as cores do tema)."""
        d = self._add(self.mod.pizza, "piz_" + chave)
        q = d["visual"]["query"]["queryState"]
        q.pop("Series", None)
        q["Category"]["projections"] = [proj(col(*categoria), None, True)]
        q["Y"]["projections"] = [proj(med(*medida))]
        d["visual"]["query"]["sortDefinition"] = {"sort": [{"field": med(*medida) if ordenar_por_valor else col(*categoria),
                                                            "direction": "Descending" if ordenar_por_valor else "Ascending"}]}
        d.pop("filterConfig", None)
        d["visual"]["visualContainerObjects"]["title"][0]["properties"]["text"] = lit("'%s'" % titulo)
        ob = {"legend": [{"properties": {"show": lit("true"), "position": lit("'RightCenter'"), "showTitle": lit("false"), "fontSize": lit("10D")}}],
              "labels": [{"properties": {"show": lit("true"), "labelStyle": lit("'%s'" % rotulo), "labelPrecision": lit("1L"),
                                         "fontSize": lit("9D"), "color": {"solid": {"color": lit("'#000000'")}}}}]}
        if cores:
            ob["dataPoint"] = [{"properties": {"fill": {"solid": {"color": lit("'%s'" % cor)}}},
                                "selector": {"data": [{"scopeId": {"Comparison": {"ComparisonKind": 0, "Left": col(*categoria),
                                                                                   "Right": {"Literal": {"Value": "'%s'" % cat}}}}}]}}
                               for cat, cor in cores.items()]
        d["visual"]["objects"] = ob
        posicionar(d, x, y, w, h, self._z(), self._z())
        return d

    def filtro_in(self, d, chave, campo, valores):
        """Filtro de visual: campo IN valores (texto)."""
        ent = campo["Column"]["Expression"]["SourceRef"]["Entity"]
        prop = campo["Column"]["Property"]
        expr = {"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": prop}}
        cond = {"In": {"Expressions": [expr], "Values": [[{"Literal": {"Value": "'%s'" % v}}] for v in valores]}}
        d.setdefault("filterConfig", {"filters": []})["filters"].append(
            {"name": novo_id(self.nome, chave), "field": campo, "type": "Categorical",
             "filter": {"Version": 2, "From": [{"Name": "t", "Entity": ent, "Type": 0}], "Where": [{"Condition": cond}]}})

    def matriz(self, titulo_molde, chave, x, y, w, h):
        """Matriz já ajustada em 08/10 (DRE, DRE por cozinha, DRE Alternativa, Fluxo de Caixa Mensal): copiada como está, só reposicionada."""
        orig = self.mod.matrizes[titulo_molde]
        d = self._add(orig, "mat_" + chave)
        d["$schema"] = orig["$schema"]
        if chave == "dre_cc" and NOME_CC != "Cozinha":      # título e nome da coluna do molde falam em "Cozinha" (cliente A)
            txt = json.dumps(d, ensure_ascii=False).replace("Demonstrativo de Resultado por Cozinha", "Demonstrativo de Resultado por " + NOME_CC)
            d = json.loads(txt.replace('"displayName": "Cozinha"', '"displayName": "%s"' % NOME_CC))
        posicionar(d, x, y, w, h, self._z(), self._z())
        return d

    def tabela_lancamentos(self, x, y, w, h):
        cols = [("Realiz.", "Data"), ("Descrição", "Descrição"), ("Detalhamento/Categoria II", "Categoria II"), ("Grupo", "Grupo"),
                ("Categoria", "Categoria"), (cfg.COL_CC, NOME_CC), (cfg.COL_CLIENTE, "Cliente / Fornecedor"), ("Conta", "Conta"),
                ("Valor", "Valor")]
        larg = (78, 220, 180, 110, 130, 60, 210, 112, 110)
        d = {"$schema": SCHEMA_VISUAL, "name": novo_id(self.nome, "tabela"),
             "visual": {"visualType": "tableEx",
                        "query": {"queryState": {"Values": {"projections": [proj(col(CF, c), r) for c, r in cols]}},
                                  "sortDefinition": {"sort": [{"field": col(CF, "Realiz."), "direction": "Descending"}]}},
                        "objects": {"values": [{"properties": {"fontSize": lit("10D")}}],
                                    "columnWidth": [{"properties": {"value": lit("%dD" % lg)}, "selector": {"metadata": "%s.%s" % (CF, c)}}
                                                    for (c, _), lg in zip(cols, larg)]},
                        "visualContainerObjects": container_padrao(),
                        "drillFilterOtherVisuals": True}}
        posicionar(d, x, y, w, h, 4000, 4000)
        self.visuais[d["name"]] = d
        return d

    def json_pagina(self):
        return {"$schema": SCHEMA_PAGE, "name": self.nome, "displayName": self.exib, "displayOption": "FitToPage", "height": 720,
                "width": 1280,
                "objects": {"background": [{"properties": {"color": {"solid": {"color": lit("'#F2F5F9'")}}, "transparency": lit("0D")}}],
                            "outspace": [{"properties": {"color": {"solid": {"color": lit("'#F2F5F9'")}}, "transparency": lit("0D")}}]}}


# ---------------------------------------------------------------- páginas
def montar(mod):
    paginas = []
    H = 704 - Y0                                                  # altura útil abaixo da faixa azul
    L, LM = 1248, 618                                             # largura útil; largura de meia coluna (2 × 618 + 12)

    # 1. Receita: cartões + Caixa × Competência (os gráficos 3, 8 e 10 diziam o mesmo; ficou só o 10, a pedido do responsável em 09/10)
    p = Pagina("Receita", mod, "Receita")
    p.filtros_padrao()
    HC = 112
    p.cartao("receita", "Receita Operacional", "Receita_Operacional_Mes", 16, Y0, 300, HC, subtitulo="Sub_Mes")
    p.cartao("resultado", "Resultado do Exercício", "Resultado_Mes", 332, Y0, 300, HC, medida_cor="Cor_Resultado_Mes", subtitulo="Sub_Mes")
    p.cartao("acumulado", "Resultado acumulado · ano", "Resultado_Acum_Ano_Mes", 648, Y0, 300, HC, medida_cor="Cor_Resultado_Acum",
             subtitulo="Sub_Acum_Ano")
    if cfg.TEM_COMPETENCIA:
        p.cartao("markup", "Markup · competência", "Markup_Competencia_Ult", 964, Y0, 300, HC, subtitulo="Sub_Competencia")
    y2 = Y0 + HC + 12
    if cfg.TEM_COMPETENCIA:
        d = p.colunas_duas("caixa_comp", "Receita Caixa × Receita Competência", "Painel_Receita_Eixo", "Caixa", "Painel_Receita_Comp_Eixo", "Competência",
                           16, y2, L, 704 - y2, rotulos=True)
    else:
        d = p.colunas("caixa", "Receita Operacional", "Painel_Receita_Eixo", 16, y2, L, 704 - y2)
    p.subtitulo(d, "Em R$ milhões · 12 meses até o mês de referência (sem Período selecionado) ou os meses selecionados")
    paginas.append(p)

    # 2. Clientes e C.C.: Pareto de clientes (no lugar da pizza) e pizza por C.C. (só se o cliente usa C.C.)
    p = Pagina(ABA_CLIENTES, mod, ABA_CLIENTES)
    p.filtros_padrao()
    d = p.pareto("pareto", "Pareto de Clientes", 16, Y0, 800 if cfg.TEM_CC else L, H)
    p.subtitulo_med(d, "Sub_Mes")
    if cfg.TEM_CC:
        d = p.pizza("cozinha", "Divisão da Receita por " + NOME_CC, (CF, cfg.COL_CC), (CF, "Receita_Operacional_Mes"), 828, Y0, 436, H)
        p.subtitulo_med(d, "Sub_Mes")
    paginas.append(p)

    # 3. Resultado e gastos: pizza do gasto + resultado do mês + acumulado do ano
    p = Pagina("Resultado e gastos", mod, "Resultado e gastos")
    p.filtros_padrao()
    d = p.pizza("gasto", "Divisão do Gasto pela Receita", ("dCategoriasGrafico", "Categoria_Graf"), ("dCategoriasGrafico", "Plot_GastoVsReceita_Mes"),
                16, Y0, 560, H, cores=AZUIS_PIZZA, rotulo="Data", ordenar_por_valor=False)
    p.subtitulo_med(d, "Sub_Mes")
    h1 = (H - 12) // 2
    d = p.colunas("resultado", "Resultado do Exercício", "Painel_Resultado_Eixo", 588, Y0, 676, h1, unidade="1000D", precisao="0L")
    p.subtitulo(d, "Resultado do mês, em R$ mil · acima do zero sobrou, abaixo do zero faltou")
    d = p.linha("acumulado", "Resultado acumulado do ano", "Painel_Resultado_Acum_Eixo", 588, Y0 + h1 + 12, 676, h1, unidade="1000D", precisao="0L",
                cor=AZUL_ESCURO)
    p.subtitulo(d, "Soma dos resultados desde janeiro, em R$ mil")
    paginas.append(p)

    # 4. Eficiência e pessoal: markup + custo/competência (TEM_COMPETENCIA) e despesas com pessoal (TEM_GRUPOS_PESSOAL).
    # Só uma das duas áreas: ela ocupa a página inteira; as duas: competência em cima e pessoal embaixo.
    if cfg.TEM_COMPETENCIA or cfg.TEM_GRUPOS_PESSOAL:
        p = Pagina(ABA_EFICIENCIA, mod, ABA_EFICIENCIA)
        p.filtros_padrao()
        h1 = (H - 12) // 2 if (cfg.TEM_COMPETENCIA and cfg.TEM_GRUPOS_PESSOAL) else H
        if cfg.TEM_COMPETENCIA:
            d = p.linha("markup", "Markup", "Eixo_Ano_Markup", 16, Y0, LM, h1, unidade="1D", precisao="2L", cor=AZUL)
            p.subtitulo(d, "Receita por competência ÷ custo operacional · até o último mês com competência")
            d = p.linha("custo", "Custo Operacional (% da Receita Competência)", "Eixo_Ano_Custo_Competencia", 16 + LM + 12, Y0, LM, h1, unidade="1D",
                        precisao="1L", cor=AZUL)
            p.subtitulo(d, "% da receita por competência · até o último mês com competência")
        if cfg.TEM_GRUPOS_PESSOAL:
            y_p = Y0 + h1 + 12 if cfg.TEM_COMPETENCIA else Y0
            d = p.linhas("pessoal", "Despesas com Pessoal detalhadas",
                         [("Eixo12M_Pessoal_Salarios", "Salários", "#094E8C"), ("Eixo12M_Pessoal_Encargos", "Encargos", "#5E9BCF"),
                          ("Eixo12M_Pessoal_Outros", "Outros", "#8A97A8")], 16, y_p, L, h1)
            p.subtitulo(d, "Em R$ mil · 12 meses até o mês de referência (sem Período selecionado) ou os meses selecionados · Encargos inclui Provisões")
        paginas.append(p)

    # 5. DRE
    p = Pagina("DRE", mod, "DRE Gerencial")
    p.filtros_padrao()
    p.matriz("Demonstrativo de Resultado", "dre", 16, Y0, L, H)
    paginas.append(p)

    # 6. DRE por C.C. (só se o cliente usa C.C.); o molde da matriz continua com o título do cliente A
    if cfg.TEM_CC:
        p = Pagina(ABA_DRE_CC, mod, "DRE por " + NOME_CC)
        p.filtros_padrao()
        p.matriz("Demonstrativo de Resultado por Cozinha", "dre_cc", 16, Y0, L, H)
        paginas.append(p)

    # 7. DFC (com a DRE Alternativa)
    p = Pagina("DFC", mod, "DFC e DRE Alternativa")
    p.filtros_padrao()
    ha = 229
    p.matriz("DRE Alternativa", "dre_alt", 16, Y0, L, ha)
    p.matriz("Fluxo de Caixa Mensal", "dfc", 16, Y0 + ha + 12, L, H - ha - 12)
    paginas.append(p)

    # 8. Lançamentos: de onde vêm os números (todos os lançamentos, com filtros próprios)
    p = Pagina("Lançamentos", mod, "Lançamentos")
    p.filtros_padrao()
    p.slicer("grupo", CF, "Grupo", "Grupo", 16, "Grupo", largura=300, y=Y0, h=58, corpo=True)
    p.slicer("categoria", CF, "Categoria", "Categoria", 332, "Categoria", largura=300, y=Y0, h=58, corpo=True)
    p.slicer("categoria2", CF, "Detalhamento/Categoria II", "Detalhamento / Categoria II", 648, "Categoria II", largura=300, y=Y0, h=58, corpo=True)
    p.slicer("conta", CF, "Conta", "Conta", 964, "Conta", largura=300, y=Y0, h=58, corpo=True)
    p.tabela_lancamentos(16, Y0 + 70, L, H - 70)
    paginas.append(p)
    return paginas


def validar(paginas, catalogo):
    erros = []
    usados = set()
    for p in paginas:
        for d in p.visuais.values():
            campos_usados(d, usados)
        json.dumps(p.json_pagina())
    for tab, nome in sorted(u for u in usados if u not in catalogo):
        erros.append("campo inexistente no modelo: %s[%s]" % (tab, nome))
    nomes = [d["name"] for p in paginas for d in p.visuais.values()]
    if len(nomes) != len(set(nomes)):
        erros.append("nomes de visual repetidos")
    for p in paginas:
        for d in p.visuais.values():
            pos = d["position"]
            if pos["x"] < 0 or pos["y"] < 0 or pos["x"] + pos["width"] > 1281 or pos["y"] + pos["height"] > 721:
                erros.append("visual fora da página %s: %s" % (p.exib, d["visual"]["visualType"]))
    return erros, usados


def gravar(paginas, destino):
    origem = zipfile.ZipFile(PBIX)
    tmp = destino + ".tmp"
    with zipfile.ZipFile(tmp, "w") as out:
        for info in origem.infolist():
            n = info.filename
            if n.startswith("Report/definition/pages/") or n == "SecurityBindings":
                continue
            novo = zipfile.ZipInfo(n, date_time=info.date_time)
            novo.compress_type = info.compress_type
            novo.external_attr = info.external_attr
            out.writestr(novo, origem.read(n))

        def add(nome, obj):
            zi = zipfile.ZipInfo(nome, date_time=datetime.datetime.now().timetuple()[:6])
            zi.compress_type = zipfile.ZIP_DEFLATED
            out.writestr(zi, json.dumps(obj, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))
        add("Report/definition/pages/pages.json", {"$schema": SCHEMA_PAGES, "pageOrder": [p.nome for p in paginas],
                                                   "activePageName": paginas[0].nome, "landingPageName": paginas[0].nome})
        for p in paginas:
            add("Report/definition/pages/%s/page.json" % p.nome, p.json_pagina())
            for d in p.visuais.values():
                add("Report/definition/pages/%s/visuals/%s/visual.json" % (p.nome, d["name"]), d)
    # confere que o DataModel não mudou um byte e que o zip abre
    novo = zipfile.ZipFile(tmp)
    assert novo.testzip() is None
    assert novo.read("DataModel") == origem.read("DataModel"), "DataModel diferente!"
    assert novo.getinfo("DataModel").compress_type == origem.getinfo("DataModel").compress_type
    novo.close()
    shutil.move(tmp, destino)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--gravar", action="store_true")
    a = ap.parse_args()
    mod = Modelos(zipfile.ZipFile(MODELOS_PBIX))
    paginas = montar(mod)
    catalogo, pasta = catalogo_modelo(TMDL)
    erros, usados = validar(paginas, catalogo)
    print("catálogo do modelo:", pasta, "|", len(catalogo), "objetos | campos usados nas páginas:", len(usados))
    for p in paginas:
        print("  %-22s %2d visuais: %s" % (p.exib, len(p.visuais), sorted({d["visual"]["visualType"] for d in p.visuais.values()})))
    if erros:
        print("ERROS:")
        for e in erros:
            print("  -", e)
        sys.exit(1)
    print("validação: OK")
    if a.gravar:
        destino = os.path.join(os.environ["TEMP"], "AnaliseFinanceira_novo_%s.pbix" % cfg.CODIGO)
        gravar(paginas, destino)
        print("gravado em", destino, "(DataModel idêntico, byte a byte)")


if __name__ == "__main__":
    main()
