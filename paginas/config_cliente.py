"""Configuração do cliente para o código-base do Padrão BI.

COMO USAR: copie este arquivo para <cliente>\\Documentos Elaborados\\00_Painel BI\\02_Automacao\\config_cliente.py
e preencha cada campo. Todo o código de 02_Automacao lê daqui; nenhum outro .py deve ter nome, cor ou regra do cliente.

Cada campo traz a ORIGEM esperada:
  [PRESUMIR] descobrir na pasta do cliente (02_Descoberta_na_Pasta_do_Cliente.md diz onde)
  [PERGUNTAR] decisão do responsável (03_Perguntas_ao_responsável.md)
  [PADRÃO]   só mude se o cliente fugir do padrão (e registre o porquê na ficha)
Exemplo preenchido: exemplos\\config_cliente_cliente_a.py (Cozinhas Modelo).
"""

# ---------------------------------------------------------------- identificação
CODIGO = "999"                      # [PRESUMIR] número da pasta do cliente ("006_Cozinhas Modelo..." -> "cliente A")
NOME = "Cozinhas Modelo"                 # [PRESUMIR] como o cliente é chamado no texto do relatório entregue
NOME_COMPLETO = "Cozinhas Modelo Alimentação Coletiva" # [PRESUMIR] nome da pasta sem o código, ou o da capa do relatório
DESCRICAO_NEGOCIO = "refeições industriais e coffee break/eventos"   # [PRESUMIR/PERGUNTAR] vai no prompt da IA (ex.: "refeições industriais e coffee break/eventos")
CLIENTE_DO = "da " + NOME                   # [PRESUMIR] artigo + nome ("da Distribuidora Modelo", "do Cozinhas Modelo"): copiar do relatório entregue
CLIENTE_DO_COMPLETO = "da " + NOME_COMPLETO   # [PRESUMIR] idem, usado na frase de apresentação

# ---------------------------------------------------------------- arquivos
PBIX = "AnáliseFinanceira"          # [PADRÃO] nome do .pbix (sem extensão) na raiz de 00_Painel BI
TEMPLATE = "Modelo_Relatorio_Financeiro_%s.docx" % CODIGO   # [PADRÃO] em 02_Automacao, criado na fase F6 a partir do modelo geral
PREFIXO_RELATORIO = "999_Relatório Financeiro_"        # [PRESUMIR] igual ao nome dos relatórios entregues
PASTA_ENTREGUES = "01_Relatorios_Entregues"                     # [PADRÃO] relativo a 00_Painel BI
PASTAS_RELATORIOS_EXTRAS = ()                   # [PRESUMIR] outras pastas com relatórios enviados (contador de meses das recomendações)
PLANILHA_REGEX = r"^Controle Financeiro (\d{4})\.xls[xm]$"      # [PADRÃO] um arquivo por ano na raiz de 00_Painel BI (.xlsx ou .xlsm, nunca os dois)
EXE = "GerarRelatorio_999.exe"                   # [PADRÃO]

# ---------------------------------------------------------------- de onde vêm os dados (11_Diferencas, P1; 05_Padrao_Planilha)
# "direta"   : a equipe lança no próprio Controle Financeiro AAAA.xlsm dentro de 00_Painel BI, e a BI lê esse arquivo (cliente A).
# "agregada" : a equipe lança numa planilha de trabalho (abas mensais) e o adaptador_dados.py do cliente regrava a agregada
#              Controle Financeiro AAAA.xlsx a cada execução; a BI lê a agregada (cliente B). Exige adaptador_dados.py em 02_Automacao.
ARQUITETURA_DADOS = "direta"      # [PADRÃO desde 19/10/2026] a equipe lança na planilha de trabalho (abas mensais) e o programa gera a agregada. "direta" só se o cliente lançar numa tabela única
PLANILHA_TRABALHO = r"..\{codigo}_Controle Financeiro {ano}.xlsm"   # [PRESUMIR] só "agregada": caminho relativo a 00_Painel BI

# ---------------------------------------------------------------- conferências
CONFERIR_ENVIO = True        # [PADRÃO] no fim da geração, confere o Word contra a BI (CONFERENCIA_ENVIO_MMMAA.txt)
CARIMBO_BASE = True          # [PADRÃO] grava saida\BASE_MMMAA.json e avisa no ALERTAS quando um mês já entregue muda na base
DFC_INTRO_CAMPOS = "Para melhor entendimento, segue descrição dos campos da DFC"   # [PRESUMIR] frase que abre a lista de campos da DFC no entregue

# ---------------------------------------------------------------- colunas da tabela ControleFinanceiro (contrato de dados)
COL_DATA = "Realiz."
COL_CLIENTE = "  Fornecedor/ Cliente"   # [PADRÃO] sim, com os dois espaços: é o nome histórico da coluna
COL_CC = "C. C."
COL_CATEGORIA = "Categoria"
COL_CATEGORIA_II = "Detalhamento/Categoria II"

# ---------------------------------------------------------------- módulos (liga/desliga partes inteiras)
TEM_CC = True                # [PRESUMIR] a planilha usa C. C. (centro de custo) com mais de um valor real?
ROTULO_CC = "Cozinha"           # [PRESUMIR] como o relatório/cliente chama o C.C. ("Cozinha", "Loja", "Obra", "C.C.")
ROTULO_CC_PLURAL = "cozinhas"    # [PRESUMIR] ("cozinhas", "lojas"...)
TEM_COMPETENCIA = True      # [PRESUMIR] o relatório entregue fala de receita por competência / markup? (aba Competencia na planilha)
COMPETENCIA_O_QUE_E = "das vendas de refeições"   # [PRESUMIR] complemento da frase "regime de competência, ou seja, ..." (só se TEM_COMPETENCIA)
TEM_GRUPOS_PESSOAL = True   # [PERGUNTAR] Pessoal dividido em Salários/Encargos/Outros (tabela dGrupoPessoal + Chk_Pessoal)?
TEM_BENCHMARK = True         # [PRESUMIR] o relatório entregue tem "Indicadores - Benchmark"?
TEM_MARGEM_CONTRIBUICAO = False   # [PERGUNTAR] o relatório entregue tem "Margem de Contribuição"? exige as medidas Eixo12M_Margem_Contribuicao e Margem_Contribuicao no modelo
TXT_BANCOS = "nos bancos Alfa, Beta e Gama e na Cooperativa Delta"   # [PRESUMIR] frase da apresentação com os bancos dos extratos (copiar do relatório entregue)
TEM_DRE_ANO_DEZEMBRO = True  # [PADRÃO] em dezembro entra o bloco "Consolidado do Ano" (G08 + T04)
TEM_RECOMENDACOES_IA = True  # [PADRÃO] rascunho das recomendações pelo Claude Code (só números da BI)
RASCUNHO_IA_NO_WORD = True   # [PADRÃO, decisão do responsável 19/10/2026] rascunho da IA sempre em amarelo no Word (abaixo das recomendações e em Riscos e Oportunidades) e em RASCUNHO_IA_MMMAA.txt; False tira só o Word
TEM_RISCOS_OPORTUNIDADES = True   # seção "Análise de Riscos e Oportunidades" (padrão para todos os clientes desde 19/10/2026)
RECOMENDACOES_IA_MAX = 3     # [PADRÃO] itens novos por seção pedidos ao Claude (cliente A usa 3, cliente B usa 2)
CONTEXTO_CLIENTE = "contexto_cliente.md"   # [PADRÃO] em 02_Automacao: fatos do cliente que o Claude pode usar (a equipe edita)
RECOMENDACOES_SECOES = {"receita": "Recomendações sobre a Receita", "gasto": "Recomendações sobre o Gasto"}   # [PRESUMIR] títulos no entregue
RECOMENDACOES_NUMEROS_FIXOS = ()   # [PRESUMIR] trechos de recomendação com número que NÃO muda (meta fixa), para não gerar alerta. Ex.: ("70%",)

# ---------------------------------------------------------------- DRE e DFC
DRE_MEDIAS = (3, 6, 12)      # [PERGUNTAR] colunas de média da DRE Gerencial (T01). Com pouco histórico: (3,) ou (3, 6)
# Nomes de linha do relatório entregue quando diferem do modelo (o modelo usa nomes curtos na DRE Alt e na DFC)
NOMES_DRE_ALT = {"Receita": "(+) Receita", "Desp. Variável": "(-) Despesa Variável", "CMV": "(-) Custo Operacional CMV",
                 "Desp. Fixa": "(-) Despesa Fixa", "Resultado": "(=) Lucro Operacional"}
NOMES_DFC = {"(-) Pagamentos": "(-) Pagamento Fornecedores", "(+) Aportes": "(+) Novos Empréstimos e Aportes",
             "(-) Retiradas": "(-) Pagamento de Empréstimos e Retiradas"}   # [PRESUMIR] copiar do relatório entregue
# Linha 10 da DFC ("(+/-) Entrada/Saída") aberta em sublinhas: [(ID, nome, medida)]. Vazio = linha única.
DFC_LINHAS_EXTRA = [(10.1, "(+/-) Fundo Financeiro Empresarial", "DFC_Fundo_Financeiro_Empresarial"),
                    (10.2, "(+/-) Aplicação", "DFC_Aplicacao")]        # [PRESUMIR] o entregue abre Entrada/Saída em sublinhas? Ex. cliente A: Fundo Financeiro e Aplicação

# ---------------------------------------------------------------- relatório: gráficos (marcador do template -> desenho)
# chave = tag do template ({{ chave }}); funcao = nome em graficos.py; largura_mm = largura no Word (do entregue);
# so_dezembro = só entra no Consolidado do Ano. Remova a linha do gráfico que o cliente não tem.
GRAFICOS = [
    {"chave": "grafico_receita_12m", "funcao": "g01_receita_12m", "largura_mm": 160},
    {"chave": "grafico_pizza_cliente", "funcao": "g02_pareto_clientes", "largura_mm": 140},
    {"chave": "grafico_pizza_cc", "funcao": "g03_receita_por_cc", "largura_mm": 130},          # exige TEM_CC
    {"chave": "grafico_markup", "funcao": "g04_markup", "largura_mm": 130},                     # exige TEM_COMPETENCIA
    {"chave": "grafico_divisao_gasto", "funcao": "g05_gasto_mes", "largura_mm": 130},
    {"chave": "grafico_custo_receita", "funcao": "g06_custo_competencia", "largura_mm": 130},   # exige TEM_COMPETENCIA
    {"chave": "grafico_resultado", "funcao": "g07_resultado_12m", "largura_mm": 160},
    {"chave": "grafico_gasto_ano", "funcao": "g08_gasto_ano", "largura_mm": 130, "so_dezembro": True},
]
# tabelas: marcador de parágrafo no template -> função em tabelas_word.py (desenhadas como imagem)
TABELAS = {"[[T01]]": "t01_dre_gerencial", "[[T03]]": "t03_dfc", "[[T04]]": "t04_dre_ano", "[[T05]]": "t05_benchmark"}
TABELAS_COMO_IMAGEM = True   # [PADRÃO] decisão do responsável (08/10/2026): tabela como imagem, ninguém edita valor no Word
PARETO_CORTE = 0.80          # [PADRÃO] G02: maiores clientes até 80% da receita do mês; o resto vira "Outros"
VARIACAO_RECEITA_ALERTA = 0.10   # [PADRÃO] alerta quando a receita varia mais que isso contra o mês anterior

# Cores fixas por C.C. no G03 (nome exato do C.C. na planilha). Vazio = paleta por posição (padrão Excel).
CORES_CC = {}                # [PRESUMIR] copiar das pizzas do relatório entregue, se as cores forem fixas por C.C.

# ---------------------------------------------------------------- textos padrão (quando a planilha não traz Categoria II)
TXT_DESPESA_FIN_PADRAO = "Tarifas e Encargos Bancários"
TXT_RECEITA_FIN_PADRAO = "Rendimentos Automáticos na Conta"
TXT_ATIVOS = "automóveis, eletrônicos, equipamentos, móveis e utensílios"   # [PERGUNTAR] exemplos de ativo da empresa (parágrafo "Aquisição de Ativos")
TXT_CUSTO_INTRO_PADRAO = "O Custo Operacional se refere aos insumos da operação."   # [PRESUMIR] copiar do entregue

# ---------------------------------------------------------------- benchmark (só se TEM_BENCHMARK)
# (nome na tabela, linha da DRE/chave, referência, margem_bom, margem_ruim). [PERGUNTAR] as referências do setor do cliente.
# Ruim se valor > ref + margem_ruim; Bom se valor < ref + margem_bom; senão Atenção. Margem de lucro = 1 - soma das refs.
BENCHMARK_REFERENCIAS = [
    ("Tributos", "(-) Tributo", 0.11, 0.01, 0.02),
    ("Custo Operacional", "(-) Custo Operacional", 0.33, 0.02, 0.04),
    ("Despesas com Pessoal", "(-) Despesa com Pessoal", 0.35, 0.02, 0.04),
    ("Despesa Fixa", "(-) Despesa Fixa", 0.07, 0.01, 0.02),
    ("Despesas Variável + Financeira", "Variável + Financeira", 0.02, 0.005, 0.01),
    ("Marketing/Comercial", "(-) Marketing/Comercial", 0.04, 0.005, 0.01),
]
BENCHMARK_NOME_MARGEM = "Resultado/Margem de Lucro"
BENCHMARK_COLUNA_CLIENTE = NOME   # cabeçalho da coluna de valores na tabela do benchmark

# ---------------------------------------------------------------- painel (ferramentas/paginas_painel.py)
PAINEL_SUBTITULO = "%s · Painel Financeiro · regime de caixa" % NOME
PAINEL_MOLDES = "moldes_paginas.zip"   # [PADRÃO] fica em 02_Automacao/ferramentas: zip só com as páginas-molde (vem do kit; sem dados)

# [PRESUMIR/PERGUNTAR] tudo abaixo vem de abrir a planilha de trabalho do cliente (rótulos de saldo, contas, abas, linhas da aba Gráficos). Exemplos: exemplos\config_cliente_cliente_a.py e _cliente_b.py

# ---------------------------------------------------------------- planilha agregada: lida pelo adaptador_dados.py (mesmo código nos dois clientes)
ABAS_MENSAIS = r"^Controle Financeiro (?P<mes>\d{2})(?:-(?P<ano>\d{4}))?$"   # aba "01-2026" (cliente A) ou "01" (cliente B); aba com ano só vale se for o do arquivo
# conta (como está na coluna Conta) -> rótulo do saldo nas abas ("Saldo [em] <rótulo> Início|Final"), forma de desembolso do ajuste e célula do saldo inicial no topo da agregada
CONTAS = {}   # arquitetura direta: não usado
CONTAS_DESCONTINUADAS = ()   # contas que saíram do controle e cuja ausência nas abas novas é esperada
CORRECOES_ESCRITA = {"Categoria": {"Aporte Sócio": "Aporte", "Despesa.fixa": "Despesa.Fixa"}, "C. C.": {"adm": "ADM"}}   # aprovadas pelo responsável em 02/10/2026
CORRECOES_LINHA = []   # [PERGUNTAR] correções de LINHA aprovadas: {"quando": {coluna: valor}, "trocar": {coluna: valor}, "motivo": texto}; só vale se achar exatamente 1 linha
AJUSTE_AUTO_A_PARTIR_DE = None         # [PERGUNTAR] (ano, mês): ajustes automáticos só a partir daí (None = todos os meses); use ao migrar um cliente cujo histórico já foi entregue
CONF_PLANILHA = {"dre": {"col_rotulo": 14, "col_valor": 15, "linha": 9, "n": 13}, "dfc": None}   # [PRESUMIR] onde ficam, na aba mensal, o bloco da DRE (rótulo, valor, 1ª linha) e o da DFC (None se a aba não tem)
TEXTOS_VENDA_ATIVO = ("Venda de Ativo Imobilizado",)
COMPETENCIA = {"aba": "Gráficos", "linha_meses": 1, "linha_valor": 20, "col_ini": 3, "col_fim": 14}   # 0 = ainda não lançado
CONF_GRAFICOS = {"aba": "Gráficos", "layout": "meses_em_colunas", "linha_meses": 1, "linha_receita": 2, "linha_resultado": 14, "col_ini": 3, "col_fim": 14}
ACUM_ANTERIOR = None
LIGAR_ANO_ANTERIOR = True        # liga o saldo final de dezembro ao inicial de janeiro (ajuste de 31/12). False só se o ano anterior estiver congelado
