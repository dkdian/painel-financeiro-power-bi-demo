# Contrato da planilha de dados

O modelo Power BI lê **todos os `Controle Financeiro AAAA.xlsx/.xlsm`** da raiz de `00_Painel BI` (pasta local, sem SharePoint) e empilha.
Um arquivo por ano, nome exato `Controle Financeiro 2026.xlsx` (ou `.xlsm`). **Nunca dois arquivos do mesmo ano na pasta** (o `.xlsx` e o
`.xlsm`, ou uma cópia de conflito do OneDrive `...-COPIA.xlsm`): a BI soma as linhas duas vezes. Ano novo: basta o arquivo do ano
existir na pasta (entra sozinho). Conferência: `python ferramentas_kit\verificar_planilha.py "<cópia do arquivo>"`.

> **Padrão desde 19/10/2026 (decisão do responsável): a arquitetura agregada** vale para o cliente A, a cliente B e os clientes novos. A direta fica descrita como legado.

## As duas arquiteturas (config `ARQUITETURA_DADOS`)

| | **direta** (cliente A) | **agregada** (cliente B) |
|---|---|---|
| Onde a equipe lança | No próprio `Controle Financeiro AAAA.xlsm` de `00_Painel BI` (tabela única) | Na planilha de trabalho `NNN_Controle Financeiro AAAA.xlsm` (abas mensais), um nível acima |
| Quem grava o que a BI lê | A equipe | O `adaptador_dados.py` do cliente, a cada clique: `Controle Financeiro AAAA.xlsx` (com coluna `Origem`, aba `Conferencia`) |
| Botão (macro `Botao.bas`) | Na planilha do painel: `EXE_RELATIVO = "02_Automacao\..."` | Na planilha de trabalho: `EXE_RELATIVO = "00_Painel BI\02_Automacao\..."` |
| Planilha aberta no Excel | O `.exe` salva e, se o Excel travou o arquivo, a BI lê uma cópia (`BI_PASTA_DADOS`) | O `.exe` salva a de trabalho e fecha a agregada (ela é regravada) |
| Lançamento depois do fechamento | Direto na planilha | `lancamentos_agregada.csv` (Origem "Adicionado pela consultoria") |
| Diferença de saldo entre meses | Linha `Ajuste de Saldo` digitada (lista aprovada pelo responsável) | Ajuste automático do adaptador (conta que some da aba; ligação dez→jan), sempre no ALERTAS |
| Quando escolher | A equipe aceita lançar numa tabela única no padrão | A equipe já tem uma planilha própria e não vai mudar o jeito de lançar |

## Abas

| Aba | Conteúdo |
|---|---|
| `Controle Financeiro` | Topo com saldos de início por conta e saldos finais calculados; cabeçalho na linha 8; **tabela do Excel chamada `ControleFinanceiro`** |
| `Lista` | Listas das validações (Conta, Categoria, Categoria II por categoria via `INDIRECT`, C. C., Fornecedor/Cliente) |
| `Competencia` | Tabela `tbCompetencia` (`Mês`, `Valor`): receita por competência, uma linha por mês (dia 1). Só se o cliente usa |
| `Saldos` | Coluna B = conta, C = saldo no **dia zero** (primeiro dia do primeiro ano); linha "Resultado acumulado anterior" (do último relatório antes do painel; 0 se não houver). Lida pela consulta `dParametros` **só no arquivo mais antigo** da pasta |
| `Relatorio` | Instruções e o botão "Gerar relatório" (forma com a macro `Botao.GerarRelatorio`; por isso a planilha onde a equipe lança vira `.xlsm`). O Office pede "Habilitar conteúdo" na primeira vez e duas confirmações no clique (Yes e OK) |

## Colunas da tabela `ControleFinanceiro` (13, nesta ordem e com estes nomes)

| Coluna | Regra |
|---|---|
| `Realiz.` | Data do pagamento/recebimento realizado. Obrigatória. Define o mês do lançamento |
| `Venc.` | Vencimento (a BI descarta) |
| `C. C.` | Centro de custo. Sem C.C. no cliente: um valor único (`GER`) |
| `  Fornecedor/ Cliente` | **Com dois espaços no início**: é o nome histórico; não corrigir |
| `Descrição` | Livre |
| `Detalhamento/Categoria II` | Subcategoria; depende da `Categoria` (lista dependente). `-` quando não há |
| `Tipo` | `Receita`, `Gasto`, `Movimentação` ou `Sócio` (a BI chama de `Grupo`) |
| `Categoria` | Uma das 14 abaixo. Sem espaço nas pontas |
| `Desembolso` | Forma (PIX, Boleto, C/D em Conta, Dinheiro) (a BI descarta) |
| `Conta` | Banco/caixa onde o dinheiro entrou ou saiu. **Obrigatória**: sem ela os saldos por conta não fecham |
| `Valor` | **Sempre positivo**; o sinal vem da Categoria. Até 2 casas decimais (mais que isso desalinha centavos da DFC: o relatório alerta) |
| `Saldo` | Fórmula (saldo corrente). A BI importa, mas o saldo da DFC é recalculado por medida |
| `Comentários` | Livre (a BI descarta) |

## As 14 categorias

| Categoria | Entra como |
|---|---|
| `Operacional` | Receita operacional |
| `Custo.Operacional` | Custo operacional (gasto) |
| `Tributo` | Tributo (gasto) |
| `Despesa.Fixa`, `Despesa.Variável`, `Despesa.com.Pessoal`, `Marketing.Comercial` | Gastos operacionais |
| `Receita.Financeira`, `Despesa.Financeira` | Resultado financeiro |
| `Investimento` | Aquisição de ativo; com Categoria II `Venda de Ativo Imobilizado` é **venda** (recebimento) |
| `Aporte`, `Retirada` | Capital do sócio |
| `Entrada`, `Saída` | Movimentação entre contas, empréstimos, aplicações (não é resultado) |

Categoria fora desta lista não é ignorada em silêncio: aparece em `Chk_Sem_Categoria` e vira alerta do relatório. Cliente com categorias
próprias: **mapear para estas** na planilha (preferível) ou levar a regra ao responsável (`03_`, #10); mudar o modelo é último recurso.

## Saldos

- **Dia zero, uma vez só** (`dParametros[SaldoDiaZero]`): saldo total das contas no primeiro dia do primeiro ano. Todo o resto vem dos lançamentos.
- O saldo final de um mês tem de ser igual ao inicial do seguinte, **por conta**. Quando a planilha digita o saldo real do banco no início do
  mês, a diferença aparece "sem lançamento". Antes de ajustar, ache a causa na origem (lançamento sem `Conta`, transferência com uma perna só).
- Sobrando diferença legítima: lançamento `Categoria = Entrada`, `Categoria II = Ajuste de Saldo`, descrição com o motivo, datado no **último dia do mês
  anterior**. **A lista (data, valor, motivo) vai ao responsável para aprovação antes de lançar.** Resíduo de até R$ 0,01 vem de casas decimais nos lançamentos.
- Conferência de fechamento por conta contra o extrato, mês a mês (no cliente A: 54 de 55 pares conta-mês fecharam em R$ 0,00 na primeira rodada).

## De onde vem a planilha

| Situação | O que fazer |
|---|---|
| Cliente já usa uma planilha no formato da consultoria (`NNN_Controle Financeiro AAAA.xlsm`, abas por mês) e vai continuar lançando nela | Arquitetura **agregada**: `adaptador_dados.py` (modelo: `referencia\cliente_b\importar_abas.py` + `atualizar_agregada.py`). Escrever o `.xlsx` agregado **pelo Excel (COM)**, nunca por `openpyxl` em arquivo existente (apaga gráficos e desarruma formatação) |
| Cliente aceita lançar numa tabela única | Arquitetura **direta**: importação única das abas antigas (modelo: `importar_2026.py` do cliente A) e daí em diante só conferir |
| Origem em outro formato (sistema, extrato, DRE pronta) | Fora do padrão: levar ao responsável antes de qualquer código (`02_`, "sinais de que foge do padrão") |

Correções de escrita na importação que o responsável já aprovou como regra: `Aporte Sócio` → `Aporte`, `Despesa.fixa` → `Despesa.Fixa`, `adm` → `ADM`. Outras: perguntar.

## Virada de ano

- O arquivo do ano novo entra sozinho na BI (consulta de pasta). O dia zero vale uma vez para a série (arquivo mais antigo): o saldo inicial de
  janeiro é o final de dezembro, porque a DFC soma o fluxo de todos os anos antes do período.
- Agregada (cliente B, testado com 2027 simulado em 16/10/2026): a planilha de trabalho do ano novo nasce por cópia e **tem de se chamar exatamente**
  `NNN_Controle Financeiro AAAA.xlsm`; o adaptador só lê linhas com data do ano do arquivo, liga o saldo de dezembro ao de janeiro (ajuste
  datado de 31/12 se diferirem) e a conferência pula as médias 3M em jan/fev (a janela cruza o ano).
- Direta: criar `Controle Financeiro AAAA.xlsm` do ano novo (cópia sem lançamentos, com a aba `Relatorio` e a macro) e lançar nele.
- Nos dois: grupo de cálculo `Análise DRE` (itens `Jan/AA..Dez/AA`) e o rótulo "Acumulado no ano" do painel, que soma a série inteira.

## Cuidados com o arquivo

- Mestre aberto pelo cliente com **AutoSave**: nunca editar por COM/openpyxl enquanto alguém o tem aberto (um salvamento foi sobrescrito por outra cópia aberta).
- Backup do `.xlsx` antes de qualquer ajuste, em `90_Backups`, **nunca** dentro de pasta que o Power Query varre.
- Planilha no OneDrive: depois de escrever, reler do disco para conferir.
