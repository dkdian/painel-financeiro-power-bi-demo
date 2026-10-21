# Painel Power BI padrão

Modelo: o painel da Distribuidora Modelo e, depois, o do cliente A (PLANO G, 09/10/2026). Faixa azul `#094E8C` de 150 px em **todas** as páginas: título, subtítulo,
carimbo "Dados até" (`Txt_Dados_Ate`), filtros sincronizados (Período, C.C. se houver, Cliente/Fornecedor), pastilha **✕ Limpar filtros** e a linha de abas
(a atual em branco, as outras translúcidas). Cada aba são 3 camadas (forma + caixa de texto + forma transparente com a ação), porque o botão de ação não
desenha preenchimento neste Desktop. Gráficos em cartões brancos com borda, sombra, título e subtítulo cinza que explica o que o gráfico mostra.
No Desktop a navegação é **Ctrl + clique**; para o cliente, no Service, é clique simples.

## Páginas (as que existem dependem dos módulos)

| Aba | Conteúdo | Depende de |
|---|---|---|
| Receita | 3 cartões (Receita Operacional, Resultado, Resultado acumulado do ano; verde/vermelho) + Receita Caixa × Competência (só caixa sem competência) + cartão Markup | `TEM_COMPETENCIA` para Markup e competência |
| Clientes (e C.C.) | Pareto de clientes + pizza por C.C. | `TEM_CC` para a pizza |
| Resultado e gastos | Pizza "Divisão do Gasto" + Resultado do mês (colunas) + Resultado acumulado (linha) | — |
| Eficiência e pessoal | Markup, Custo/Receita Competência, Despesas com Pessoal detalhadas | `TEM_COMPETENCIA` e/ou `TEM_GRUPOS_PESSOAL`; some se nenhum |
| DRE | Matriz da DRE (grupo de cálculo) | — |
| DRE por C.C. | Matriz da DRE por C.C. | `TEM_CC` |
| DFC | DRE Alternativa + Fluxo de Caixa Mensal | — |
| Lançamentos | Filtros no corpo (Grupo, Categoria, Categoria II, Conta) + tabela de lançamentos, mais recentes primeiro | — |

Gráficos em ordem: **gráficos antes das tabelas**. Regra de período: gráficos de 12 meses terminam no mês selecionado (sem seleção, no último mês com
dados); gráficos "do mês" e cartões, sem seleção, mostram o último mês com dados; com Período selecionado, só os meses selecionados.

## Como as páginas são escritas

`02_Automacao\ferramentas\paginas_painel.py` escreve o relatório PBIR de dentro do `.pbix` **com o Desktop fechado**: copia todas as entradas byte a byte
(o `DataModel` não muda; o script confere), troca só `Report/definition/pages/**`, **remove `SecurityBindings`** (o Desktop recusa `.pbix` editado por fora com ela)
e valida os campos contra o export TMDL mais recente (`03_Documentacao\gerado\Modelo_TMDL_*`, o último por nome).

```
python paginas_painel.py            # só valida e resume; o arquivo vai para %TEMP%, nunca sobre o original
python paginas_painel.py --gravar   # grava AnaliseFinanceira_novo_NNN.pbix em %TEMP%
```

O `--gravar` **nunca** sobrescreve o `.pbix`: o responsável (ou o Claude, com aviso) compara e substitui, depois de **backup datado**.
Moldes (slicer, forma, caixas, colunas, linha, pizza e as 4 matrizes) vêm de `ferramentas\moldes_paginas.zip` (só `Report/definition/pages/**`, sem dados).
Rodar de novo refaz as páginas e **desfaz ajustes feitos à mão no Desktop**: ajuste de aparência que valha sempre entra no script, não só no Desktop.

## Ordem de trabalho (uma camada por sessão)

1. Modelo pelo MCP com o Desktop **aberto** → salvar → exportar TMDL → **fechar o Desktop**.
2. `paginas_painel.py --gravar` com o Desktop **fechado** → conferir → substituir o `.pbix` (backup antes).
3. Abrir, atualizar, conferir cada página na tela contra o plano; responsável salva.

## Armadilhas do Desktop 2.158

- Abrir segunda instância: `Start-Process <pbix>` (o exe da WindowsApps dá "Acesso negado"); esperar ~110 s antes de automatizar.
- Com **dois Desktops abertos** (dois clientes têm o mesmo nome de arquivo), escolher a instância pela linha de comando (`*006_Hora*` / `*023_Casa*`); nunca encerrar a do outro.
  Duas instâncias do **mesmo** arquivo: Salvar falha com aviso modal ("already open") e a data do arquivo não muda; fechar a outra primeiro.
- Salvar por automação: botão `Save` por `InvokePattern` (Ctrl+S via SendKeys não chega). Banner "Refresh now" para tabela calculada nova (o refresh pelo MCP falha com "base version negative").
- A conexão do MCP **morre depois que o Desktop salva**: reconectar (`ListLocalInstances`).
- Renderização: gráfico combinado não desenha rótulo de coluna e ignora cor por regra; colunas empilhadas escondem rótulo com barra estreita (12 meses); molde com
  `labelPosition: OutsideEnd` é inválido para empilhado; cor por regra/campo em barras não aplicou. **Cor estática de série e cor de cartão vinda de medida de texto funcionam.**
- Medida que ignora filtro de linha mas mantém o Período: `REMOVEFILTERS` nas colunas listadas, não `ALL(ControleFinanceiro)`.
- Páginas de exportação para o Word (páginas R + PDF) foram **abandonadas** no cliente A: rótulos somem, tabela corta, cabeçalho não aceita texto dinâmico. O relatório Word usa os gráficos do Python.
- Nunca "Publicar na web" (público e indexável). O compartilhamento correto é o do workspace, com usuário nomeado e licença.
- Capturar a janela por `PrintWindow`, com 12 a 15 s de espera por página (senão o PNG sai uma página atrás). Janela modal aparece como `Pane` sem nome na UIA: ler por captura de tela.

## Limites conhecidos do gerador de páginas

- As matrizes (DRE, DRE por C.C., DRE Alternativa, Fluxo de Caixa) são **copiadas dos moldes do cliente A** e só reposicionadas. O título e o nome da coluna da
  "DRE por C.C." usam o rótulo do config (`ROTULO_CC`), mas a matriz **herda um filtro de visual do cliente A que oculta o C.C. literal `ADM`**: num cliente cujo
  C.C. administrativo tenha outro nome (ou que não rateie), isso precisa ser revisto na matriz.
- Os nomes de coluna dos itens do grupo `Análise DRE` (`Jan/25..Dez/25`) vêm do modelo; se o ano do cliente for outro, as colunas dos moldes podem
  referenciar meses que não existem: conferir a matriz da DRE na tela.
- Cartões e gráficos dependem das medidas de painel do modelo (`Painel_*`, `Sub_*`, `Cor_*`, `Eixo12M_*`); `verificar_modelo.py` as exige.
- O subtítulo da faixa é `PAINEL_SUBTITULO` do config ("Nome · Painel Financeiro · regime de caixa"); se o cliente usa competência, o texto "regime de caixa" pode precisar mudar.

## O que o `.pbix` não faz sozinho

Publicação, convites e licenças, e os acordos com o cliente: `07_checklist_piloto.md` do cliente A é a referência do que perguntar. Desempenho da matriz da DRE:
o grupo de cálculo é ~6× mais lento que medidas simples; a otimização conhecida foi remover o `FormatStringExpression` dos itens de %. Medir no Analisador de Desempenho antes de reclamar.
