// Consulta dParametros (padrão do kit, 19/10/2026; vem do cliente B, 16/10/2026). Substitui o SaldoDiaZero digitado dentro do modelo (cliente A).
// Lê a aba "Saldos" do arquivo MAIS ANTIGO da pasta: saldo do dia zero por conta (coluna B = conta, C = valor) e, na linha
// "Resultado acumulado anterior", o resultado acumulado antes do primeiro mês (vem do último relatório antes do painel; 0 se não houver).
// O dia zero vale uma vez só para a série inteira: o saldo inicial de janeiro de um ano novo é o saldo final de dezembro (Valor_DFC soma o
// fluxo de todos os anos antes do período). Cliente novo: preencher a aba Saldos; nada a editar no Power BI.
let
    // Pasta do painel neste computador. O gerador de relatório troca o valor de PastaLocal a cada execução (bi_conexao.definir_pasta_local); mantenha esta linha como está.
    PastaLocal = "C:\caminho\do\cliente\Documentos Elaborados\00_Painel BI",
    Arquivos = Table.Sort(Table.SelectRows(Folder.Contents(PastaLocal), each List.Contains({".xlsx", ".xlsm"}, Text.Lower([Extension])) and Text.StartsWith([Name], "Controle Financeiro 20") and not Text.StartsWith([Name], "~$")), {{"Name", Order.Ascending}}),
    Fonte = Excel.Workbook(Arquivos{0}[Content], null, true),
    Saldos = Fonte{[Item="Saldos",Kind="Sheet"]}[Data],
    Numericas = Table.SelectRows(Saldos, each [Column2] <> null and Value.Is([Column3], type number)),
    Contas = Table.SelectRows(Numericas, each not Text.StartsWith(Text.From([Column2]), "Resultado acumulado")),
    AcumLinhas = Table.SelectRows(Numericas, each Text.StartsWith(Text.From([Column2]), "Resultado acumulado anterior")),
    PorConta = List.Transform(Table.ToRows(Table.SelectColumns(Contas, {"Column2", "Column3"})), each {"SaldoDiaZero|" & Text.From(_{0}), _{1}}),
    Acum = if Table.RowCount(AcumLinhas) > 0 then AcumLinhas{0}[Column3] else 0,
    Resultado = #table(type table [Parametro = text, Valor = number],
        {{"SaldoDiaZero", List.Sum(Contas[Column3])}, {"ResultadoAcumuladoAnterior", Acum}} & PorConta)
in
    Resultado
