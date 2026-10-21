// Consulta Receita_Competencia_Manual (padrão do kit, 19/10/2026; vem do cliente A, 15/10/2026). Só se o cliente usa competência (TEM_COMPETENCIA).
// Tabela tbCompetencia (aba Competencia: Mês = dia 1 do mês, Valor) de todos os arquivos da pasta. Arquivo sem a tabela é ignorado.
let
    // Pasta do painel neste computador. O gerador de relatório troca o valor de PastaLocal a cada execução (bi_conexao.definir_pasta_local); mantenha esta linha como está.
    PastaLocal = "C:\caminho\do\cliente\Documentos Elaborados\00_Painel BI",
    Arquivos = Table.Sort(Table.SelectRows(Folder.Contents(PastaLocal), each List.Contains({".xlsx", ".xlsm"}, Text.Lower([Extension])) and Text.StartsWith([Name], "Controle Financeiro 20") and not Text.StartsWith([Name], "~$")), {{"Name", Order.Ascending}}),
    Fontes = List.Transform(Arquivos[Content], each Excel.Workbook(_, null, true)),
    ComTabela = List.Select(Fontes, each Table.RowCount(Table.SelectRows(_, each [Item] = "tbCompetencia" and [Kind] = "Table")) > 0),
    Competencia_Table = Table.Combine(List.Transform(ComTabela, each _{[Item="tbCompetencia",Kind="Table"]}[Data])),
    #"Colunas Renomeadas" = Table.RenameColumns(Competencia_Table,{{"Mês", "Data_Competencia"}, {"Valor", "Valor_Competencia"}}),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Colunas Renomeadas",{{"Data_Competencia", type date}, {"Valor_Competencia", type number}})
in
    #"Tipo Alterado"
