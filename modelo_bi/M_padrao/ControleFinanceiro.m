// Consulta ControleFinanceiro (padrão do kit, 19/10/2026). Junta o cliente A (15-16/10: pasta local, .xlsx e .xlsm) e a cliente B (16/10: coluna Origem).
// Lê TODOS os "Controle Financeiro AAAA.xlsx/.xlsm" da pasta do painel: o arquivo do ano novo entra sozinho, sem login nem nível de privacidade.
// ATENÇÃO: um arquivo por ano. Se houver o .xlsx e o .xlsm do mesmo ano (ou uma cópia de conflito do OneDrive), as linhas entram duas vezes.
let
    // Pasta do painel neste computador. O gerador de relatório troca o valor de PastaLocal a cada execução (bi_conexao.definir_pasta_local); mantenha esta linha como está.
    PastaLocal = "C:\caminho\do\cliente\Documentos Elaborados\00_Painel BI",
    Arquivos = Table.Sort(Table.SelectRows(Folder.Contents(PastaLocal), each List.Contains({".xlsx", ".xlsm"}, Text.Lower([Extension])) and Text.StartsWith([Name], "Controle Financeiro 20") and not Text.StartsWith([Name], "~$")), {{"Name", Order.Ascending}}),
    Fontes = List.Transform(Arquivos[Content], each Excel.Workbook(_, null, true)),
    Tabela = Table.Combine(List.Transform(Fontes, each _{[Item="ControleFinanceiro",Kind="Table"]}[Data])),
    // Origem: só existe na planilha agregada (arquitetura "agregada"); na direta vira coluna vazia, para o modelo ser o mesmo.
    ComOrigem = if Table.HasColumns(Tabela, "Origem") then Tabela else Table.AddColumn(Tabela, "Origem", each null, type text),
    #"Tipo Alterado" = Table.TransformColumnTypes(ComOrigem,{{"Realiz.", type date}, {"Venc.", type date}, {"C. C.", type text}, {"  Fornecedor/ Cliente", type text}, {"Descrição", type text}, {"Detalhamento/Categoria II", type text}, {"Tipo", type text}, {"Categoria", type text}, {"Desembolso", type text}, {"Conta", type text}, {"Valor", type number}, {"Saldo", type number}, {"Comentários", type any}, {"Origem", type text}}),
    #"Mês Inserido" = Table.AddColumn(#"Tipo Alterado", "Mês_Num", each Date.Month([#"Realiz."]), Int64.Type),
    #"Nome do Mês Inserido" = Table.AddColumn(#"Mês Inserido", "Mês_Nome", each Date.MonthName([#"Realiz."]), type text),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Nome do Mês Inserido",{"Realiz.", "Mês_Num", "Mês_Nome", "Venc.", "C. C.", "  Fornecedor/ Cliente", "Descrição", "Detalhamento/Categoria II", "Tipo", "Categoria", "Desembolso", "Conta", "Valor", "Saldo", "Comentários", "Origem"}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Colunas Reordenadas",{{"Tipo", "Grupo"}}),
    #"Colunas Removidas" = Table.RemoveColumns(#"Colunas Renomeadas",{"Desembolso", "Comentários", "Venc."}, MissingField.Ignore)
in
    #"Colunas Removidas"
