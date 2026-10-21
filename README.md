# Painel financeiro em Power BI: demonstração

Versão de demonstração de um projeto que uso em um cliente real. O projeto, com o problema, o resultado e um vídeo de 3 minutos, está no portfólio:
https://dkdian.github.io/projeto.html?p=painel-financeiro-power-bi

> **Todos os dados, nomes, valores e telas deste repositório são fictícios.** A empresa ("Cozinhas Modelo"), as cozinhas, os clientes,
> os fornecedores e as contas bancárias foram inventados só para esta demonstração. O que é real é o método.

## O que tem aqui

| Pasta | O que é |
|---|---|
| `dados_demo/` | Gerador da planilha sintética no contrato de dados do painel (`gerar_dados_demo.py`: 15 meses, 5 contas, 3 cozinhas) e da planilha antiga e da conferência com extratos usadas no vídeo |
| `modelo_bi/M_padrao/` | As três consultas Power Query do modelo: lê **todos** os arquivos `Controle Financeiro AAAA` da pasta, empilha, e lê o saldo de abertura da aba `Saldos` |
| `modelo_bi/tmdl/` | Export TMDL do modelo (tabelas, relações e **todas as medidas DAX**, incluindo o grupo de cálculo da DRE) |
| `paginas/` | `paginas_painel.py` e `moldes_paginas.zip`: escrevem as 8 páginas do painel (formato PBIR) com o Power BI Desktop fechado |
| `docs/` | Contrato da planilha e padrão das páginas |
| `video/` | O montador do vídeo de demonstração (Pillow + ffmpeg) e o roteiro |

O arquivo `.pbix` **não** está aqui: o painel se reconstrói a partir do TMDL, das consultas e das páginas.

## Como funciona

1. A equipe lança numa tabela por ano (`Controle Financeiro AAAA.xlsx`, 13 colunas fixas, 14 categorias padronizadas, conta em cada lançamento).
2. A consulta `ControleFinanceiro` usa `Folder.Contents` e empilha os anos: o arquivo do ano novo entra sozinho.
3. As medidas DAX calculam receita, resultado, DRE (mensal, por cozinha), DFC e indicadores; o rateio de tributos e administrativo segue a regra da planilha.
4. As 8 abas mostram receita, clientes e cozinhas, resultado e gastos, eficiência e pessoal, DRE, DRE por cozinha, DFC e lançamentos.

## Como rodar a parte de dados

```bash
pip install -r requirements.txt
python dados_demo/gerar_dados_demo.py "C:\\caminho\\da\\pasta\\00_Painel BI"
```

Cria `Controle Financeiro 2025.xlsx` e `Controle Financeiro 2026.xlsx` com lançamentos inventados. Em seguida, no Power BI Desktop, aponte a linha
`PastaLocal` das consultas para essa pasta.

## Limites conhecidos

A matriz da DRE expandida é mais lenta na tela do que a consulta equivalente. O painel precisa do Power BI Desktop no Windows.
