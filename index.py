import pandas as pd
from pathlib import Path

pasta_projeto = Path(__file__).resolve().parent

caminho_entrada = pasta_projeto / "data" / "Base Vendas - 2021.xlsx"
caminho_saida = pasta_projeto / "Analise_vendas.xlsx" 

df = pd.read_excel(caminho_entrada)

#Verificar valores 
df["Data da Venda"] = pd.to_datetime(df["Data da Venda"], errors="coerce", dayfirst=True)

colunas = ["ID Cliente","SKU","Qtd Vendida","ID Loja","Data da Venda"]

print("\nValores em falta:")
print(df[colunas].isna().sum())

if df[colunas].isna().any().any():
    raise ValueError("Existem dados em falta. Rever antes da analise.")

duplicados = df.duplicated(keep=False)

print(f"Linhas envolvidas em duplicações: {duplicados.sum()}")
print(f"Repetições além da primeira ocorrência: {df.duplicated().sum()}")

if duplicados.any():
    print("\nLinhas duplicadas para revisão:")
    print(df.loc[duplicados].sort_values("Ordem de Compra"))
else:
    print("\nNão foram encontradas linhas duplicadas.")

#Somar as quantidades por cliente e Prodduto
compras_clientes = (df.groupby(["ID Cliente","SKU"], as_index=False)["Qtd Vendida"].sum())

#Produtos mais comprados
top_produtos = (
    df.groupby("SKU", as_index=False)["Qtd Vendida"]
    .sum().sort_values("Qtd Vendida", ascending=False)
    .reset_index(drop=True))
print("\nTop produtos mais comprados:")
print(top_produtos)

# top 5 Produtos mais comprados por cada cliente
compras_clientes = compras_clientes.sort_values(["ID Cliente","Qtd Vendida","SKU"], ascending=[True, False, True])
top_5_por_cliente = compras_clientes.groupby("ID Cliente").head(5)
print("\nTop 5 de produtos por cliente:")
print(top_5_por_cliente)

#Top lojas com mais vendas
ranking_loja = (df.groupby(["ID Loja"], as_index=False)["Qtd Vendida"]
                .sum()
                .sort_values("Qtd Vendida", ascending=False)
                .reset_index(drop=True))

print("\nRanking de lojas com mais vendas:")
print(ranking_loja)

#Mes com mais vendas
df["Mes"] = df["Data da Venda"].dt.to_period("M").astype(str)

ranking_meses = (df.groupby(["Mes"], as_index=False)["Qtd Vendida"] 
                 .sum() .sort_values("Qtd Vendida", ascending=False) 
                 .reset_index(drop=True))
print("\nRanking de meses com mais vendas:")
print(ranking_meses)

if not ranking_meses.empty:
    mes_mais_vendas = ranking_meses.iloc[0]["Mes"]
    print(f"\nMês com mais unidades vendidas: {mes_mais_vendas}" 
        f" — {ranking_meses.iloc[0]['Qtd Vendida']} unidades")
    
vendas_mensais = ranking_meses.sort_values("Mes")

#Tabela cronológica de vendas por mês
tabela_cronologica = vendas_mensais.copy()

tabela_cronologica.insert(0, "Ordem", range(1, len(tabela_cronologica) + 1))

# Os rankings já estão ordenados pela quantidade vendida.
top_5_produtos = top_produtos.head(5)
top_5_lojas = ranking_loja.head(5)


with pd.ExcelWriter(caminho_saida, engine="xlsxwriter") as writer:

    formato_titulo = writer.book.add_format({"bold": True, "font_color": "white", "bg_color": "#244062","font_size":12 , "align": "center", "valign": "vcenter"})

    formato_quantidade_centrado = writer.book.add_format({"align": "center" , "valign": "vcenter","num_format": "#,##0"})
    
    ranking_meses.insert(0, "Posição", range(1, len(ranking_meses) + 1))

    # Guardar as tabelas
    top_5_por_cliente.to_excel(
        writer, sheet_name="Top 5 Produtos por Cliente", startrow=2 , index=False 
    )
    ranking_loja.to_excel(
        writer, sheet_name="Ranking de Lojas", startrow=2, index=False 
    )
    top_produtos.to_excel(
        writer, sheet_name="Ranking de Produtos", startrow=2, index=False 
    )
    tabela_cronologica.to_excel(
        writer, sheet_name="Analise Mensal", startrow=2, startcol=0, index=False 
    )
    ranking_meses.to_excel(
        writer, sheet_name="Analise Mensal", startrow=2, startcol=4, index=False
    )
    #Folha Top 5 Produtos por Cliente
    folha_top5 = writer.sheets["Top 5 Produtos por Cliente"]

    folha_top5.merge_range("A1:C1", "Top 5 Produtos por Cliente", formato_titulo)
    folha_top5.set_column("A:A", 12, formato_quantidade_centrado)
    folha_top5.set_column("B:B", 12, formato_quantidade_centrado)
    folha_top5.set_column("C:C", 12, formato_quantidade_centrado)

    #Folha Ranking Lojas
    folha_lojas = writer.sheets["Ranking de Lojas"]

    folha_lojas.merge_range("A1:B1", "Ranking de Lojas", formato_titulo)
    folha_lojas.set_column("A:A", 9, formato_quantidade_centrado)
    folha_lojas.set_column("B:B", 9, formato_quantidade_centrado)

    #Folha Ranking Produtos
    folha_produtos = writer.sheets["Ranking de Produtos"]

    folha_produtos.merge_range("A1:B1", "Ranking de Produtos", formato_titulo)
    folha_produtos.set_column("A:A", 10, formato_quantidade_centrado)
    folha_produtos.set_column("B:B", 10, formato_quantidade_centrado)

    #Folha Analise mensal 
    folha_mensal = writer.sheets["Analise Mensal"]

    folha_mensal.merge_range("A1:C1","Vendas por ordem Cronológica", formato_titulo)

    folha_mensal.merge_range("E1:G1","Ranking por unidades vendidas", formato_titulo)

    folha_mensal.set_column("A:A", 8, formato_quantidade_centrado)
    folha_mensal.set_column("B:B", 10, formato_quantidade_centrado)
    folha_mensal.set_column("C:C", 10, formato_quantidade_centrado)
    folha_mensal.set_column("D:D", 3, formato_quantidade_centrado)
    folha_mensal.set_column("E:E", 8, formato_quantidade_centrado)
    folha_mensal.set_column("F:F", 10, formato_quantidade_centrado)
    folha_mensal.set_column("G:G", 10, formato_quantidade_centrado)
    folha_mensal.freeze_panes(3, 0)

    #Graficos
    workbook = writer.book
    painel = workbook.add_worksheet("Graficos")

    # 1. Unidades vendidas por mês.
    grafico_meses = workbook.add_chart({"type": "line"})

    grafico_meses.add_series({
    "name": "Unidades vendidas",
    "categories": ["Analise Mensal", 3, 1, len(vendas_mensais) + 2, 1],
    "values": ["Analise Mensal", 3, 2, len(vendas_mensais) + 2, 2],
    "marker": {"type": "circle", "size": 6},
})

    grafico_meses.set_title({"name": "Unidades vendidas por mês"})
    grafico_meses.set_x_axis({"name": "Mês"})
    grafico_meses.set_y_axis({"name": "Unidades vendidas", "min": 0})
    grafico_meses.set_legend({"none": True})
    grafico_meses.set_size({"width": 760, "height": 380})

    painel.insert_chart("B2", grafico_meses)

    # 2. Top 5 produtos.
    numero_produtos = min(5, len(top_5_produtos))
    grafico_produtos = workbook.add_chart({"type": "bar"})

    grafico_produtos.add_series({
        "name": "Unidades vendidas",
        "categories": ["Ranking de Produtos", 3, 0, numero_produtos + 2, 0],
        "values": ["Ranking de Produtos", 3, 1, numero_produtos + 2, 1],
        "data_labels": {"value": True},
    })

    grafico_produtos.set_title({"name": "Produtos com mais unidades vendidas"})
    grafico_produtos.set_x_axis({"name": "Unidades vendidas", "min": 0})
    grafico_produtos.set_y_axis({"name": "SKU", "reverse": True})
    grafico_produtos.set_legend({"none": True})
    grafico_produtos.set_size({"width": 760, "height": 380})

    painel.insert_chart("B23", grafico_produtos)

    # 3. Top 5 lojas.
    numero_lojas = min(5, len(top_5_lojas))
    grafico_lojas = workbook.add_chart({"type": "bar"})

    grafico_lojas.add_series({
        "name": "Unidades vendidas",
        "categories": ["Ranking de Lojas", 3, 0, numero_lojas + 2, 0],
        "values": ["Ranking de Lojas", 3, 1, numero_lojas + 2, 1],
        "data_labels": {"value": True},
    })

    grafico_lojas.set_title({"name": "Lojas com mais unidades vendidas"})
    grafico_lojas.set_x_axis({"name": "Unidades vendidas", "min": 0})
    grafico_lojas.set_y_axis({"name": "ID Loja", "reverse": True})
    grafico_lojas.set_legend({"none": True})
    grafico_lojas.set_size({"width": 760, "height": 380})

    painel.insert_chart("B44", grafico_lojas)

    # Mostrar a folha dos gráficos ao abrir o Excel.
    painel.activate()

print(f"\nResultados guardados em {caminho_saida}")


