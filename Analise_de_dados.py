import pandas as pd
from datetime import datetime

# === Carrega os dados ===
df = pd.read_csv("dados_dos_imoveis.csv")

# === Pré-processamento ===
df.columns = df.columns.str.strip()
df = df.dropna(subset=["Código", "Status / Datas"])
df["Código"] = df["Código"].astype(str)

# === Extração de informações da coluna "Status / Datas" ===
df["Cliques"] = df["Status / Datas"].str.extract(r"Cliques (\d+)").astype(float)
df["Status"] = df["Status / Datas"].str.extract(r"^(Ativo|Inativo)")
df["Cadastrado"] = df["Status / Datas"].str.extract(r"Cadastrado (\d{2}/\d{2}/\d{4})")[0]
df["Atualizado"] = df["Status / Datas"].str.extract(r"Atualizado (\d{2}/\d{2}/\d{4})")[0]
df["DiasSinceCadastro"] = (datetime.now() - pd.to_datetime(df["Cadastrado"], format="%d/%m/%Y")).dt.days
df["Engajamento"] = df["Cliques"] / df["DiasSinceCadastro"]

# === Classificação da demanda ===
def classificar_demanda(dias, cliques):
    if dias <= 90:
        alta = 40
    elif dias <= 180:
        alta = 80
    else:
        alta = 100

    if cliques >= alta:
        return "Alta"
    elif cliques >= alta - 10:
        return "Boa"
    elif cliques >= alta - 20:
        return "Normal"
    elif cliques >= alta - 40:
        return "Baixa"
    else:
        return "Muito Baixa"

df["Demanda"] = df.apply(lambda row: classificar_demanda(row["DiasSinceCadastro"], row["Cliques"]), axis=1)

# === Top 10 imóveis mais acessados ===
top_10 = df.sort_values(by="Cliques", ascending=False).head(10)
top_10.to_excel("Top_10_Imoveis_Mais_Acessados.xlsx", index=False)

# === Relatório analítico completo ===
analitico = df[[
    "Código", "Título", "Condomínio", "Referência", "Status", "Cadastrado", "Atualizado",
    "Cliques", "DiasSinceCadastro", "Engajamento", "Demanda"
]]
analitico.to_excel("Relatorio_Analitico_Imoveis.xlsx", index=False)

# === Estatísticas descritivas ===
estatisticas = {
    "Total de Imóveis": len(df),
    "Ativos": (df["Status"] == "Ativo").sum(),
    "Inativos": (df["Status"] == "Inativo").sum(),
    "Tempo Médio desde Cadastro (dias)": round(df["DiasSinceCadastro"].mean(), 2),
    "Desvio Padrão de Dias": round(df["DiasSinceCadastro"].std(), 2),
    "Média Geral de Cliques": round(df["Cliques"].mean(), 2),
    "Maior número de Cliques": int(df["Cliques"].max()),
    "Menor número de Cliques": int(df["Cliques"].min()),
    "Imóvel Mais Antigo (dias)": int(df["DiasSinceCadastro"].max()),
    "Imóvel Mais Recente (dias)": int(df["DiasSinceCadastro"].min()),
    "Maior Engajamento (cliques/dia)": round(df["Engajamento"].max(), 2),
    "Menor Engajamento (cliques/dia)": round(df["Engajamento"].min(), 2)
}

estatisticas_por_status = df.groupby("Status")["Cliques"].agg(["mean", "median", "std", "min", "max"]).reset_index()
estatisticas_por_status.columns = ["Status", "Média", "Mediana", "Desvio Padrão", "Mínimo", "Máximo"]

distribuicao_demanda = df["Demanda"].value_counts().reset_index()
distribuicao_demanda.columns = ["Demanda", "Quantidade"]

# Extras para insights
mais_cliques = df[df["Cliques"] == df["Cliques"].max()]
mais_engajado = df[df["Engajamento"] == df["Engajamento"].max()]
mais_antigo = df[df["DiasSinceCadastro"] == df["DiasSinceCadastro"].max()]
mais_recente = df[df["DiasSinceCadastro"] == df["DiasSinceCadastro"].min()]

# === Exporta estatísticas em múltiplas abas ===
with pd.ExcelWriter("Estatisticas_Descritivas_Expandida.xlsx", engine="openpyxl") as writer:
    pd.DataFrame(list(estatisticas.items()), columns=["Indicador", "Valor"]).to_excel(writer, sheet_name="Resumo Geral", index=False)
    estatisticas_por_status.to_excel(writer, sheet_name="Por Status", index=False)
    distribuicao_demanda.to_excel(writer, sheet_name="Por Demanda", index=False)
    mais_cliques.to_excel(writer, sheet_name="Maior Cliques", index=False)
    mais_engajado.to_excel(writer, sheet_name="Maior Engajamento", index=False)
    mais_antigo.to_excel(writer, sheet_name="Mais Antigo", index=False)
    mais_recente.to_excel(writer, sheet_name="Mais Recente", index=False)

print("✅ Relatórios finalizados com sucesso!")
