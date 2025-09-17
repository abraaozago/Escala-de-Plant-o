import pandas as pd

# -----------------------------
# Caminhos dos arquivos
# -----------------------------
arquivo_delegados = r"C:\Users\user\Desktop\Escala-de-Plant-o\delegados.xlsx"
arquivo_escala = r"C:\Users\user\Desktop\Escala-de-Plant-o\ESCALA_FINAL_NOMES_COMPLETA.xlsx"
saida = r"C:\Users\user\Desktop\Escala-de-Plant-o\conflitos_ferias_nomes.xlsx"

# -----------------------------
# Lê delegados e férias
# -----------------------------
df_delegados = pd.read_excel(arquivo_delegados)

# Converte colunas de datas
for col in ["Inicio Férias 1", "Término Férias 1", "Inicio Férias 2", "Término Férias 2"]:
    df_delegados[col] = pd.to_datetime(df_delegados[col], errors="coerce")

# -----------------------------
# Lê escala com nomes
# -----------------------------
df_escala = pd.read_excel(arquivo_escala)

# Pega a coluna de datas pela posição (coluna A = índice 0)
df_escala["Data"] = pd.to_datetime(df_escala.iloc[:, 0], errors="coerce")

# -----------------------------
# Verifica conflitos
# -----------------------------
conflitos = []

for _, escala in df_escala.iterrows():
    data = escala["Data"]

    # Pega diurno e noturno pelas posições (coluna C e D)
    for plantao, idx_col in [("Diurno", 2), ("Noturno", 3)]:
        nome = escala.iloc[idx_col]
        if pd.isna(nome):
            continue

        # Busca férias desse delegado pelo nome
        ferias = df_delegados[df_delegados["Nome"].str.strip() == str(nome).strip()]

        if not ferias.empty:
            for _, f in ferias.iterrows():
                # Período 1
                if pd.notna(f["Inicio Férias 1"]) and pd.notna(f["Término Férias 1"]):
                    if f["Inicio Férias 1"].date() <= data.date() <= f["Término Férias 1"].date():
                        conflitos.append([nome, data.date(), plantao, "Período 1"])
                # Período 2
                if pd.notna(f["Inicio Férias 2"]) and pd.notna(f["Término Férias 2"]):
                    if f["Inicio Férias 2"].date() <= data.date() <= f["Término Férias 2"].date():
                        conflitos.append([nome, data.date(), plantao, "Período 2"])

# -----------------------------
# Resultado
# -----------------------------
df_conflitos = pd.DataFrame(conflitos, columns=["Nome", "Data", "Plantão", "Período de Férias"])

# Salva em Excel
df_conflitos.to_excel(saida, index=False)

print(f"✅ Análise concluída. Conflitos salvos em: {saida}")
