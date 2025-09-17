import pandas as pd

# Caminhos dos arquivos
arquivo_delegados = r"C:\Users\user\Desktop\Escala-de-Plant-o\delegados.xlsx"
arquivo_escala = r"C:\Users\user\Desktop\Escala-de-Plant-o\ESCALA 2º Semestre 2025 plantonistas.xlsx"

# --- Lê delegados e férias ---
df_delegados = pd.read_excel(arquivo_delegados)

# Converte colunas de datas
for col in ["Inicio Férias 1", "Término Férias 1", "Inicio Férias 2", "Término Férias 2"]:
    df_delegados[col] = pd.to_datetime(df_delegados[col], errors="coerce")

# --- Lê escala ---
df_escala = pd.read_excel(arquivo_escala)

# Converte a coluna de datas da escala (troque se o nome for diferente)
df_escala["Data"] = pd.to_datetime(df_escala["Data"], errors="coerce")

# Colunas dos plantões na escala
col_plantao_diurno = "diurno"
col_plantao_noturno = "noturno"

# --- Verifica conflitos ---
conflitos = []

for _, escala in df_escala.iterrows():
    data = escala["Data"]

    for plantao, col in [("Diurno", col_plantao_diurno), ("Noturno", col_plantao_noturno)]:
        codigo = escala[col]

        if pd.isna(codigo):
            continue

        # Busca férias desse delegado
        ferias = df_delegados[df_delegados["Código"] == codigo]

        if not ferias.empty:
            for _, f in ferias.iterrows():
                # Período 1
                if pd.notna(f["Inicio Férias 1"]) and pd.notna(f["Término Férias 1"]):
                    if f["Inicio Férias 1"] <= data <= f["Término Férias 1"]:
                        conflitos.append([codigo, data, plantao, "Período 1"])

                # Período 2
                if pd.notna(f["Inicio Férias 2"]) and pd.notna(f["Término Férias 2"]):
                    if f["Inicio Férias 2"] <= data <= f["Término Férias 2"]:
                        conflitos.append([codigo, data, plantao, "Período 2"])

# --- Resultado ---
df_conflitos = pd.DataFrame(conflitos, columns=["Código", "Data", "Plantão", "Período de Férias"])

# Salva em Excel
saida = r"C:\Users\user\Desktop\Escala-de-Plant-o\conflitos_ferias.xlsx"
df_conflitos.to_excel(saida, index=False)

print(f"✅ Análise concluída. Conflitos salvos em: {saida}")
