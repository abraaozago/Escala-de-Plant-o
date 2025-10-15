import pandas as pd
from datetime import datetime

# -----------------------------
# Arquivos
# -----------------------------
arquivo_escala = r"C:\Users\user\Desktop\Escala-de-Plant-o\ESCALA_FINAL_NOMES_COMPLETA_DIURNO.xlsx"
saida = r"C:\Users\user\Desktop\Escala-de-Plant-o\quantitativo_plantoes.xlsx"

# -----------------------------
# Lê planilha
# -----------------------------
df = pd.read_excel(arquivo_escala)

# Ajusta nome das colunas (A=Data, C=Diurno, D=Noturno)
df = df.rename(columns={
    df.columns[0]: "Data",
    df.columns[2]: "Diurno",
    df.columns[3]: "Noturno"
})

# Converte datas
df["Data"] = pd.to_datetime(df["Data"], errors="coerce")

# -----------------------------
# Inicializa estrutura de contagem
# -----------------------------
contagem = {}

def add_contagem(nome, tipo):
    if not nome or pd.isna(nome):
        return
    if nome not in contagem:
        contagem[nome] = {"Semana_Diurno": 0, "Semana_Noturno": 0,
                          "FDS_Diurno": 0, "FDS_Noturno": 0}
    contagem[nome][tipo] += 1

# -----------------------------
# Percorre cada linha e conta
# -----------------------------
for _, row in df.iterrows():
    data = row["Data"]
    if pd.isna(data):
        continue

    dia_semana = data.weekday()  # 0=segunda ... 6=domingo
    diurno = row["Diurno"]
    noturno = row["Noturno"]

    # Regras diferenciando sexta noturno
    if dia_semana < 4:  # segunda (0) até quinta (3)
        add_contagem(diurno, "Semana_Diurno")
        add_contagem(noturno, "Semana_Noturno")

    elif dia_semana == 4:  # sexta
        add_contagem(diurno, "Semana_Diurno")   # sexta diurno = semana
        add_contagem(noturno, "FDS_Noturno")    # sexta noturno = fim de semana

    elif dia_semana in (5, 6):  # sábado ou domingo
        add_contagem(diurno, "FDS_Diurno")
        add_contagem(noturno, "FDS_Noturno")

# -----------------------------
# Salvar resultado
# -----------------------------
df_resultado = pd.DataFrame.from_dict(contagem, orient="index")
df_resultado = df_resultado.reset_index().rename(columns={"index": "Delegado"})

# soma total de plantões por delegado
df_resultado = pd.DataFrame.from_dict(contagem, orient="index")
df_resultado = df_resultado.reset_index().rename(columns={"index": "Delegado"})

# soma total de plantões por delegado
colunas_numericas = ["Semana_Diurno", "Semana_Noturno", "FDS_Diurno", "FDS_Noturno"]
df_resultado["Total"] = df_resultado[colunas_numericas].sum(axis=1)

# exportar
df_resultado.to_excel(saida, index=False)
print(f"✅ Quantitativo de plantões salvo em: {saida}") 