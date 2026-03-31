import pandas as pd

# Cargar archivo
file_path = "Motor de cálculo reinstalables_actuaria V3.xlsx"
df = pd.read_excel(file_path, sheet_name="Simplificada")

# Limpieza inicial
df_clean = df[df["ID"].notna()].copy()
df_clean["ID"] = df_clean["ID"].astype(int)

# Número de asegurados por ID
df_num_asegurados = df_clean["ID"].value_counts().reset_index()
df_num_asegurados.columns = ["TCID", "NUM ASEGURADOS"]
df_num_asegurados = df_num_asegurados.sort_values("TCID").reset_index(drop=True)

# Prima esperada
df_clean["PRIMA ESPERADA"] = df_clean["Pma + Der"]
df_prima_esperada = df_clean.groupby("ID", as_index=False)["PRIMA ESPERADA"].sum()

# Poliza
df_poliza = df_clean.groupby("ID")["Póliza"].first().reset_index()
df_poliza.columns = ["TCID", "POLIZA"]
df_poliza["POLIZA"] = df_poliza["POLIZA"].apply(lambda x: str(str(x)) if pd.notnull(x) else "")

# Tipo de Deducible
col_deducible = [c for c in df_clean.columns if "Tipo de Deducible" in c][0]
df_tipo_deducible = df_clean.groupby("ID")[col_deducible].first().reset_index()
df_tipo_deducible.columns = ["TCID", "TIPO DE DEDUCIBLE"]

# Tipo de Coaseguro
col_coaseguro = [c for c in df_clean.columns if "Tipo de Coaseguro" in c][0]
df_tipo_coaseguro = df_clean.groupby("ID")[col_coaseguro].first().reset_index()
df_tipo_coaseguro.columns = ["TCID", "TIPO DE COASEGURO"]

# Plan
col_plan = [c for c in df_clean.columns if "Plan" in c][0]
df_plan = df_clean.groupby("ID")[col_plan].first().reset_index()
df_plan.columns = ["TCID", "PLANB"]

# Tipo de cambio
col_cambio = [c for c in df_clean.columns if "Tipo de cambio" in c][0]
df_tipo_cambio = df_clean.groupby("ID")[col_cambio].first().reset_index()
df_tipo_cambio.columns = ["TCID", "TIPO DE CAMBIO"]


###########################################################################################################
# Consolidar todos los datos en el orden deseado
df_resultado = df_num_asegurados
df_resultado = df_resultado.merge(df_prima_esperada, left_on="TCID", right_on="ID").drop(columns=["ID"])
df_resultado = df_resultado.merge(df_poliza, on="TCID")
df_resultado = df_resultado.merge(df_tipo_deducible, on="TCID")
df_resultado = df_resultado.merge(df_tipo_coaseguro, on="TCID")
df_resultado["FLUJO"] = "cambios"
df_resultado["TIPO"] = "GMM"
df_resultado["RUN"] = 0
df_resultado["PLAN"] = "CAMBIO DE CONDICION"
df_resultado = df_resultado.merge(df_plan, on="TCID")
df_resultado = df_resultado.merge(df_tipo_cambio, on="TCID")

###########################################################################################################
# Exportar
df_resultado.to_csv("matriz_cambios_2026.csv", index=False)
print("✅ Archivo 'matriz_cambios.csv' generado con exito.")