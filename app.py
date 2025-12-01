import streamlit as st
import pandas as pd
import io

# ===========================================================
# CONFIGURACIÓN DE LA APP
# ===========================================================
st.set_page_config(page_title="Actualizador de Procesos", layout="centered")
st.title("📝 Actualizador de Procesos (Versión Web Simple)")

st.write("""
Sube el archivo Excel **sin datos sensibles** (solo para pruebas internas)
y edita los campos de cada proceso usando filtros desplegables.
""")

# ===========================================================
# SUBIR ARCHIVO
# ===========================================================
uploaded = st.file_uploader("Sube el archivo Excel (hoja: MAPA_ACTUAL_JUL_2025)", type=["xlsx", "xlsm"])

if not uploaded:
    st.stop()

try:
    df = pd.read_excel(uploaded, sheet_name="MAPA_ACTUAL_JUL_2025")
except Exception as e:
    st.error("❌ Error al cargar la hoja MAPA_ACTUAL_JUL_2025: " + str(e))
    st.stop()

# Normalizar columnas
df.columns = df.columns.str.strip()

# Columnas clave
COL_TIPO = "Tipo de Proceso"
COL_MACRO = "Nivel 0 - Macroproceso"
COL_PROCESO = "Nivel 1 - Proceso (Final)"
COL_SUB = "Nivel 2 - Subproceso (Final)"

# Columnas editables
COL_COM = "COMENTARIOS"
COL_AREA = "Área Responsable"
COL_RESP = "Responsable"
COL_EST = "ESTADO"
COL_FECHA = "FECHA LEVANTAMIENTO / PROGRAMADO"

# ===========================================================
# FILTROS DEPENDIENTES
# ===========================================================
st.subheader("🔎 Selecciona el proceso a modificar")

# 1. Tipo de Proceso
tipos = df[COL_TIPO].dropna().unique()
tipo_sel = st.selectbox("Tipo de Proceso", tipos)

df1 = df[df[COL_TIPO] == tipo_sel]

# 2. Macroproceso
macros = df1[COL_MACRO].dropna().unique()
macro_sel = st.selectbox("Macroproceso", macros)

df2 = df1[df1[COL_MACRO] == macro_sel]

# 3. Proceso
procesos = df2[COL_PROCESO].dropna().unique()
proc_sel = st.selectbox("Proceso (Final)", procesos)

df3 = df2[df2[COL_PROCESO] == proc_sel]

# 4. Subproceso
subs = df3[COL_SUB].dropna().unique()
sub_sel = st.selectbox("Subproceso (Final)", subs)

df_target = df3[df3[COL_SUB] == sub_sel]

if len(df_target) != 1:
    st.error("❌ La combinación no corresponde a una fila única.")
    st.stop()

idx = df_target.index[0]
row = df.loc[idx]

st.success("✔ Fila encontrada")

# ===========================================================
# FORMULARIO PARA EDITAR CAMPOS
# ===========================================================
st.subheader("✏️ Editar campos")

with st.form("form_edit"):
    comentarios = st.text_area("Comentarios", row[COL_COM])
    area = st.text_input("Área Responsable", row[COL_AREA])
    responsable = st.text_input("Responsable", row[COL_RESP])

    estado = st.selectbox(
        "Estado",
        ["EN PROCESO", "FINALIZADO"],
        index=0 if row[COL_EST] == "EN PROCESO" else 1
    )

    st.text_input("Fecha Levantamiento / Programado (NO editable)", str(row[COL_FECHA]), disabled=True)

    submit = st.form_submit_button("✅ Aplicar Cambios")

# ===========================================================
# APLICAR MODIFICACIONES
# ===========================================================
if submit:
    df.loc[idx, COL_COM] = comentarios
    df.loc[idx, COL_AREA] = area
    df.loc[idx, COL_RESP] = responsable
    df.loc[idx, COL_EST] = estado

    # Generar Excel modificado
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="MAPA_ACTUAL_JUL_2025")

    st.success("✔ Cambios aplicados")

    st.download_button(
        "📥 Descargar Excel Actualizado",
        data=buffer.getvalue(),
        file_name="procesos_actualizado.xlsx",
        mime="application/vnd.ms-excel"
    )
