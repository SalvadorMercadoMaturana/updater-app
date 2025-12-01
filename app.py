import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

# ===========================================================
# CONFIGURACIÓN DE LA APP
# ===========================================================
st.set_page_config(page_title="Actualizador de Procesos", layout="centered")
st.title("Actualizador de Procesos")

# ============================
# PROTECCIÓN CON CONTRASEÑA
# ============================
PASSWORD = "RipleyRiesgos"

if "auth" not in st.session_state:
    st.session_state["auth"] = False

if not st.session_state["auth"]:
    pwd = st.text_input("Ingrese contraseña para acceder:", type="password")
    if pwd == PASSWORD:
        st.session_state["auth"] = True
        st.success("Acceso concedido")
    else:
        st.stop()

st.write("""
Sube el archivo Excel real. La app mantiene TODAS las hojas, macros y formatos,
y solo edita la fila seleccionada dentro de la hoja MAPA_ACTUAL_JUL_2025.
""")

# ===========================================================
# SUBIR ARCHIVO
# ===========================================================
uploaded = st.file_uploader("Sube el archivo Excel (hoja: MAPA_ACTUAL_JUL_2025)", type=["xlsx", "xlsm"])

if not uploaded:
    st.stop()

# ========= MUY IMPORTANTE =========
# Guardamos una copia ORIGINAL del archivo en memoria
# antes de consumirlo con pandas
# =================================
original_bytes = uploaded.getvalue()
buffer_for_pandas = io.BytesIO(original_bytes)

# Ahora pandas puede leer normalmente
try:
    df = pd.read_excel(buffer_for_pandas, sheet_name="MAPA_ACTUAL_JUL_2025")
except Exception as e:
    st.error("Error al cargar la hoja MAPA_ACTUAL_JUL_2025: " + str(e))
    st.stop()

df.columns = df.columns.str.strip()

# Columnas clave
COL_TIPO = "Tipo de Proceso"
COL_MACRO = "Nivel 0 - Macroproceso"
COL_PROCESO = "Nivel 1 - Proceso (Final)"
COL_SUB = "Nivel 2 -Subproceso (Final)"

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

tipos = df[COL_TIPO].dropna().unique()
tipo_sel = st.selectbox("Tipo de Proceso", tipos)

df1 = df[df[COL_TIPO] == tipo_sel]

macros = df1[COL_MACRO].dropna().unique()
macro_sel = st.selectbox("Macroproceso", macros)

df2 = df1[df1[COL_MACRO] == macro_sel]

procesos = df2[COL_PROCESO].dropna().unique()
proc_sel = st.selectbox("Proceso (Final)", procesos)

df3 = df2[df2[COL_PROCESO] == proc_sel]

subs = df3[COL_SUB].dropna().unique()
sub_sel = st.selectbox("Subproceso (Final)", subs)

df_target = df3[df3[COL_SUB] == sub_sel]

if len(df_target) != 1:
    st.error("La combinación no corresponde a una fila única.")
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

    submit = st.form_submit_button("Aplicar Cambios")

# ===========================================================
# APLICAR MODIFICACIONES EN EL EXCEL ORIGINAL
# ===========================================================
if submit:
    # Cargar el Excel original desde los bytes guardados
    wb = load_workbook(io.BytesIO(original_bytes), keep_vba=True)
    ws = wb["MAPA_ACTUAL_JUL_2025"]

    excel_row = idx + 2  # pandas 0-based + header

    columnas = list(df.columns)

    cambios = {
        COL_COM: comentarios,
        COL_AREA: area,
        COL_RESP: responsable,
        COL_EST: estado,
    }

    for col_name, new_val in cambios.items():
        col_idx = columnas.index(col_name) + 1
        ws.cell(row=excel_row, column=col_idx).value = new_val

    buffer_out = io.BytesIO()
    wb.save(buffer_out)
    buffer_out.seek(0)

    st.success("✔ Cambios aplicados correctamente (libro COMPLETO preservado).")

    st.download_button(
        "📥 Descargar Excel Actualizado",
        data=buffer_out,
        file_name="procesos_actualizado.xlsm" if uploaded.name.endswith(".xlsm") else "procesos_actualizado.xlsx",
        mime="application/vnd.ms-excel"
    )
