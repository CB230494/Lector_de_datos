import streamlit as st
import pandas as pd
import io

st.title("📊 Consolidado de Indicadores - DASHBOARD")
st.write("Carga archivos Excel con hoja 'DASHBOARD' desbloqueada para generar un resumen por delegación y líder estratégico.")

archivo = st.file_uploader("📁 Sube un archivo .xlsm o .xlsx", type=["xlsm", "xlsx"])

@st.cache_data
def procesar_dashboard(uploaded_file):
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
    consolidado = []

    if "DASHBOARD" not in xls.sheet_names:
        st.error("❌ El archivo no contiene una hoja llamada 'DASHBOARD'")
        return pd.DataFrame()

    df = pd.read_excel(xls, sheet_name="DASHBOARD", header=None, engine="openpyxl")

    try:
        delegacion = str(df.iloc[3, 1]).strip()

        # Datos Gobierno Local (columna 7)
        gl_completos = int(df.iloc[7, 7]) if pd.notna(df.iloc[7, 7]) else 0
        gl_con_act = int(df.iloc[8, 7]) if pd.notna(df.iloc[8, 7]) else 0
        gl_sin_act = int(df.iloc[9, 7]) if pd.notna(df.iloc[9, 7]) else 0

        # Datos Fuerza Pública (columna 7)
        fp_completos = int(df.iloc[18, 7]) if pd.notna(df.iloc[18, 7]) else 0
        fp_con_act = int(df.iloc[19, 7]) if pd.notna(df.iloc[19, 7]) else 0
        fp_sin_act = int(df.iloc[20, 7]) if pd.notna(df.iloc[20, 7]) else 0

        consolidado.append({
            "Delegación": delegacion,
            "Líder Estratégico": "Gobierno Local",
            "Completados": gl_completos,
            "Con Actividades": gl_con_act,
            "Sin Actividades": gl_sin_act
        })

        consolidado.append({
            "Delegación": delegacion,
            "Líder Estratégico": "Fuerza Pública",
            "Completados": fp_completos,
            "Con Actividades": fp_con_act,
            "Sin Actividades": fp_sin_act
        })

    except Exception as e:
        st.error(f"❌ Error al procesar la hoja 'DASHBOARD': {e}")
        return pd.DataFrame()

    return pd.DataFrame(consolidado)

# Procesamiento principal
if archivo:
    df_resultado = procesar_dashboard(archivo)

    if not df_resultado.empty:
        st.success("✅ Archivo procesado correctamente.")
        st.dataframe(df_resultado)

        # Descargar Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_resultado.to_excel(writer, index=False, sheet_name="Resumen")

        st.download_button(
            label="📥 Descargar resumen en Excel",
            data=output.getvalue(),
            file_name="resumen_dashboard.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

