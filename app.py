import streamlit as st
import pandas as pd
import io

st.title("📊 Consolidado de Indicadores - DASHBOARD")
st.write("Carga archivos Excel con la hoja 'DASHBOARD' desbloqueada para generar un resumen por delegación y líder estratégico.")

# Subida de archivo
archivo = st.file_uploader("📁 Sube un archivo .xlsm o .xlsx", type=["xlsm", "xlsx"])

@st.cache_data
def procesar_dashboard(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")

        if "DASHBOARD" not in xls.sheet_names:
            st.error("❌ El archivo no contiene una hoja llamada 'DASHBOARD'")
            return pd.DataFrame()

        df = pd.read_excel(xls, sheet_name="DASHBOARD", header=None, engine="openpyxl")

        delegacion = str(df.iloc[3, 1]).strip()

        # ✅ Leer columna 8 (índice 8) que contiene los valores enteros reales, no los porcentajes
        gl_completos = int(df.iloc[7, 8]) if pd.notna(df.iloc[7, 8]) else 0
        gl_con_act = int(df.iloc[8, 8]) if pd.notna(df.iloc[8, 8]) else 0
        gl_sin_act = int(df.iloc[9, 8]) if pd.notna(df.iloc[9, 8]) else 0

        fp_completos = int(df.iloc[18, 8]) if pd.notna(df.iloc[18, 8]) else 0
        fp_con_act = int(df.iloc[19, 8]) if pd.notna(df.iloc[19, 8]) else 0
        fp_sin_act = int(df.iloc[20, 8]) if pd.notna(df.iloc[20, 8]) else 0

        consolidado = [
            {
                "Delegación": delegacion,
                "Líder Estratégico": "Gobierno Local",
                "Completados": gl_completos,
                "Con Actividades": gl_con_act,
                "Sin Actividades": gl_sin_act
            },
            {
                "Delegación": delegacion,
                "Líder Estratégico": "Fuerza Pública",
                "Completados": fp_completos,
                "Con Actividades": fp_con_act,
                "Sin Actividades": fp_sin_act
            }
        ]

        return pd.DataFrame(consolidado)

    except Exception as e:
        st.error(f"❌ Error al procesar la hoja 'DASHBOARD': {e}")
        return pd.DataFrame()

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

