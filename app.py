import streamlit as st
import pandas as pd
import io

st.title("📊 Consolidado de Indicadores - DASHBOARD")
st.write("Carga **uno o varios archivos Excel** con hoja 'DASHBOARD' desbloqueada para generar el resumen.")

# Subida múltiple de archivos
archivos = st.file_uploader("📁 Sube archivos .xlsm o .xlsx", type=["xlsm", "xlsx"], accept_multiple_files=True)

@st.cache_data
def procesar_varios_dashboards(lista_archivos):
    consolidado = []

    for archivo in lista_archivos:
        try:
            xls = pd.ExcelFile(archivo, engine="openpyxl")

            if "DASHBOARD" not in xls.sheet_names:
                st.warning(f"⚠️ El archivo '{archivo.name}' no tiene hoja 'DASHBOARD'. Se omite.")
                continue

            df = pd.read_excel(xls, sheet_name="DASHBOARD", header=None, engine="openpyxl")

            delegacion = str(df.iloc[3, 1]).strip()

            # Leer columna 8 (índice 8) con valores reales
            gl_completos = int(df.iloc[7, 8]) if pd.notna(df.iloc[7, 8]) else 0
            gl_con_act = int(df.iloc[8, 8]) if pd.notna(df.iloc[8, 8]) else 0
            gl_sin_act = int(df.iloc[9, 8]) if pd.notna(df.iloc[9, 8]) else 0

            fp_completos = int(df.iloc[18, 8]) if pd.notna(df.iloc[18, 8]) else 0
            fp_con_act = int(df.iloc[19, 8]) if pd.notna(df.iloc[19, 8]) else 0
            fp_sin_act = int(df.iloc[20, 8]) if pd.notna(df.iloc[20, 8]) else 0

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
            st.error(f"❌ Error en archivo '{archivo.name}': {e}")

    return pd.DataFrame(consolidado)

# Procesamiento principal
if archivos:
    df_resultado = procesar_varios_dashboards(archivos)

    if not df_resultado.empty:
        st.success("✅ Archivos procesados correctamente.")
        st.dataframe(df_resultado)

        # Descargar Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_resultado.to_excel(writer, index=False, sheet_name="Resumen")

        st.download_button(
            label="📥 Descargar resumen consolidado",
            data=output.getvalue(),
            file_name="resumen_dashboard_consolidado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

