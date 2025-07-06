import streamlit as st
import pandas as pd
import io

st.title("📊 Consolidado de Indicadores por Delegación")
st.write("Carga un archivo Excel con múltiples hojas para extraer automáticamente el avance de líneas de acción por delegación y líder estratégico.")

# Subida de archivo
archivo = st.file_uploader("Sube el archivo .xlsm", type=["xlsm", "xlsx"])

@st.cache_data
def procesar_archivo_excel(uploaded_file):
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
    consolidado = []

    for hoja in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=hoja, engine="openpyxl")

            # Delegación (columna 7, fila 1)
            delegacion = df.iloc[1, 7] if pd.notna(df.iloc[1, 7]) else hoja

            for i in range(len(df)):
                lider = df.iloc[i, 3]
                if lider in ["Municipalidad", "Fuerza Pública"]:
                    fila_inicio = i + 1
                    for j in range(fila_inicio, len(df)):
                        estado = df.iloc[j, -1]
                        if pd.isna(estado):
                            fila_fin = j
                            break
                    else:
                        fila_fin = len(df)

                    sub_df = df.iloc[fila_inicio:fila_fin]
                    completados = (sub_df.iloc[:, -1] == "Completado").sum()
                    con_actividades = (sub_df.iloc[:, -1] == "Con actividades").sum()
                    sin_actividades = (sub_df.iloc[:, -1] == "Sin actividades").sum()

                    consolidado.append({
                        "Delegación": delegacion,
                        "Líder Estratégico": lider,
                        "Completados": completados,
                        "Con Actividades": con_actividades,
                        "Sin Actividades": sin_actividades
                    })

        except Exception as e:
            st.warning(f"⚠️ Error al procesar la hoja '{hoja}': {e}")

    return pd.DataFrame(consolidado)

if archivo:
    df_resultado = procesar_archivo_excel(archivo)
    st.success("✅ Archivo procesado correctamente.")
    st.dataframe(df_resultado)

    # Descargar Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_resultado.to_excel(writer, index=False, sheet_name="Resumen")

    st.download_button(
        label="📥 Descargar resumen en Excel",
        data=output.getvalue(),
        file_name="resumen_indicadores.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
