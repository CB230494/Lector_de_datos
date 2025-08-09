import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Consolidado de Indicadores", layout="wide")
st.title("📋 Consolidado de Indicadores - Informe de Avance")
st.write("Carga uno o varios archivos Excel para extraer indicadores desde la hoja **'Informe de avance'**.")

# ================= Utilidades =================
def _s(x):
    """Texto limpio."""
    return "" if pd.isna(x) else re.sub(r"\s+", " ", str(x)).strip()

def _is_zeroish(val: str) -> bool:
    return True if re.fullmatch(r"\s*0+(\.0+)?\s*", val or "") else False

def _parse_single(df: pd.DataFrame, archivo: str) -> pd.DataFrame:
    """
    Parser adaptado a la plantilla observada.
    Mapa de columnas (0‑based) por bloque:
      - Encabezado global:
          Delegación -> (2,7)  [G3]
          N° Líneas  -> (4,7)  [G5]
      - Por "Línea de Acción #X" (fila con texto en col=3):
          Problemática -> misma fila, col=5
          Fila de encabezados del bloque -> r0 + 3
          Columnas datos:
              Lider (3), Indicadores (5), Meta (7)
              Resultados T1 (13), Resultados T2 (19)
    """
    out_rows = []

    # Delegación y número de líneas (según tu hoja de ejemplo)
    delegacion = _s(df.iat[2, 7]) if df.shape[0] > 2 and df.shape[1] > 7 else ""
    try:
        num_lineas = int(float(_s(df.iat[4, 7])))
    except Exception:
        num_lineas = 0

    # Localizar las filas donde inicia cada "Línea de Acción #"
    line_rows = []
    for r in range(df.shape[0]):
        txt = _s(df.iat[r, 3]) if df.shape[1] > 3 else ""
        m = re.search(r"L[ií]nea de Acci[oó]n\s*#\s*(\d+)", txt, flags=re.I)
        if m:
            line_rows.append((r, int(m.group(1))))

    # Limitar a las líneas indicadas en el encabezado
    if num_lineas > 0:
        line_rows = [t for t in line_rows if t[1] <= num_lineas]

    # Si no encontramos nada, devolvemos vacío
    if not line_rows:
        return pd.DataFrame(columns=[
            "Archivo","Delegación","N° Líneas","Línea #","Problemática","Líder",
            "Indicador","Meta","Resultados T1","Resultados T2"
        ])

    for idx, (r0, linea_num) in enumerate(line_rows):
        r_next = line_rows[idx + 1][0] if idx + 1 < len(line_rows) else df.shape[0]

        # Problemática (misma fila, col=5)
        problema = _s(df.iat[r0, 5]) if df.shape[1] > 5 else ""

        # Encabezado del bloque suele estar 3 filas abajo del rótulo
        header_row = r0 + 3
        c_lider, c_ind, c_meta = 3, 5, 7
        c_res1, c_res2 = 13, 19

        # Recorremos filas de datos del bloque hasta el inicio del siguiente
        for r in range(header_row + 1, r_next):
            # Indicador es clave para identificar una fila válida
            indicador = _s(df.iat[r, c_ind]) if df.shape[1] > c_ind else ""
            if not indicador:
                continue

            lider = _s(df.iat[r, c_lider]) if df.shape[1] > c_lider else ""
            meta = _s(df.iat[r, c_meta]) if df.shape[1] > c_meta else ""
            t1 = _s(df.iat[r, c_res1]) if df.shape[1] > c_res1 else ""
            t2 = _s(df.iat[r, c_res2]) if df.shape[1] > c_res2 else ""

            # Regla: si es 0, queda en blanco
            t1 = "" if _is_zeroish(t1) else t1
            t2 = "" if _is_zeroish(t2) else t2

            out_rows.append({
                "Archivo": archivo,
                "Delegación": delegacion,
                "N° Líneas": num_lineas,
                "Línea #": linea_num,
                "Problemática": problema,
                "Líder": lider,
                "Indicador": indicador,
                "Meta": meta,
                "Resultados T1": t1,
                "Resultados T2": t2,
            })

    return pd.DataFrame(out_rows)

@st.cache_data
def procesar_informes(files) -> pd.DataFrame:
    todos = []
    for f in files:
        try:
            xls = pd.ExcelFile(f, engine="openpyxl")
            if "Informe de avance" not in xls.sheet_names:
                st.warning(f"⚠️ '{f.name}' no tiene la hoja 'Informe de avance'. Se omite.")
                continue

            df = pd.read_excel(xls, sheet_name="Informe de avance", header=None, engine="openpyxl")
            df_parsed = _parse_single(df, f.name)
            todos.append(df_parsed)
        except Exception as e:
            st.error(f"❌ Error procesando '{f.name}': {e}")

    if not todos:
        return pd.DataFrame()
    return pd.concat(todos, ignore_index=True)

# ================= UI =================
archivos = st.file_uploader("📁 Sube archivos .xlsm o .xlsx", type=["xlsm", "xlsx"], accept_multiple_files=True)

if archivos:
    df_out = procesar_informes(archivos)

    if df_out.empty:
        st.warning("No se encontraron datos con el formato esperado.")
    else:
        st.success(f"✅ Procesados {df_out['Archivo'].nunique()} archivo(s). Registros: {len(df_out)}")
        st.dataframe(df_out, use_container_width=True, height=420)

        # Resumen por línea
        with st.expander("📌 Resumen por línea de acción"):
            resumen = (
                df_out.groupby(["Archivo","Delegación","Línea #","Problemática"], dropna=False)["Indicador"]
                .count()
                .reset_index(name="Total Indicadores")
                .sort_values(["Archivo","Línea #"])
            )
            st.dataframe(resumen, use_container_width=True)

        # Descargar Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_out.to_excel(writer, index=False, sheet_name="Detalle")
            if 'resumen' in locals() and not resumen.empty:
                resumen.to_excel(writer, index=False, sheet_name="Resumen por línea")
        st.download_button(
            "📥 Descargar Excel",
            data=output.getvalue(),
            file_name="resumen_informe_avance.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Sube uno o varios archivos para comenzar.")

