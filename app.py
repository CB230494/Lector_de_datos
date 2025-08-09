import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Consolidado de Indicadores", layout="wide")
st.title("📋 Consolidado de Indicadores - Informe de Avance")
st.write("Carga uno o varios archivos Excel para extraer indicadores desde la hoja **'Informe de avance'**.")

# ================= Utilidades =================
def _s(x):
    """Devuelve texto limpio sin espacios repetidos."""
    return "" if pd.isna(x) else re.sub(r"\s+", " ", str(x)).strip()

def _is_zeroish(val: str) -> bool:
    """True si el valor es 0 (o 0.0, etc.)."""
    return True if re.fullmatch(r"\s*0+(\.0+)?\s*", (val or "")) else False

def _looks_like_header(lider: str, indicador: str, meta: str) -> bool:
    """Detecta filas de encabezado/rótulos dentro del bloque para no incluirlas."""
    L = (lider or "").strip().lower()
    I = (indicador or "").strip().lower()
    M = (meta or "").strip().lower()
    if (L == "lider" and I == "indicadores" and M == "meta"):
        return True
    if I in {"indicadores", "indicador", "descripción", "descripcion", "resultados", "evidencia"}:
        return True
    if M == "meta":
        return True
    return False

def _parse_single(df: pd.DataFrame, archivo: str) -> pd.DataFrame:
    """
    Parser adaptado a la plantilla observada.
    Mapa (0‑based):
      - Delegación -> (2,7)  [G3]
      - N° Líneas  -> (4,7)  [G5]
      - Por 'Línea de Acción #X' (texto en col=3):
          Problemática -> misma fila col=5
          Encabezado del bloque -> r0 + 3
          Columnas datos:
            Lider (3), Indicadores (5), Meta (7)
            Resultados T1 (13), Resultados T2 (19)
    """
    out_rows = []

    # Delegación y número de líneas
    delegacion = _s(df.iat[2, 7]) if df.shape[0] > 2 and df.shape[1] > 7 else ""
    try:
        num_lineas = int(float(_s(df.iat[4, 7])))
    except Exception:
        num_lineas = 0

    # Localizar líneas de acción
    line_rows = []
    for r in range(df.shape[0]):
        txt = _s(df.iat[r, 3]) if df.shape[1] > 3 else ""
        m = re.search(r"L[ií]nea de Acci[oó]n\s*#\s*(\d+)", txt, flags=re.I)
        if m:
            line_rows.append((r, int(m.group(1))))

    if num_lineas > 0:
        line_rows = [t for t in line_rows if t[1] <= num_lineas]

    if not line_rows:
        return pd.DataFrame(columns=[
            "Archivo","Delegación","N° Líneas","Línea #","Problemática","Líder",
            "Indicador","Meta","Resultados T1","Resultados T2"
        ])

    for idx, (r0, linea_num) in enumerate(line_rows):
        r_next = line_rows[idx + 1][0] if idx + 1 < len(line_rows) else df.shape[0]

        # Problemática
        problema = _s(df.iat[r0, 5]) if df.shape[1] > 5 else ""

        # Encabezado del bloque y columnas
        header_row = r0 + 3
        c_lider, c_ind, c_meta = 3, 5, 7
        c_res1, c_res2 = 13, 19

        # Recorrer filas de datos del bloque
        blank_streak = 0
        for r in range(header_row + 1, r_next):
            indicador = _s(df.iat[r, c_ind]) if df.shape[1] > c_ind else ""
            lider = _s(df.iat[r, c_lider]) if df.shape[1] > c_lider else ""
            meta = _s(df.iat[r, c_meta]) if df.shape[1] > c_meta else ""

            # Si reaparece una fila de encabezados/rótulos → cortar bloque
            if _looks_like_header(lider, indicador, meta):
                break

            # Si no hay indicador, contar vacíos y cortar tras 2 seguidos
            if not indicador:
                blank_streak += 1
                if blank_streak >= 2:
                    break
                continue
            blank_streak = 0

            t1 = _s(df.iat[r, c_res1]) if df.shape[1] > c_res1 else ""
            t2 = _s(df.iat[r, c_res2]) if df.shape[1] > c_res2 else ""
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

def _clean_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Limpieza final: quita encabezados repetidos y filas sin indicador útil."""
    if df.empty:
        return df

    def s(x): return "" if pd.isna(x) else str(x).strip().lower()
    bad_ind_texts = {"", "indicadores", "indicador", "descripcion", "descripción", "resultados", "evidencia"}

    mask_headerish = (
        df.get("Líder", "").map(s).eq("lider")
        & df.get("Indicador", "").map(s).eq("indicadores")
        & df.get("Meta", "").map(s).eq("meta")
    )
    mask_bad_indicator = df.get("Indicador", "").map(s).isin(bad_ind_texts)

    # Conservar solo filas buenas
    clean = df.loc[~(mask_headerish | mask_bad_indicator)].copy()

    # Opcional: quitar espacios sobrantes en todas las columnas de texto
    for col in ["Archivo","Delegación","Problemática","Líder","Indicador","Meta","Resultados T1","Resultados T2"]:
        if col in clean.columns:
            clean[col] = clean[col].apply(_s)

    # Reordenar por archivo y línea
    sort_cols = [c for c in ["Archivo","Línea #"] if c in clean.columns]
    if sort_cols:
        clean = clean.sort_values(sort_cols, kind="stable").reset_index(drop=True)

    return clean

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

    result = pd.concat(todos, ignore_index=True)
    result = _clean_rows(result)   # <<< limpieza final
    return result

# ================= UI =================
archivos = st.file_uploader("📁 Sube archivos .xlsm o .xlsx", type=["xlsm", "xlsx"], accept_multiple_files=True)

if archivos:
    df_out = procesar_informes(archivos)

    if df_out.empty:
        st.warning("No se encontraron datos con el formato esperado.")
    else:
        st.success(f"✅ Procesados {df_out['Archivo'].nunique()} archivo(s). Registros: {len(df_out)}")
        st.dataframe(df_out, use_container_width=True, height=420)

        # Resumen por línea de acción
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


