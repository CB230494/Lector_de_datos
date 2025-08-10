# app.py
import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Seguimiento por Trimestre", layout="wide")
st.title("📘 Seguimiento por Trimestre (1T y 2T) — Lector + Formulario")

st.markdown(
    "1) Sube el archivo con **dos hojas** (1er y 2do trimestre). "
    "2) Filtra por **Delegaciones 2**. "
    "3) Agrega registros con el **formulario** y descarga un Excel nuevo o actualizado."
)

# ------------------------ helpers ------------------------
def clean_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def find_col(df: pd.DataFrame, target: str):
    """Busca columna por nombre exacto o aproximado (case-insensitive)."""
    cols = list(df.columns)
    for c in cols:
        if c.lower().strip() == target.lower().strip():
            return c
    # aproximado
    rx = re.compile(rf"{re.escape(target)}", re.I)
    for c in cols:
        if rx.search(c):
            return c
    return None

def take_H_to_N(df: pd.DataFrame):
    """Devuelve las columnas H..N (posición 8..14 en Excel → 0-based 7..13) que existan."""
    start, end = 7, 14  # iloc slicing: 7..13
    # si el df tiene menos columnas, acotar
    end = min(end, df.shape[1])
    if start >= end:
        return []
    return list(df.columns[start:end])

def add_trim_label(df: pd.DataFrame, label: str) -> pd.DataFrame:
    df = df.copy()
    df["Trimestre"] = label
    return df

def export_xlsx(dfs_by_sheet: dict, filename: str = "seguimiento_trimestres.xlsx"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet, dfx in dfs_by_sheet.items():
            dfx.to_excel(writer, index=False, sheet_name=sheet[:31])
    st.download_button(
        "📥 Descargar Excel",
        data=output.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ------------------------ carga del archivo principal ------------------------
st.subheader("1) Cargar archivo base (1T y 2T)")
archivo_base = st.file_uploader("📁 Sube el Excel con 1er y 2do trimestre", type=["xlsx", "xlsm"])

if not archivo_base:
    st.info("Sube el archivo de 1er/2do trimestre para continuar.")
    st.stop()

# Detectar hojas
xls = pd.ExcelFile(archivo_base)
sheet_names = xls.sheet_names

# Heurística de nombres (permite escoger si no coincide)
def guess_sheet(patterns):
    for p in patterns:
        for s in sheet_names:
            if re.search(p, s, re.I):
                return s
    return None

sheet_1t = guess_sheet([r"^1", r"primer", r"i\s*trim"])
sheet_2t = guess_sheet([r"^2", r"seg", r"ii\s*trim"])

col1, col2 = st.columns(2)
with col1:
    sheet_1t = st.selectbox("Hoja del 1er Trimestre", sheet_names, index=sheet_names.index(sheet_1t) if sheet_1t in sheet_names else 0)
with col2:
    sheet_2t = st.selectbox("Hoja del 2do Trimestre", sheet_names, index=sheet_names.index(sheet_2t) if sheet_2t in sheet_names else min(1, len(sheet_names)-1))

# Leer hojas con encabezado en la primera fila
df1 = pd.read_excel(xls, sheet_name=sheet_1t)
df2 = pd.read_excel(xls, sheet_name=sheet_2t)

df1, df2 = clean_cols(df1), clean_cols(df2)
df1, df2 = add_trim_label(df1, "I"), add_trim_label(df2, "II")

# Detectar columnas clave
col_deleg = find_col(df1, "Delegaciones 2") or find_col(df2, "Delegaciones 2")
col_tipo  = find_col(df1, "Tipo de actividad") or find_col(df2, "Tipo de actividad")
col_obs   = find_col(df1, "Observaciones") or find_col(df2, "Observaciones")

cols_HN_1 = take_H_to_N(df1)
cols_HN_2 = take_H_to_N(df2)
cols_HN = cols_HN_1 if len(cols_HN_1) >= len(cols_HN_2) else cols_HN_2  # toma el bloque más largo

# Consolidado
df_all = pd.concat([df1, df2], ignore_index=True)

# ------------------------ filtros ------------------------
st.subheader("2) Filtros")
if not col_deleg:
    st.error("No se encontró la columna **Delegaciones 2** en el archivo. Renómbrala exactamente o revisa la hoja.")
    st.stop()

delegaciones = sorted([d for d in df_all[col_deleg].dropna().unique().tolist() if str(d).strip() != ""])
deleg_sel = st.selectbox("Delegación (columna 'Delegaciones 2')", options=["(Todas)"] + delegaciones, index=0)

trims_sel = st.multiselect("Trimestres", options=["I","II"], default=["I","II"])

df_filtrado = df_all.copy()
if deleg_sel != "(Todas)":
    df_filtrado = df_filtrado[df_filtrado[col_deleg] == deleg_sel]
if trims_sel:
    df_filtrado = df_filtrado[df_filtrado["Trimestre"].isin(trims_sel)]

# Selección de columnas a mostrar
cols_base = [c for c in [col_deleg, "Trimestre", col_tipo, col_obs] if c]
cols_mostrar = cols_base + [c for c in cols_HN if c not in cols_base]

st.subheader("3) Vista rápida")
if cols_mostrar:
    st.dataframe(df_filtrado[cols_mostrar], use_container_width=True, height=420)
else:
    st.dataframe(df_filtrado, use_container_width=True, height=420)

# ------------------------ formulario para agregar registros ------------------------
st.subheader("4) Agregar registros (formulario)")

with st.form("agregar_registro"):
    c1, c2, c3 = st.columns(3)
    trim_new = c1.selectbox("Trimestre", ["I","II","III","IV"], index=0)
    vao_new  = c2.selectbox("Validación PAO", ["Sí", "No"], index=0)
    deleg_new = c3.selectbox("Delegaciones 2", delegaciones if delegaciones else [""])

    # Tipo de actividad (multi)
    tipos_catalogo = ["Rendición de cuentas", "Seguimiento", "Líneas de acción", "Informe territorial"]
    tipo_multi = st.multiselect("Tipo de actividad", tipos_catalogo)
    tipo_new = "; ".join(tipo_multi) if tipo_multi else ""

    # Observaciones
    obs_new = st.text_area("Observaciones", height=100, placeholder="Agrega observaciones…")

    st.markdown("**Completar columnas H–N**")
    valores_hn = {}
    for col in cols_HN:
        valores_hn[col] = st.text_input(col, value="")

    enviado = st.form_submit_button("➕ Agregar a la tabla")

if enviado:
    nuevo = {col_deleg: deleg_new, "Trimestre": trim_new}
    # si existen nombres canónicos, agrégalos
    if col_tipo: nuevo[col_tipo] = tipo_new
    if col_obs:  nuevo[col_obs]  = obs_new
    # Validación PAO: si existe la columna agrégala; si no, créala
    col_pao = find_col(df_all, "Validación PAO") or "Validación PAO"
    nuevo[col_pao] = vao_new

    # columnas H..N
    for col in cols_HN:
        nuevo[col] = valores_hn.get(col, "")

    # agrega fila
    df_all = pd.concat([df_all, pd.DataFrame([nuevo])], ignore_index=True)
    st.success("Registro agregado temporalmente. No olvides **descargar** el Excel más abajo.")

# ------------------------ actualizar desde excel anterior (opcional) ------------------------
st.subheader("5) Generar Excel nuevo o actualizar con uno anterior")

modo = st.radio("¿Cómo quieres generar el archivo final?", ["Empezar uno nuevo", "Actualizar un Excel anterior"], index=0)

df_final = df_all.copy()

if modo == "Actualizar un Excel anterior":
    prev = st.file_uploader("📎 Excel anterior (opcional)", type=["xlsx","xlsm"], key="prev_x")
    if prev:
        try:
            xold = pd.ExcelFile(prev)
            # si trae varias hojas, unimos todas
            frames = []
            for sh in xold.sheet_names:
                frames.append(pd.read_excel(xold, sheet_name=sh))
            old_df = pd.concat(frames, ignore_index=True)
            old_df = clean_cols(old_df)
            # unimos viejo + nuevo
            df_final = pd.concat([old_df, df_all], ignore_index=True)
            st.info(f"Archivo anterior detectado ({len(old_df)} filas). Se **sumó** al actual.")
        except Exception as e:
            st.error(f"No se pudo leer el archivo anterior: {e}")

# Quitar duplicados exactos
df_final = df_final.drop_duplicates()

# Vista previa final (filtrable)
with st.expander("🔎 Vista previa del Excel a generar"):
    st.dataframe(df_final, use_container_width=True, height=420)

# Exportar con hojas por trimestre (I, II, III, IV si existieran)
sheets = {}
for t in ["I","II","III","IV"]:
    parte = df_final[df_final["Trimestre"]==t]
    if not parte.empty:
        sheets[f"{t} Trimestre"] = parte

if not sheets:
    sheets = {"Datos": df_final}

export_xlsx(sheets, filename="seguimiento_trimestres_generado.xlsx")

st.caption("Sugerencias: si más adelante cambian los nombres de columnas, ajusta el **select** de hojas y verifica que exista ‘Delegaciones 2’.")





