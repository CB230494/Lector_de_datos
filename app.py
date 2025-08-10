import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Seguimiento por Trimestre (IT/IIT)", layout="wide")
st.title("📘 Seguimiento por Trimestre — Lector + Formulario")

# ------------------------ helpers ------------------------
def clean_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def take_cols_H_to_N(df: pd.DataFrame):
    """Devuelve nombres de columnas H..N por posición (Excel H..N → 0-based 7..13)."""
    start, end = 7, 14
    end = min(end, df.shape[1])
    return list(df.columns[start:end]) if start < end else []

def add_trimestre(df: pd.DataFrame, label: str) -> pd.DataFrame:
    df = df.copy()
    df["Trimestre"] = label
    return df

def standardize_delegacion_from_colD(df: pd.DataFrame) -> pd.DataFrame:
    """Crea columna estándar 'Delegación' SIEMPRE desde la columna D (índice 3)."""
    df = df.copy()
    if df.shape[1] > 3:
        df["Delegación"] = df.iloc[:, 3]
    else:
        df["Delegación"] = ""
    # Elimina cualquier otra columna que parezca 'Delegaciones...' o 'Delegación...'
    drop_like = [c for c in df.columns if c != "Delegación" and re.search(r"delegaci[oó]n", str(c), re.I)]
    df = df.drop(columns=drop_like)
    return df

def export_xlsx(dfs_by_sheet: dict, filename: str):
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

# ------------------------ Carga principal ------------------------
st.subheader("1) Cargar archivo base (IT y IIT)")
archivo_base = st.file_uploader("📁 Sube el Excel (el tuyo trae IT e IIT)", type=["xlsx", "xlsm"])
if not archivo_base:
    st.info("Sube el archivo para continuar.")
    st.stop()

xls = pd.ExcelFile(archivo_base)
sheet_names = xls.sheet_names

# Proponer automáticamente IT/IIT; permitir cambiar si fuera necesario
def suggest(name_list, patterns):
    for p in patterns:
        for s in name_list:
            if re.search(p, s, re.I):
                return s
    return name_list[0] if name_list else None

sheet_it  = suggest(sheet_names, [r"^it$", r"\b1t\b", r"\bprimer", r"i\s*tr"])
sheet_iit = suggest(sheet_names, [r"^iit$", r"\b2t\b", r"\bseg", r"ii\s*tr"])

col1, col2 = st.columns(2)
with col1:
    sheet_it  = st.selectbox("Hoja del I Trimestre (IT)", sheet_names, index=sheet_names.index(sheet_it) if sheet_it in sheet_names else 0)
with col2:
    sheet_iit = st.selectbox("Hoja del II Trimestre (IIT)", sheet_names, index=sheet_names.index(sheet_iit) if sheet_iit in sheet_names else min(1, len(sheet_names)-1))

# Leer las hojas
df_it  = pd.read_excel(xls, sheet_name=sheet_it)
df_iit = pd.read_excel(xls, sheet_name=sheet_iit)
df_it, df_iit = clean_cols(df_it), clean_cols(df_iit)

# Estandarizar delegación desde columna D SIEMPRE
df_it  = standardize_delegacion_from_colD(df_it)
df_iit = standardize_delegacion_from_colD(df_iit)

# Añadir rótulo de trimestre
df_it  = add_trimestre(df_it, "I")
df_iit = add_trimestre(df_iit, "II")

# Columnas H..N por posición (nombres reales)
cols_HN_it  = take_cols_H_to_N(df_it)
cols_HN_iit = take_cols_H_to_N(df_iit)
cols_HN = cols_HN_it if len(cols_HN_it) >= len(cols_HN_iit) else cols_HN_iit

# Detectar “Tipo de actividad” y “Observaciones” por nombre si existen
def find_col_by_exact(df, pat):
    for c in df.columns:
        if re.fullmatch(pat, c, flags=re.I):
            return c
    return None

col_tipo = find_col_by_exact(df_it, r"tipo\s*de\s*actividad\.?") or find_col_by_exact(df_iit, r"tipo\s*de\s*actividad\.?")
col_obs  = find_col_by_exact(df_it, r"observaciones\.?") or find_col_by_exact(df_iit, r"observaciones\.?")

# Consolidado
df_all = pd.concat([df_it, df_iit], ignore_index=True)

# ------------------------ Filtros ------------------------
st.subheader("2) Filtros")
delegaciones = sorted([d for d in df_all["Delegación"].dropna().astype(str).map(str.strip).unique() if d])
deleg_sel = st.selectbox("Delegación (columna D de cada hoja)", options=["(Todas)"] + delegaciones, index=0)
trims_sel = st.multiselect("Trimestres", options=["I","II"], default=["I","II"])

df_filtrado = df_all.copy()
if deleg_sel != "(Todas)":
    df_filtrado = df_filtrado[df_filtrado["Delegación"] == deleg_sel]
if trims_sel:
    df_filtrado = df_filtrado[df_filtrado["Trimestre"].isin(trims_sel)]

# Columnas a mostrar: Delegación, Trimestre, Tipo de actividad, Observaciones, H–N
cols_base = ["Delegación", "Trimestre"] + [c for c in [col_tipo, col_obs] if c]
cols_mostrar = cols_base + [c for c in cols_HN if c not in cols_base]

st.subheader("3) Vista rápida")
if cols_mostrar:
    st.dataframe(df_filtrado[cols_mostrar], use_container_width=True, height=420)
else:
    st.dataframe(df_filtrado, use_container_width=True, height=420)

# ------------------------ Formulario para agregar ------------------------
st.subheader("4) Agregar registros (formulario)")

with st.form("form_add"):
    c1, c2, c3 = st.columns(3)
    trim_new = c1.selectbox("Trimestre", ["I","II","III","IV"], index=0)
    pao_new  = c2.selectbox("Validación PAO", ["Sí", "No"], index=0)
    deleg_new = c3.selectbox("Delegación", delegaciones if delegaciones else [""])

    tipos_catalogo = ["Rendición de cuentas", "Seguimiento", "Líneas de acción", "Informe territorial"]
    tipo_multi = st.multiselect("Tipo de actividad (multi)", tipos_catalogo)
    tipo_new = "; ".join(tipo_multi) if tipo_multi else ""

    obs_new = st.text_area("Observaciones", height=100, placeholder="Agrega observaciones…")

    st.markdown("Completar columnas")
    valores_hn = {}
    for col in cols_HN:
        valores_hn[col] = st.text_input(col, value="")

    enviado = st.form_submit_button("➕ Agregar a la tabla")

if enviado:
    nuevo = {"Delegación": deleg_new, "Trimestre": trim_new}
    # Validación PAO: reusar si existe, si no crear
    col_pao = next((c for c in df_all.columns if re.search(r"validaci[oó]n\s*pao", c, re.I)), "Validación PAO")
    nuevo[col_pao] = pao_new
    if col_tipo: nuevo[col_tipo] = tipo_new
    if col_obs:  nuevo[col_obs]  = obs_new
    for col in cols_HN:
        nuevo[col] = valores_hn.get(col, "")
    df_all = pd.concat([df_all, pd.DataFrame([nuevo])], ignore_index=True)
    st.success("Registro agregado de forma temporal. Descárgalo abajo para guardarlo.")

# ------------------------ Generar Excel (nuevo o actualizando) ------------------------
st.subheader("5) Generar Excel nuevo o actualizar uno anterior")
modo = st.radio("¿Cómo quieres generar el archivo final?", ["Empezar uno nuevo", "Actualizar un Excel anterior"], index=0)

df_final = df_all.copy()

if modo == "Actualizar un Excel anterior":
    prev = st.file_uploader("📎 Excel anterior para combinar (opcional)", type=["xlsx","xlsm"], key="prev_x")
    if prev:
        try:
            xold = pd.ExcelFile(prev)
            frames = [pd.read_excel(xold, sheet_name=sh) for sh in xold.sheet_names]
            old_df = pd.concat(frames, ignore_index=True)
            old_df = clean_cols(old_df)

            # Estandarizar 'Delegación' también en el anterior si no existe
            if "Delegación" not in old_df.columns:
                if old_df.shape[1] > 3:
                    old_df["Delegación"] = old_df.iloc[:, 3]
                else:
                    old_df["Delegación"] = ""

            # Quitar columnas tipo delegación duplicadas
            drop_like = [c for c in old_df.columns if c != "Delegación" and re.search(r"delegaci[oó]n", str(c), re.I)]
            old_df = old_df.drop(columns=drop_like)

            df_final = pd.concat([old_df, df_all], ignore_index=True)
            st.info(f"Se combinó el archivo anterior ({len(old_df)} filas) con el actual.")
        except Exception as e:
            st.error(f"No se pudo leer el archivo anterior: {e}")

# Duplicados exactos fuera
df_final = df_final.drop_duplicates()

with st.expander("🔎 Vista previa del Excel a generar"):
    st.dataframe(df_final, use_container_width=True, height=420)

# Exportar con hojas por trimestre
sheets = {}
for t in ["I","II","III","IV"]:
    parte = df_final[df_final["Trimestre"] == t]
    if not parte.empty:
        sheets[f"{t} Trimestre"] = parte
if not sheets:
    sheets = {"Datos": df_final}

export_xlsx(sheets, filename="seguimiento_trimestres_generado.xlsx")

st.caption("Nota: El filtro y la columna 'Delegación' se construyen SIEMPRE desde la **columna D** por posición, para evitar columnas duplicadas.")









