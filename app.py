import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Seguimiento por Trimestre", layout="wide")
st.title("📘 Seguimiento por Trimestre — Lector + Formulario (Delegación = Columna D)")

# ------------------------ helpers ------------------------
def clean_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def take_H_to_N(df: pd.DataFrame):
    """Columnas H..N por posición (Excel H..N → 0-based 7..13)."""
    start, end = 7, 14
    end = min(end, df.shape[1])
    return list(df.columns[start:end]) if start < end else []

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

# ------------------------ carga del archivo base ------------------------
st.subheader("1) Cargar archivo base (1T y 2T)")
archivo_base = st.file_uploader("📁 Sube el Excel con 1er y 2do trimestre", type=["xlsx", "xlsm"])
if not archivo_base:
    st.info("Sube el archivo de 1er/2do trimestre para continuar.")
    st.stop()

xls = pd.ExcelFile(archivo_base)
sheet_names = xls.sheet_names

def guess_sheet(patterns):
    for p in patterns:
        for s in sheet_names:
            if re.search(p, s, re.I):
                return s
    return sheet_names[0]

sheet_1t = guess_sheet([r"^1", r"primer", r"i\s*trim"])
sheet_2t = guess_sheet([r"^2", r"seg", r"ii\s*trim"])

col1, col2 = st.columns(2)
with col1:
    sheet_1t = st.selectbox("Hoja del 1er Trimestre", sheet_names, index=sheet_names.index(sheet_1t))
with col2:
    sheet_2t = st.selectbox("Hoja del 2do Trimestre", sheet_names, index=sheet_names.index(sheet_2t) if sheet_2t in sheet_names else 0)

df1 = pd.read_excel(xls, sheet_name=sheet_1t)
df2 = pd.read_excel(xls, sheet_name=sheet_2t)
df1, df2 = clean_cols(df1), clean_cols(df2)

# ===== Delegación SIEMPRE desde columna D (índice 3) =====
def standardize_deleg_column(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # crea columna Delegación desde la columna D real
    if df.shape[1] > 3:
        df["Delegación"] = df.iloc[:, 3]
    else:
        df["Delegación"] = ""
    # elimina cualquier otra columna que parezca "Delegac..."
    drop_like = [c for c in df.columns if c != "Delegación" and re.search(r"delegaci[oó]n", str(c), re.I)]
    df = df.drop(columns=drop_like)
    return df

df1 = standardize_deleg_column(df1)
df2 = standardize_deleg_column(df2)

# Trimestre
df1 = add_trim_label(df1, "I")
df2 = add_trim_label(df2, "II")

# Columnas H..N (por posición) para armar formulario
cols_HN_1 = take_H_to_N(df1)
cols_HN_2 = take_H_to_N(df2)
cols_HN = cols_HN_1 if len(cols_HN_1) >= len(cols_HN_2) else cols_HN_2

# Consolidado
df_all = pd.concat([df1, df2], ignore_index=True)

# ------------------------ filtros ------------------------
st.subheader("2) Filtros")
delegaciones = sorted([d for d in df_all["Delegación"].dropna().astype(str).map(str.strip).unique() if d != ""])
deleg_sel = st.selectbox("Delegación (de la columna D)", options=["(Todas)"] + delegaciones, index=0)
trims_sel = st.multiselect("Trimestres", options=["I", "II"], default=["I", "II"])

df_filtrado = df_all.copy()
if deleg_sel != "(Todas)":
    df_filtrado = df_filtrado[df_filtrado["Delegación"] == deleg_sel]
if trims_sel:
    df_filtrado = df_filtrado[df_filtrado["Trimestre"].isin(trims_sel)]

# Intentar detectar "Tipo de actividad" y "Observaciones" por nombre (si existen)
col_tipo = next((c for c in df_all.columns if re.fullmatch(r"tipo\s*de\s*actividad", c, re.I)), None)
col_obs  = next((c for c in df_all.columns if re.fullmatch(r"observaciones?", c, re.I)), None)

cols_base = ["Delegación", "Trimestre"] + [c for c in [col_tipo, col_obs] if c]
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
    trim_new = c1.selectbox("Trimestre", ["I", "II", "III", "IV"], index=0)
    vao_new  = c2.selectbox("Validación PAO", ["Sí", "No"], index=0)
    deleg_new = c3.selectbox("Delegación", delegaciones if delegaciones else [""])

    tipos_catalogo = ["Rendición de cuentas", "Seguimiento", "Líneas de acción", "Informe territorial"]
    tipo_multi = st.multiselect("Tipo de actividad (multi)", tipos_catalogo)
    tipo_new = "; ".join(tipo_multi) if tipo_multi else ""

    obs_new = st.text_area("Observaciones", height=100, placeholder="Agrega observaciones…")

    st.markdown("**Completar columnas H–N**")
    valores_hn = {}
    for col in cols_HN:
        valores_hn[col] = st.text_input(col, value="")

    enviado = st.form_submit_button("➕ Agregar a la tabla")

if enviado:
    nuevo = {"Delegación": deleg_new, "Trimestre": trim_new}
    # Validación PAO (si existe una columna parecida la usamos; si no, creamos una nueva)
    col_pao = next((c for c in df_all.columns if re.search(r"validaci[oó]n\s*pao", c, re.I)), "Validación PAO")
    nuevo[col_pao] = vao_new
    if col_tipo: nuevo[col_tipo] = tipo_new
    if col_obs:  nuevo[col_obs]  = obs_new
    for col in cols_HN:
        nuevo[col] = valores_hn.get(col, "")

    df_all = pd.concat([df_all, pd.DataFrame([nuevo])], ignore_index=True)
    st.success("Registro agregado temporalmente. Recuerda descargar el Excel.")

# ------------------------ generar Excel (nuevo o actualizar) ------------------------
st.subheader("5) Generar Excel nuevo o actualizar uno anterior")
modo = st.radio("¿Cómo quieres generar el archivo final?", ["Empezar uno nuevo", "Actualizar un Excel anterior"], index=0)

df_final = df_all.copy()

if modo == "Actualizar un Excel anterior":
    prev = st.file_uploader("📎 Excel anterior (opcional)", type=["xlsx", "xlsm"], key="prev_x")
    if prev:
        try:
            xold = pd.ExcelFile(prev)
            frames = [pd.read_excel(xold, sheet_name=sh) for sh in xold.sheet_names]
            old_df = pd.concat(frames, ignore_index=True)
            old_df = clean_cols(old_df)

            # Estandarizar también la columna Delegación del archivo anterior (si no la trae)
            if "Delegación" not in old_df.columns:
                if old_df.shape[1] > 3:
                    old_df["Delegación"] = old_df.iloc[:, 3]
                else:
                    old_df["Delegación"] = ""

            # Quitar otras columnas "delegación" duplicadas en el anterior
            drop_like = [c for c in old_df.columns if c != "Delegación" and re.search(r"delegaci[oó]n", str(c), re.I)]
            old_df = old_df.drop(columns=drop_like)

            df_final = pd.concat([old_df, df_all], ignore_index=True)
            st.info(f"Archivo anterior detectado ({len(old_df)} filas). Se sumó al actual.")
        except Exception as e:
            st.error(f"No se pudo leer el archivo anterior: {e}")

# Duplicados exactos fuera
df_final = df_final.drop_duplicates()

with st.expander("🔎 Vista previa del Excel a generar"):
    st.dataframe(df_final, use_container_width=True, height=420)

# Exportar con hojas por trimestre si existen
sheets = {}
for t in ["I", "II", "III", "IV"]:
    parte = df_final[df_final["Trimestre"] == t]
    if not parte.empty:
        sheets[f"{t} Trimestre"] = parte
if not sheets:
    sheets = {"Datos": df_final}

export_xlsx(sheets, filename="seguimiento_trimestres_generado.xlsx")

st.caption("Ahora la delegación se toma SIEMPRE de la **columna D** y se muestra en una sola columna estándar: **Delegación**.")







