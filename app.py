# app.py
import streamlit as st
import pandas as pd
import io
import re
import uuid

st.set_page_config(page_title="Seguimiento por Trimestre — Editor y Generador", layout="wide")
st.title("📘 Seguimiento por Trimestre — Lector + Editor + Formulario")

# ===================== Helpers =====================
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
    """Crea columna estándar 'Delegación' SIEMPRE desde la columna D (índice 3) y elimina otras 'delegación*'."""
    df = df.copy()
    if df.shape[1] > 3:
        df["Delegación"] = df.iloc[:, 3]
    else:
        df["Delegación"] = ""
    drop_like = [c for c in df.columns if c != "Delegación" and re.search(r"delegaci[oó]n", str(c), re.I)]
    if drop_like:
        df = df.drop(columns=drop_like)
    return df

def find_col_by_exact(df, pat):
    """Busca columna por nombre exacto (regex) insensible a mayúsculas."""
    for c in df.columns:
        if re.fullmatch(pat, c, flags=re.I):
            return c
    return None

def ensure_row_id(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega un ID estable por fila para poder reconciliar ediciones."""
    df = df.copy()
    if "_row_id" not in df.columns:
        df["_row_id"] = [str(uuid.uuid4()) for _ in range(len(df))]
    return df

# ---- Sí/No detection/normalization
def _norm_yesno(x: str) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).strip().lower()
    if s in {"si", "sí", "s", "yes", "y"}: return "Sí"
    if s in {"no", "n"}: return "No"
    return ""

def _is_yesno_column(series: pd.Series) -> bool:
    """Detecta si una columna parece ser 'Sí/No' (admite vacíos)."""
    if series.empty:
        return False
    vals = set(_norm_yesno(v) for v in series.dropna().unique())
    return vals.issubset({"Sí", "No"}) and len(vals) <= 2

def export_xlsx_force_4_sheets(dfs_by_trim: dict, filename: str):
    """Escribe SIEMPRE las 4 hojas I/II/III/IV. Si un trimestre está vacío, crea la hoja con encabezados."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Encabezados estándar (de la primera tabla no vacía)
        sample = next((df for df in dfs_by_trim.values() if df is not None and not df.empty), None)
        cols = list(sample.columns) if sample is not None else []

        for t, sheet_name in [("I","I Trimestre"),("II","II Trimestre"),("III","III Trimestre"),("IV","IV Trimestre")]:
            df = dfs_by_trim.get(t)
            if df is None or df.empty:
                pd.DataFrame(columns=cols).to_excel(writer, index=False, sheet_name=sheet_name[:31])
            else:
                df.to_excel(writer, index=False, sheet_name=sheet_name[:31])

    st.download_button(
        "📥 Descargar Excel",
        data=output.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ===================== 1) Cargar archivo base (IT y IIT) =====================
st.subheader("1) Cargar archivo base (IT y IIT)")
archivo_base = st.file_uploader("📂 Sube el Excel (contiene IT e IIT)", type=["xlsx", "xlsm"])
if not archivo_base:
    st.info("Sube el archivo para continuar.")
    st.stop()

xls = pd.ExcelFile(archivo_base)
sheet_names = xls.sheet_names

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

# Leer
df_it  = pd.read_excel(xls, sheet_name=sheet_it)
df_iit = pd.read_excel(xls, sheet_name=sheet_iit)
df_it, df_iit = clean_cols(df_it), clean_cols(df_iit)

# Delegación desde D
df_it  = standardize_delegacion_from_colD(df_it)
df_iit = standardize_delegacion_from_colD(df_iit)

# Trimestre
df_it  = add_trimestre(df_it, "I")
df_iit = add_trimestre(df_iit, "II")

# H..N por posición (nombres reales)
cols_HN_it  = take_cols_H_to_N(df_it)
cols_HN_iit = take_cols_H_to_N(df_iit)
cols_HN = cols_HN_it if len(cols_HN_it) >= len(cols_HN_iit) else cols_HN_iit

# Detectar Tipo y Observaciones por nombre si existen
col_tipo = find_col_by_exact(df_it, r"tipo\s*de\s*actividad\.?") or find_col_by_exact(df_iit, r"tipo\s*de\s*actividad\.?")
col_obs  = find_col_by_exact(df_it, r"observaciones?\.?") or find_col_by_exact(df_iit, r"observaciones?\.?")

# Consolidado + ID
df_all = pd.concat([df_it, df_iit], ignore_index=True)
df_all = ensure_row_id(df_all)

# Detectar/normalizar PAO y columnas Sí/No (incluye H–N)
col_pao = next((c for c in df_all.columns if re.search(r"validaci[oó]n\s*pao", c, re.I)), "Validación PAO")
if col_pao not in df_all.columns:
    df_all[col_pao] = ""

yesno_cols = [col_pao]
for c in df_all.columns:
    if c in {"Delegación", "Trimestre", "_row_id"} or c in yesno_cols:
        continue
    if df_all[c].dtype == "O" and _is_yesno_column(df_all[c]):
        yesno_cols.append(c)
# Normalizar Sí/No
for c in yesno_cols:
    df_all[c] = df_all[c].map(_norm_yesno)

# ===================== 2) Filtros =====================
st.subheader("2) Filtros")
delegaciones = sorted([d for d in df_all["Delegación"].dropna().astype(str).map(str.strip).unique() if d])
deleg_sel = st.selectbox("🏢 Delegación (columna D)", options=["(Todas)"] + delegaciones, index=0)
trims_sel = st.multiselect("🗓️ Trimestres", options=["I","II","III","IV"], default=["I","II"])

df_filtrado = df_all.copy()
if deleg_sel != "(Todas)":
    df_filtrado = df_filtrado[df_filtrado["Delegación"] == deleg_sel]
if trims_sel:
    df_filtrado = df_filtrado[df_filtrado["Trimestre"].isin(trims_sel)]

# Columnas visibles/editar
cols_base = ["Delegación", "Trimestre"] + [c for c in [col_tipo, col_obs] if c]
cols_mostrar = cols_base + [c for c in cols_HN if c not in cols_base] + [col_pao]
cols_editor = cols_mostrar + ["_row_id"]

# ===================== 3) Editor =====================
st.subheader("3) Editor por delegación (editar, agregar o eliminar filas)")

df_ed = df_filtrado[cols_editor].copy()
df_ed["Eliminar"] = False

# Config: Select Sí/No para columnas binarias
col_config = {
    "_row_id": st.column_config.TextColumn("ID (interno)", disabled=True),
    "Eliminar": st.column_config.CheckboxColumn("Eliminar"),
}
for c in yesno_cols:
    if c in df_ed.columns:
        col_config[c] = st.column_config.SelectboxColumn(c, options=["", "Sí", "No"], required=False)

edited = st.data_editor(
    df_ed,
    num_rows="dynamic",                   # permite agregar nuevas filas
    use_container_width=True,
    height=420,
    column_config=col_config,
    hide_index=True,
    key="editor",
)

colE1, colE2, colE3, colE4 = st.columns(4)
with colE1:
    do_quick_add_iii = st.button("➕ Fila base a III", use_container_width=True)
with colE2:
    do_quick_add_iv  = st.button("➕ Fila base a IV", use_container_width=True)
with colE3:
    delete_now       = st.button("🗑️ Eliminar seleccionados", use_container_width=True)
with colE4:
    apply_changes    = st.button("💾 Guardar cambios", use_container_width=True)

def blank_row_for_trim(trim_label: str):
    base = {k: "" for k in cols_mostrar}
    base["Delegación"] = (deleg_sel if deleg_sel != "(Todas)" else "")
    base["Trimestre"]  = trim_label
    for c in yesno_cols:
        if c in base: base[c] = ""
    base["_row_id"]    = str(uuid.uuid4())
    return base

if do_quick_add_iii:
    df_all = pd.concat([df_all, pd.DataFrame([blank_row_for_trim("III")])], ignore_index=True)
    st.success("Fila base creada en III.")

if do_quick_add_iv:
    df_all = pd.concat([df_all, pd.DataFrame([blank_row_for_trim("IV")])], ignore_index=True)
    st.success("Fila base creada en IV.")

# Eliminar seleccionados
if delete_now:
    to_delete_ids = set(edited.loc[edited["Eliminar"] == True, "_row_id"].astype(str).tolist())
    if to_delete_ids:
        df_all = df_all[~df_all["_row_id"].astype(str).isin(to_delete_ids)]
        st.success(f"Eliminadas {len(to_delete_ids)} fila(s).")
    else:
        st.info("Marca 'Eliminar' en al menos una fila.")

# Guardar cambios del editor en el consolidado
if apply_changes:
    edited_clean = edited.drop(columns=["Eliminar"]).copy()

    for _, row in edited_clean.iterrows():
        rid = str(row["_row_id"]).strip() if pd.notna(row["_row_id"]) else ""
        if not rid:
            # nueva fila
            rid = str(uuid.uuid4())
            row["_row_id"] = rid
            new_entry = {c: row.get(c, "") for c in cols_mostrar + ["_row_id"]}
            df_all = pd.concat([df_all, pd.DataFrame([new_entry])], ignore_index=True)
        else:
            # actualizar fila existente
            mask = df_all["_row_id"].astype(str).eq(rid)
            for c in cols_mostrar:
                if c in edited_clean.columns:
                    df_all.loc[mask, c] = row.get(c, "")
    # normalizar Sí/No otra vez por si se editaron
    for c in yesno_cols:
        if c in df_all.columns:
            df_all[c] = df_all[c].map(_norm_yesno)

    st.success("Cambios guardados.")

# ===================== 4) Formulario rápido =====================
st.subheader("4) Formulario rápido para agregar filas")
with st.form("form_add_quick"):
    c1, c2, c3 = st.columns(3)
    trim_new   = c1.selectbox("Trimestre", ["I", "II", "III", "IV"], index=2)
    deleg_new  = c3.selectbox("Delegación", sorted([deleg_sel] + delegaciones) if delegaciones else [""])

    # PAO (Sí/No)
    pao_new    = c2.selectbox("Validación PAO", ["", "Sí", "No"], index=0)

    tipo_new = ""
    if col_tipo:
        tipos_catalogo = ["Rendición de cuentas", "Seguimiento", "Líneas de acción", "Informe territorial"]
        tipo_new = st.multiselect("Tipo de actividad (multi)", tipos_catalogo, default=[])
        tipo_new = "; ".join(tipo_new) if tipo_new else ""

    obs_new = st.text_area(col_obs or "Observaciones", height=100)

    st.markdown("**Completar columnas**")
    valores_hn = {}
    for col in cols_HN:
        if col in yesno_cols:
            valores_hn[col] = st.selectbox(col, ["", "Sí", "No"], index=0)
        else:
            valores_hn[col] = st.text_input(col, value="")

    enviado = st.form_submit_button("➕ Agregar registro")

if enviado:
    nuevo = {"Delegación": deleg_new, "Trimestre": trim_new, "_row_id": str(uuid.uuid4())}
    nuevo["Validación PAO" if "Validación PAO" in df_all.columns else next((c for c in yesno_cols if re.search(r"validaci[oó]n\s*pao", c, re.I)), "Validación PAO")] = pao_new
    if col_tipo: nuevo[col_tipo] = tipo_new
    if col_obs:  nuevo[col_obs]  = obs_new
    for col in cols_HN:
        nuevo[col] = valores_hn.get(col, "")
    df_all = pd.concat([df_all, pd.DataFrame([nuevo])], ignore_index=True)
    st.success("Registro agregado.")

# ===================== 5) Vista por 'hojas' (tabs) =====================
st.subheader("📑 Vista por 'hojas' (I/II/III/IV)")
t1, t2, t3, t4 = st.tabs(["I Trimestre", "II Trimestre", "III Trimestre", "IV Trimestre"])
with t1:
    st.dataframe(df_all[df_all["Trimestre"]=="I"], use_container_width=True, height=300)
with t2:
    st.dataframe(df_all[df_all["Trimestre"]=="II"], use_container_width=True, height=300)
with t3:
    st.dataframe(df_all[df_all["Trimestre"]=="III"], use_container_width=True, height=300)
with t4:
    st.dataframe(df_all[df_all["Trimestre"]=="IV"], use_container_width=True, height=300)

# ===================== 6) Exportación (siempre 4 hojas) =====================
st.subheader("6) Descargar Excel (siempre con 4 hojas)")
# Quitar duplicados exactos ignorando _row_id
export_cols = [c for c in df_all.columns if c != "_row_id"]
df_export = df_all[export_cols].drop_duplicates()

dfs_by_trim = {
    "I":   df_export[df_export["Trimestre"]=="I"],
    "II":  df_export[df_export["Trimestre"]=="II"],
    "III": df_export[df_export["Trimestre"]=="III"],
    "IV":  df_export[df_export["Trimestre"]=="IV"],
}
export_xlsx_force_4_sheets(dfs_by_trim, filename="seguimiento_trimestres_generado.xlsx")

st.caption("Delegación se toma SIEMPRE de la columna D. Editor con Sí/No automáticos en PAO y H–N. Exportación fija con 4 hojas (I/II/III/IV).")











