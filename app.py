# app.py
import streamlit as st
import pandas as pd
import io
import re
import uuid
from datetime import date

st.set_page_config(page_title="Seguimiento por Trimestre — Editor y Generador", layout="wide")
st.title("📘 Seguimiento por Trimestre — Lector + Editor + Formulario (Delegación = Columna D)")

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

# ===================== 1) Cargar archivo base (1–4 trimestres) =====================
st.subheader("1) Cargar archivo base (admite 1–4 trimestres)")
archivo_base = st.file_uploader("📂 Sube el Excel (puede contener IT/IIT o I/II/III/IV)", type=["xlsx", "xlsm"])
if not archivo_base:
    st.info("Sube el archivo para continuar.")
    st.stop()

xls = pd.ExcelFile(archivo_base)
sheet_names = xls.sheet_names

# Detección automática de trimestre por nombre de hoja + selector de override
TRIM_MAP_PATTERNS = [
    (r"^(it|i\s*tr|1t|primer|1)\b", "I"),
    (r"^(iit|ii\s*tr|2t|seg|segundo|2)\b", "II"),
    (r"^(iii|iii\s*tr|3t|terc|tercero|3)\b", "III"),
    (r"^(iv|iv\s*tr|4t|cuart|cuarto|4)\b", "IV"),
]

def guess_trim(sheet_name: str) -> str:
    s = sheet_name.strip().lower()
    for pat, label in TRIM_MAP_PATTERNS:
        if re.search(pat, s):
            return label
    return ""  # sin detectar

st.write("### Mapear hojas a Trimestre")
sheet_to_trim = {}
for sh in sheet_names:
    g = guess_trim(sh)
    sheet_to_trim[sh] = st.selectbox(
        f"Hoja: **{sh}** → Trimestre",
        options=["", "I", "II", "III", "IV"],
        index=(["", "I", "II", "III", "IV"].index(g) if g in {"I","II","III","IV"} else 0),
        key=f"map_{sh}"
    )

# Leer y construir consolidado con las hojas mapeadas
frames = []
for sh, tri in sheet_to_trim.items():
    if tri == "":
        continue
    df_sh = pd.read_excel(xls, sheet_name=sh)
    df_sh = clean_cols(df_sh)
    df_sh = standardize_delegacion_from_colD(df_sh)
    df_sh = add_trimestre(df_sh, tri)
    frames.append(df_sh)

if not frames:
    st.error("No mapeaste ninguna hoja a un trimestre. Selecciona al menos una.")
    st.stop()

# Consolidado + ID
df_all = pd.concat(frames, ignore_index=True)
df_all = ensure_row_id(df_all)

# Detectar columnas H..N (para inputs rápidos)
cols_HN = []
for df_sample in frames:
    cols_HN = max(cols_HN, take_cols_H_to_N(df_sample), key=lambda l: len(l))  # la más larga

# Detectar Tipo y Observaciones por nombre si existen (en cualquiera de las hojas)
def find_in_frames(frames, pat):
    for d in frames:
        c = find_col_by_exact(d, pat)
        if c: return c
    return None

col_tipo = find_in_frames(frames, r"tipo\s*de\s*actividad\.?")
col_obs  = find_in_frames(frames, r"observaciones?\.?")

# Asegurar columnas nuevas que pediste: Fecha e Instituciones (si no existen)
if "Fecha" not in df_all.columns:
    df_all["Fecha"] = pd.NaT
if "Instituciones" not in df_all.columns:
    df_all["Instituciones"] = ""

# Detectar/normalizar PAO y columnas Sí/No (incluye H–N)
col_pao = next((c for c in df_all.columns if re.search(r"validaci[oó]n\s*pao", c, re.I)), "Validación PAO")
if col_pao not in df_all.columns:
    df_all[col_pao] = ""

yesno_cols = [col_pao]
for c in df_all.columns:
    if c in {"Delegación", "Trimestre", "_row_id", "Fecha", "Instituciones"} or c in yesno_cols:
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
trims_sel = st.multiselect("🗓️ Trimestres", options=["I","II","III","IV"], default=["I","II","III","IV"])

df_filtrado = df_all.copy()
if deleg_sel != "(Todas)":
    df_filtrado = df_filtrado[df_filtrado["Delegación"] == deleg_sel]
if trims_sel:
    df_filtrado = df_filtrado[df_filtrado["Trimestre"].isin(trims_sel)]

# Columnas visibles/editar
cols_base = ["Fecha", "Delegación", "Trimestre"] + [c for c in [col_tipo, col_obs, "Instituciones"] if c]
cols_mostrar = cols_base + [c for c in cols_HN if c not in cols_base] + [col_pao]
# Asegurar que existen
for c in cols_mostrar:
    if c not in df_all.columns:
        df_all[c] = "" if c not in {"Fecha"} else pd.NaT
cols_editor = [c for c in cols_mostrar if c in df_all.columns] + ["_row_id"]

# ===================== 3) Editor =====================
st.subheader("3) Editor por delegación (editar, **agregar filas y columnas**, eliminar)")

df_ed = df_filtrado[cols_editor].copy()
df_ed["Eliminar"] = False

# Config: Select Sí/No para columnas binarias
col_config = {
    "_row_id": st.column_config.TextColumn("ID (interno)", disabled=True),
    "Eliminar": st.column_config.CheckboxColumn("Eliminar"),
    "Fecha": st.column_config.DateColumn("Fecha"),
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

# Controles extra (agregar/eliminar columnas + filas base III/IV)
colE1, colE2, colE3, colE4, colE5, colE6 = st.columns(6)
with colE1:
    do_quick_add_iii = st.button("➕ Fila base a III", use_container_width=True)
with colE2:
    do_quick_add_iv  = st.button("➕ Fila base a IV", use_container_width=True)
with colE3:
    delete_now       = st.button("🗑️ Eliminar seleccionados", use_container_width=True)
with colE4:
    apply_changes    = st.button("💾 Guardar cambios", use_container_width=True)
with colE5:
    new_col_name     = st.text_input("Nueva columna", placeholder="Nombre de columna…")
with colE6:
    add_col_now      = st.button("➕ Agregar columna", use_container_width=True)

# Agregar columna global
PROTECTED_COLS = {"_row_id", "Delegación", "Trimestre"}
if add_col_now and new_col_name:
    if new_col_name in df_all.columns:
        st.warning("Ya existe una columna con ese nombre.")
    elif new_col_name in PROTECTED_COLS:
        st.warning("Ese nombre está reservado.")
    else:
        df_all[new_col_name] = ""
        st.success(f"Columna '{new_col_name}' agregada.")
        # Si la columna es binaria de Sí/No, se normalizará al guardar

def blank_row_for_trim(trim_label: str):
    base = {k: "" for k in cols_mostrar}
    base["Fecha"] = pd.NaT
    base["Delegación"] = (deleg_sel if deleg_sel != "(Todas)" else "")
    base["Trimestre"]  = trim_label
    for c in yesno_cols:
        if c in base: base[c] = ""
    base["Instituciones"] = base.get("Instituciones", "")
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

    # Reconciliar filas nuevas/existentes
    existing_ids = set(df_all["_row_id"].astype(str))
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

    # Re‑detectar nuevas columnas Sí/No por si agregaste alguna
    yesno_cols = [col_pao]
    for c in df_all.columns:
        if c in {"Delegación", "Trimestre", "_row_id", "Fecha", "Instituciones"} or c in yesno_cols:
            continue
        if df_all[c].dtype == "O" and _is_yesno_column(df_all[c]):
            yesno_cols.append(c)
    # Normalizar Sí/No
    for c in yesno_cols:
        if c in df_all.columns:
            df_all[c] = df_all[c].map(_norm_yesno)

    st.success("Cambios guardados.")

# ===================== 4) Formulario rápido =====================
st.subheader("4) Formulario rápido para agregar filas")
with st.form("form_add_quick"):
    c0, c1, c2, c3 = st.columns(4)
    fecha_new  = c0.date_input("Fecha", value=date.today())
    trim_new   = c1.selectbox("Trimestre", ["I", "II", "III", "IV"], index=2)
    deleg_new  = c2.selectbox("Delegación", sorted([deleg_sel] + delegaciones) if delegaciones else [""])
    # PAO (Sí/No)
    pao_new    = c3.selectbox("Validación PAO", ["", "Sí", "No"], index=0)

    tipo_new = ""
    if col_tipo:
        tipos_catalogo = ["Rendición de cuentas", "Seguimiento", "Líneas de acción", "Informe territorial"]
        tipo_new = st.multiselect("Tipo de actividad (multi)", tipos_catalogo, default=[])
        tipo_new = "; ".join(tipo_new) if tipo_new else ""

    obs_new = st.text_area(col_obs or "Observaciones", height=100)
    inst_new = st.text_input("Instituciones", value="", placeholder="Ingrese instituciones involucradas…")

    st.markdown("**Completar columnas H–N**")
    valores_hn = {}
    for col in cols_HN:
        # si la columna fue detectada como Sí/No → select
        if col in yesno_cols:
            valores_hn[col] = st.selectbox(col, ["", "Sí", "No"], index=0)
        else:
            valores_hn[col] = st.text_input(col, value="")

    enviado = st.form_submit_button("➕ Agregar registro")

if enviado:
    nuevo = {
        "Fecha": pd.to_datetime(fecha_new),
        "Delegación": deleg_new,
        "Trimestre": trim_new,
        "_row_id": str(uuid.uuid4()),
        "Instituciones": inst_new,
    }
    # PAO
    nuevo_col_pao = col_pao if col_pao in df_all.columns else "Validación PAO"
    nuevo[nuevo_col_pao] = pao_new
    if col_tipo: nuevo[col_tipo] = tipo_new
    if col_obs:  nuevo[col_obs]  = obs_new
    for col in cols_HN:
        if col not in df_all.columns:
            df_all[col] = ""  # crear si no existe aún
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

st.caption("Delegación se toma SIEMPRE de la columna D. Agrega Fecha e Instituciones en el formulario. Editor permite agregar filas y columnas, editar y eliminar. Exportación fija con 4 hojas.")




