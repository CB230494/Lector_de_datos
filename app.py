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
    """Nombres de columnas H..N (Excel H..N → 0-based 7..13)."""
    start, end = 7, 14
    end = min(end, df.shape[1])
    return list(df.columns[start:end]) if start < end else []

def add_trimestre(df: pd.DataFrame, label: str) -> pd.DataFrame:
    df = df.copy()
    df["Trimestre"] = label
    return df

def standardize_delegacion_from_colD(df: pd.DataFrame) -> pd.DataFrame:
    """Crea columna estándar 'Delegación' desde columna D (índice 3) y elimina otras similares."""
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
    for c in df.columns:
        if re.fullmatch(pat, c, flags=re.I):
            return c
    return None

def ensure_row_id(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "_row_id" not in df.columns:
        df["_row_id"] = [str(uuid.uuid4()) for _ in range(len(df))]
    return df

# Sí/No
def _norm_yesno(x: str) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)): return ""
    s = str(x).strip().lower()
    if s in {"si", "sí", "s", "yes", "y"}: return "Sí"
    if s in {"no", "n"}: return "No"
    return ""

def _is_yesno_column(series: pd.Series) -> bool:
    if series.empty: return False
    vals = set(_norm_yesno(v) for v in series.dropna().unique())
    return vals.issubset({"Sí","No"}) and len(vals) <= 2

def export_xlsx_force_4_sheets(dfs_by_trim: dict, filename: str):
    """Escribe SIEMPRE hojas I/II/III/IV. Vacías → solo encabezados."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sample = next((df for df in dfs_by_trim.values() if df is not None and not df.empty), None)
        cols = list(sample.columns) if sample is not None else []
        for t, sheet_name in [("I","I Trimestre"),("II","II Trimestre"),("III","III Trimestre"),("IV","IV Trimestre")]:
            df = dfs_by_trim.get(t)
            (df if df is not None and not df.empty else pd.DataFrame(columns=cols))\
                .to_excel(writer, index=False, sheet_name=sheet_name[:31])
    st.download_button("📥 Descargar Excel", data=output.getvalue(),
                       file_name=filename,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ===================== 1) Cargar archivo base (1–4 trimestres, auto) =====================
st.subheader("1) Cargar archivo base (auto-detección 1–4 trimestres)")
archivo_base = st.file_uploader("📂 Sube el Excel (IT/IIT o I/II/III/IV)", type=["xlsx","xlsm"])
if not archivo_base:
    st.info("Sube el archivo para continuar.")
    st.stop()

xls = pd.ExcelFile(archivo_base)
sheet_names = xls.sheet_names

# Detección automática de trimestre por nombre
TRIM_MAP_PATTERNS = [
    (r"^(it|i\s*tr|1t|primer|1)\b",  "I"),
    (r"^(iit|ii\s*tr|2t|seg|segundo|2)\b", "II"),
    (r"^(iii|iii\s*tr|3t|terc|tercero|3)\b", "III"),
    (r"^(iv|iv\s*tr|4t|cuart|cuarto|4)\b",  "IV"),
]
def guess_trim(sheet_name: str) -> str:
    s = sheet_name.strip().lower()
    for pat, lab in TRIM_MAP_PATTERNS:
        if re.search(pat, s): return lab
    return ""

# Asignar por nombre, y lo que falte por orden
mapped = {}
used = set()
for sh in sheet_names:
    lab = guess_trim(sh)
    if lab and lab not in used:
        mapped[sh] = lab
        used.add(lab)
remaining_labels = [l for l in ["I","II","III","IV"] if l not in used]
for sh in sheet_names:
    if sh not in mapped and remaining_labels:
        mapped[sh] = remaining_labels.pop(0)
# Si hay más de 4 hojas, ignoramos las sobrantes
mapped = {sh: mapped[sh] for sh in mapped if mapped[sh] in {"I","II","III","IV"}}

frames = []
for sh, tri in mapped.items():
    df_sh = pd.read_excel(xls, sheet_name=sh)
    df_sh = clean_cols(df_sh)
    df_sh = standardize_delegacion_from_colD(df_sh)
    df_sh = add_trimestre(df_sh, tri)
    frames.append(df_sh)

if not frames:
    st.error("No pude detectar trimestres. Renombra las hojas (ej.: IT, IIT, I, II, III, IV) o numéralas.")
    st.stop()

# Consolidado + ID
df_all = pd.concat(frames, ignore_index=True)
df_all = ensure_row_id(df_all)

# Detectar columnas H..N (para inputs)
cols_HN = []
for df_sample in frames:
    if len(take_cols_H_to_N(df_sample)) > len(cols_HN):
        cols_HN = take_cols_H_to_N(df_sample)

# Detectar Tipo y Observaciones (si existen en alguna hoja)
def find_in_frames(frames, pat):
    for d in frames:
        c = find_col_by_exact(d, pat)
        if c: return c
    return None
col_tipo = find_in_frames(frames, r"tipo\s*de\s*actividad\.?")
col_obs  = find_in_frames(frames, r"observaciones?\.?")

# Asegurar Fecha e Instituciones
if "Fecha" not in df_all.columns: df_all["Fecha"] = pd.NaT
if "Instituciones" not in df_all.columns: df_all["Instituciones"] = ""

# PAO + columnas Sí/No (incluye H–N)
col_pao = next((c for c in df_all.columns if re.search(r"validaci[oó]n\s*pao", c, re.I)), "Validación PAO")
if col_pao not in df_all.columns: df_all[col_pao] = ""

yesno_cols = [col_pao]
for c in df_all.columns:
    if c in {"Delegación","Trimestre","_row_id","Fecha","Instituciones"} or c in yesno_cols:
        continue
    if df_all[c].dtype == "O" and _is_yesno_column(df_all[c]):
        yesno_cols.append(c)
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
cols_base = ["Fecha","Delegación","Trimestre"] + [c for c in [col_tipo, col_obs, "Instituciones"] if c]
cols_mostrar = cols_base + [c for c in cols_HN if c not in cols_base] + [col_pao]
for c in cols_mostrar:
    if c not in df_all.columns:
        df_all[c] = "" if c != "Fecha" else pd.NaT
cols_editor = [c for c in cols_mostrar if c in df_all.columns] + ["_row_id"]

# ===================== 3) Editor =====================
st.subheader("3) Editor por delegación (editar, agregar filas/columnas, eliminar)")

df_ed = df_filtrado[cols_editor].copy()
df_ed["Eliminar"] = False

col_config = {
    "_row_id": st.column_config.TextColumn("ID (interno)", disabled=True),
    "Eliminar": st.column_config.CheckboxColumn("Eliminar"),
    "Fecha": st.column_config.DateColumn("Fecha"),
}
for c in yesno_cols:
    if c in df_ed.columns:
        col_config[c] = st.column_config.SelectboxColumn(c, options=["", "Sí", "No"])

edited = st.data_editor(
    df_ed,
    num_rows="dynamic",
    use_container_width=True,
    height=420,
    column_config=col_config,
    hide_index=True,
    key="editor",
)

c1, c2, c3, c4, c5, c6 = st.columns(6)
with c1: do_add_iii = st.button("➕ Fila base a III", use_container_width=True)
with c2: do_add_iv  = st.button("➕ Fila base a IV", use_container_width=True)
with c3: delete_now = st.button("🗑️ Eliminar seleccionados", use_container_width=True)
with c4: save_now   = st.button("💾 Guardar cambios", use_container_width=True)
with c5: new_col    = st.text_input("Nueva columna", placeholder="Nombre de columna…")
with c6: add_col    = st.button("➕ Agregar columna", use_container_width=True)

PROTECTED = {"_row_id","Delegación","Trimestre"}
if add_col and new_col:
    if new_col in df_all.columns:
        st.warning("Ya existe esa columna.")
    elif new_col in PROTECTED:
        st.warning("Nombre reservado.")
    else:
        df_all[new_col] = ""
        st.success(f"Columna '{new_col}' agregada.")

def blank_row(trim_label: str):
    base = {k: "" for k in cols_mostrar}
    base["Fecha"] = pd.NaT
    base["Delegación"] = (deleg_sel if deleg_sel != "(Todas)" else "")
    base["Trimestre"]  = trim_label
    for c in yesno_cols:
        if c in base: base[c] = ""
    base["_row_id"] = str(uuid.uuid4())
    return base

if do_add_iii:
    df_all = pd.concat([df_all, pd.DataFrame([blank_row("III")])], ignore_index=True)
    st.success("Fila base creada en III.")
if do_add_iv:
    df_all = pd.concat([df_all, pd.DataFrame([blank_row("IV")])], ignore_index=True)
    st.success("Fila base creada en IV.")

if delete_now:
    ids = set(edited.loc[edited["Eliminar"] == True, "_row_id"].astype(str).tolist())
    if ids:
        df_all = df_all[~df_all["_row_id"].astype(str).isin(ids)]
        st.success(f"Eliminadas {len(ids)} fila(s).")
    else:
        st.info("Marca 'Eliminar' en al menos una fila.")

if save_now:
    edited_clean = edited.drop(columns=["Eliminar"]).copy()
    for _, row in edited_clean.iterrows():
        rid = str(row["_row_id"]).strip() if pd.notna(row["_row_id"]) else ""
        if not rid:
            rid = str(uuid.uuid4()); row["_row_id"] = rid
            new_entry = {c: row.get(c, "") for c in cols_mostrar + ["_row_id"]}
            df_all = pd.concat([df_all, pd.DataFrame([new_entry])], ignore_index=True)
        else:
            mask = df_all["_row_id"].astype(str).eq(rid)
            for c in cols_mostrar:
                if c in edited_clean.columns:
                    df_all.loc[mask, c] = row.get(c, "")
    # re-normalizar Sí/No
    for c in yesno_cols:
        if c in df_all.columns:
            df_all[c] = df_all[c].map(_norm_yesno)
    st.success("Cambios guardados.")

# ===================== 4) Formulario =====================
st.subheader("4) Formulario rápido para agregar filas")
with st.form("form_add"):
    a, b, c, d = st.columns(4)
    fecha_new = a.date_input("Fecha", value=date.today())
    trim_new  = b.selectbox("Trimestre", ["I","II","III","IV"], index=2)
    deleg_new = c.selectbox("Delegación", sorted([deleg_sel] + delegaciones) if delegaciones else [""])
    pao_new   = d.selectbox("Validación PAO", ["", "Sí", "No"], index=0)

    tipo_new = ""
    if col_tipo:
        tipos_cat = ["Rendición de cuentas","Seguimiento","Líneas de acción","Informe territorial"]
        tipo_new = st.multiselect("Tipo de actividad (multi)", tipos_cat, default=[])
        tipo_new = "; ".join(tipo_new) if tipo_new else ""

    obs_new  = st.text_area(col_obs or "Observaciones", height=100)
    inst_new = st.text_input("Instituciones", "", placeholder="Ingrese instituciones involucradas…")

    st.markdown("**Completar columnas H–N**")
    valores_hn = {}
    for col in cols_HN:
        if col in yesno_cols:
            valores_hn[col] = st.selectbox(col, ["", "Sí", "No"], index=0)
        else:
            valores_hn[col] = st.text_input(col, value="")

    enviar = st.form_submit_button("➕ Agregar registro")

if enviar:
    nuevo = {
        "Fecha": pd.to_datetime(fecha_new),
        "Delegación": deleg_new,
        "Trimestre": trim_new,
        "Instituciones": inst_new,
        "_row_id": str(uuid.uuid4()),
    }
    nuevo[col_pao if col_pao in df_all.columns else "Validación PAO"] = pao_new
    if col_tipo: nuevo[col_tipo] = tipo_new
    if col_obs:  nuevo[col_obs]  = obs_new
    for col in cols_HN:
        if col not in df_all.columns: df_all[col] = ""
        nuevo[col] = valores_hn.get(col, "")
    df_all = pd.concat([df_all, pd.DataFrame([nuevo])], ignore_index=True)
    st.success("Registro agregado.")

# ===================== 5) Vista por tabs =====================
st.subheader("📑 Vista por 'hojas' (I/II/III/IV)")
t1, t2, t3, t4 = st.tabs(["I Trimestre","II Trimestre","III Trimestre","IV Trimestre"])
with t1: st.dataframe(df_all[df_all["Trimestre"]=="I"],  use_container_width=True, height=300)
with t2: st.dataframe(df_all[df_all["Trimestre"]=="II"], use_container_width=True, height=300)
with t3: st.dataframe(df_all[df_all["Trimestre"]=="III"],use_container_width=True, height=300)
with t4: st.dataframe(df_all[df_all["Trimestre"]=="IV"], use_container_width=True, height=300)

# ===================== 6) Exportación (siempre 4 hojas) =====================
st.subheader("6) Descargar Excel (siempre con 4 hojas)")
export_cols = [c for c in df_all.columns if c != "_row_id"]
df_export  = df_all[export_cols].drop_duplicates()

dfs_by_trim = {
    "I":   df_export[df_export["Trimestre"]=="I"],
    "II":  df_export[df_export["Trimestre"]=="II"],
    "III": df_export[df_export["Trimestre"]=="III"],
    "IV":  df_export[df_export["Trimestre"]=="IV"],
}
export_xlsx_force_4_sheets(dfs_by_trim, filename="seguimiento_trimestres_generado.xlsx")

st.caption("Auto-detección de hojas → trimestres; Delegación siempre desde columna D; editor con agregar/editar/eliminar filas y columnas; formulario con Fecha e Instituciones; exportación fija con 4 hojas.")



