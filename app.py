import streamlit as st
import pandas as pd
import io
import re
import uuid
import unicodedata
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

# Columnas canónicas de Sí/No para usar en editor y formulario
COL_SEGUIMIENTO = "Seguimiento líneas de acción"
COL_ACUERDOS    = "¿Hubo acuerdos inter-institucionales concretos en esta sesión?"

# ===================== 1) Cargar archivo base (1–4 trimestres, auto) =====================
st.subheader("1) Cargar archivo base (auto-detección 1–4 trimestres)")
archivo_base = st.file_uploader("📂 Sube el Excel (IT/IIT o I/II/III/IV)", type=["xlsx","xlsm"])
if not archivo_base:
    st.info("Sube el archivo para continuar.")
    st.stop()

# ---------- Persistencia en sesión ----------
file_key = f"{archivo_base.name}-{getattr(archivo_base, 'size', None)}"
if "file_key" not in st.session_state or st.session_state["file_key"] != file_key:
    xls = pd.ExcelFile(archivo_base)
    sheet_names = xls.sheet_names

    # ===== Detección flexible de nombres de hoja =====
    def _norm(s: str) -> str:
        s = s.strip().lower()
        s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
        s = re.sub(r'\s+', ' ', s)
        return s

    PAT_I = [
        r'^i($|\b)', r'^i\s*tri', r'^it\b',
        r'^1($|\b)', r'^1\s*tri', r'^t1($|\b)', r'^q1($|\b)',
        r'^1er\b', r'^primer\b', r'^primero\b',
        r'^trimestre\s*i\b', r'^trim\s*i\b',
    ]
    PAT_II = [
        r'^ii($|\b)', r'^ii\s*tri', r'^iit\b',
        r'^2($|\b)', r'^2\s*tri', r'^t2($|\b)', r'^q2($|\b)',
        r'^2do\b', r'^segundo?\b',
        r'^trimestre\s*ii\b', r'^trim\s*ii\b',
    ]
    PAT_III = [
        r'^iii($|\b)', r'^iii\s*tri',
        r'^3($|\b)', r'^3\s*tri', r'^t3($|\b)', r'^q3($|\b)',
        r'^3er\b', r'^tercer(o)?\b',
        r'^trimestre\s*iii\b', r'^trim\s*iii\b',
    ]
    PAT_IV = [
        r'^iv($|\b)', r'^iv\s*tri',
        r'^4($|\b)', r'^4\s*tri', r'^t4($|\b)', r'^q4($|\b)',
        r'^4to\b', r'^cuarto?\b',
        r'^trimestre\s*iv\b', r'^trim\s*iv\b',
    ]

    def _match_any(s: str, pats: list[str]) -> bool:
        return any(re.search(p, s, re.I) for p in pats)

    def guess_trim(sheet_name: str) -> str:
        s = _norm(sheet_name)
        if _match_any(s, PAT_I):   return "I"
        if _match_any(s, PAT_II):  return "II"
        if _match_any(s, PAT_III): return "III"
        if _match_any(s, PAT_IV):  return "IV"
        return ""

    # Solo mapeamos hojas que COINCIDEN; NO se rellenan por orden
    mapped, used = {}, set()
    for sh in sheet_names:
        lab = guess_trim(sh)
        if lab and lab not in used:
            mapped[sh] = lab
            used.add(lab)

    with st.expander("🔎 Ver mapeo de hojas detectado"):
        if not mapped:
            st.warning("No se detectaron hojas de trimestres.")
        else:
            st.write({sh: mapped[sh] for sh in mapped})

    # ===== Leer hojas detectadas =====
    frames = []
    for sh, tri in mapped.items():
        df_sh = pd.read_excel(xls, sheet_name=sh)
        df_sh = clean_cols(df_sh)
        df_sh = standardize_delegacion_from_colD(df_sh)
        df_sh = add_trimestre(df_sh, tri)
        frames.append(df_sh)

    if not frames:
        st.error("No pude detectar hojas IT/IIT/I/II/III/IV. Renombra las hojas o verifica el archivo.")
        st.stop()

    df_all = pd.concat(frames, ignore_index=True)
    df_all = ensure_row_id(df_all)

    # Detectar columnas H..N
    cols_HN = []
    for df_sample in frames:
        cand = take_cols_H_to_N(df_sample)
        if len(cand) > len(cols_HN): cols_HN = cand

    # Detectar Tipo y Observaciones
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

    # ===== Forzar columnas Sí/No =====
    col_pao = next((c for c in df_all.columns if re.search(r"validaci[oó]n\s*pao", c, re.I)), "Validación PAO")
    if col_pao not in df_all.columns: df_all[col_pao] = ""
    yesno_cols = {col_pao, COL_SEGUIMIENTO, COL_ACUERDOS}
    YESNO_NAME_HINTS = [
        r"validaci[oó]n\s*pao",
        r"^seguimiento\s+líneas\s+de\s+acci[oó]n$",
        r"^¿\s*hubo\s+acuerdos\s+inter[- ]?institucionales.*",
    ]
    for c in df_all.columns:
        if c in {"Delegación","Trimestre","_row_id","Fecha","Instituciones"}:
            continue
        if any(re.search(pat, c, re.I) for pat in YESNO_NAME_HINTS) or \
           (df_all[c].dtype == "O" and _is_yesno_column(df_all[c])):
            yesno_cols.add(c)

    for c in yesno_cols:
        if c not in df_all.columns: df_all[c] = ""
        df_all[c] = df_all[c].map(_norm_yesno)

    # Guardar en sesión
    st.session_state.update({
        "file_key": file_key,
        "df_all": df_all,
        "cols_HN": cols_HN,
        "col_tipo": col_tipo,
        "col_obs": col_obs,
        "col_pao": col_pao,
        "yesno_cols": list(yesno_cols),
    })

# Usar lo de sesión
df_all     = st.session_state["df_all"].copy()
cols_HN    = st.session_state["cols_HN"]
col_tipo   = st.session_state["col_tipo"]
col_obs    = st.session_state["col_obs"]
col_pao    = st.session_state["col_pao"]
yesno_cols = set(st.session_state["yesno_cols"])

# ===================== 2) Filtros =====================
st.subheader("2) Filtros")
delegaciones = sorted([d for d in df_all["Delegación"].dropna().astype(str).map(str.strip).unique() if d])
deleg_sel = st.selectbox("🏢 Delegación (columna D)", options=["(Todas)"] + delegaciones, index=0)
trims_sel = st.multiselect("🗓️ Trimestres", options=["I","II","III","IV"], default=["I","II","III","IV"])

# Filtro de filas (aún sin tocar columnas)
df_filtrado = df_all.copy()
if deleg_sel != "(Todas)":
    df_filtrado = df_filtrado[df_filtrado["Delegación"] == deleg_sel]
if trims_sel:
    df_filtrado = df_filtrado[df_filtrado["Trimestre"].isin(trims_sel)]

# Columnas visibles/editar
cols_base = ["Fecha","Delegación","Trimestre"] + [c for c in [col_tipo, col_obs, "Instituciones"] if c]
cols_mostrar = cols_base + [c for c in cols_HN if c not in cols_base] + [col_pao, COL_SEGUIMIENTO, COL_ACUERDOS]

# Asegurar que EXISTAN en df_all
for c in cols_mostrar:
    if c not in df_all.columns:
        df_all[c] = "" if c != "Fecha" else pd.NaT

# ⚠️ REALINEAR df_filtrado después de crear columnas nuevas
df_filtrado = df_all.loc[df_filtrado.index]

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
for c in yesno_cols.union({COL_SEGUIMIENTO, COL_ACUERDOS}):
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
        st.session_state["df_all"] = df_all  # persistir

def blank_row(trim_label: str):
    base = {k: "" for k in cols_mostrar}
    base["Fecha"] = pd.NaT
    base["Delegación"] = (deleg_sel if deleg_sel != "(Todas)" else "")
    base["Trimestre"]  = trim_label
    for c in yesno_cols.union({COL_SEGUIMIENTO, COL_ACUERDOS}):
        if c in base: base[c] = ""
    base["_row_id"] = str(uuid.uuid4())
    return base

if do_add_iii:
    df_all = pd.concat([df_all, pd.DataFrame([blank_row("III")])], ignore_index=True)
    st.success("Fila base creada en III.")
    st.session_state["df_all"] = df_all

if do_add_iv:
    df_all = pd.concat([df_all, pd.DataFrame([blank_row("IV")])], ignore_index=True)
    st.success("Fila base creada en IV.")
    st.session_state["df_all"] = df_all

if delete_now:
    ids = set(edited.loc[edited["Eliminar"] == True, "_row_id"].astype(str).tolist())
    if ids:
        df_all = df_all[~df_all["_row_id"].astype(str).isin(ids)]
        st.success(f"Eliminadas {len(ids)} fila(s).")
        st.session_state["df_all"] = df_all
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
    for c in yesno_cols.union({COL_SEGUIMIENTO, COL_ACUERDOS}):
        if c in df_all.columns:
            df_all[c] = df_all[c].map(_norm_yesno)
    st.success("Cambios guardados.")
    st.session_state["df_all"] = df_all

# ===================== 4) Formulario =====================
st.subheader("4) Formulario rápido para agregar filas")
with st.form("form_add"):
    a, b, c, d = st.columns(4)
    fecha_new = a.date_input("Fecha", value=date.today())
    trim_new  = b.selectbox("Trimestre", ["I","II","III","IV"], index=2)
    deleg_new = c.selectbox("Delegación", sorted([deleg_sel] + delegaciones) if delegaciones else [""])
    pao_new   = d.selectbox("Validación PAO", ["", "Sí", "No"], index=0)

    # Sí/No explícitos para las 2 columnas canónicas
    seg_new = st.selectbox(COL_SEGUIMIENTO, ["", "Sí", "No"], index=0)
    acu_new = st.selectbox(COL_ACUERDOS, ["", "Sí", "No"], index=0)

    tipo_new = ""
    if st.session_state["col_tipo"]:
        tipos_cat = ["Rendición de cuentas","Seguimiento","Líneas de acción","Informe territorial"]
        tipo_new = st.multiselect("Tipo de actividad (multi)", tipos_cat, default=[])
        tipo_new = "; ".join(tipo_new) if tipo_new else ""

    obs_new  = st.text_area(st.session_state["col_obs"] or "Observaciones", height=100)
    inst_new = st.text_input("Instituciones", "", placeholder="Ingrese instituciones involucradas…")

    st.markdown("**Completar columnas H–N**")
    valores_hn = {}
    for col in cols_HN:
        key = f"hn_{abs(hash(col))}"  # clave única por columna
        if col in yesno_cols:
            valores_hn[col] = st.selectbox(col, ["", "Sí", "No"], index=0, key=key)
        else:
            valores_hn[col] = st.text_input(col, value="", key=key)

    enviar = st.form_submit_button("➕ Agregar registro")

if enviar:
    nuevo = {
        "Fecha": pd.to_datetime(fecha_new),
        "Delegación": deleg_new,
        "Trimestre": trim_new,
        "Instituciones": inst_new,
        "_row_id": str(uuid.uuid4()),
        COL_SEGUIMIENTO: seg_new,
        COL_ACUERDOS: acu_new,
    }
    nuevo[st.session_state["col_pao"] if st.session_state["col_pao"] in df_all.columns else "Validación PAO"] = pao_new
    if st.session_state["col_tipo"]: nuevo[st.session_state["col_tipo"]] = tipo_new
    if st.session_state["col_obs"]:  nuevo[st.session_state["col_obs"]]  = obs_new
    for col in cols_HN:
        if col not in df_all.columns: df_all[col] = ""
        nuevo[col] = valores_hn.get(col, "")
    df_all = pd.concat([df_all, pd.DataFrame([nuevo])], ignore_index=True)
    for c in yesno_cols.union({COL_SEGUIMIENTO, COL_ACUERDOS}):
        if c in df_all.columns: df_all[c] = df_all[c].map(_norm_yesno)
    st.session_state["df_all"] = df_all
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

st.caption("Detección flexible de nombres de hoja (I/1er/T1/Q1/etc.), datos persistentes en session_state y formulario con Sí/No para Seguimiento y Acuerdos.")


