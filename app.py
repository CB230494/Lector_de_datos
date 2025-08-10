# app.py
import streamlit as st
import pandas as pd
import io
import re
import uuid
from datetime import date

st.set_page_config(page_title="Seguimiento por Trimestre — Hojas independientes", layout="wide")
st.title("📘 Seguimiento por Trimestre — Hojas I/II/III/IV independientes")

# ================= Utilidades =================
def clean_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def standardize_delegacion_from_colD(df: pd.DataFrame) -> pd.DataFrame:
    """Crea/actualiza 'Delegación' SIEMPRE desde la columna D (índice 3) y elimina otras similares."""
    df = df.copy()
    if df.shape[1] > 3:
        df["Delegación"] = df.iloc[:, 3]
    else:
        if "Delegación" not in df.columns:
            df["Delegación"] = ""
    drop_like = [c for c in df.columns if c != "Delegación" and re.search(r"delegaci[oó]n", str(c), re.I)]
    if drop_like:
        df = df.drop(columns=drop_like)
    return df

def ensure_columns(df: pd.DataFrame, needed: list[str]) -> pd.DataFrame:
    df = df.copy()
    for c in needed:
        if c not in df.columns:
            df[c] = pd.NaT if c == "Fecha" else ""
    return df

def ensure_row_id(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "_row_id" not in df.columns:
        df["_row_id"] = [str(uuid.uuid4()) for _ in range(len(df))]
    return df

# --- Sí/No helpers
def _norm_yesno(x):
    if x is None or (isinstance(x, float) and pd.isna(x)): return ""
    s = str(x).strip().lower()
    if s in {"si","sí","s","yes","y"}: return "Sí"
    if s in {"no","n"}: return "No"
    return ""

def _is_yesno_col(series: pd.Series) -> bool:
    vals = set(_norm_yesno(v) for v in series.dropna().unique())
    return vals.issubset({"Sí","No"}) and len(vals) <= 2

def detect_yesno_cols(df: pd.DataFrame, extra_binary_cols=None):
    yesno = set(extra_binary_cols or [])
    for c in df.columns:
        if c in {"_row_id","Fecha","Delegación","Trimestre"}: continue
        if df[c].dtype == "O" and _is_yesno_col(df[c]): yesno.add(c)
    return list(yesno)

def normalize_yesno(df: pd.DataFrame, yesno_cols: list[str]) -> pd.DataFrame:
    df = df.copy()
    for c in yesno_cols:
        if c in df.columns: df[c] = df[c].map(_norm_yesno)
    return df

def export_xlsx_force_4_sheets(dfs_by_trim: dict, filename: str):
    """Escribe SIEMPRE I/II/III/IV; si una hoja no tiene filas, se exporta con solo encabezados."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sample = next((d for d in dfs_by_trim.values() if d is not None and not d.empty), None)
        cols = list(sample.columns.drop("_row_id")) if sample is not None else []
        for lab, title in [("I","I Trimestre"),("II","II Trimestre"),("III","III Trimestre"),("IV","IV Trimestre")]:
            df = dfs_by_trim.get(lab)
            if df is None or df.drop(columns=[c for c in ["_row_id"] if c in df.columns]).empty:
                pd.DataFrame(columns=cols).to_excel(writer, index=False, sheet_name=title[:31])
            else:
                df.drop(columns=[c for c in ["_row_id"] if c in df.columns])\
                  .to_excel(writer, index=False, sheet_name=title[:31])
    st.download_button("📥 Descargar Excel (4 hojas)",
                       data=output.getvalue(),
                       file_name=filename,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

def pao_col_name(df: pd.DataFrame) -> str:
    c = next((c for c in df.columns if re.search(r"validaci[oó]n\s*pao", c, re.I)), None)
    return c or "Validación PAO"

# ================= Cargar archivo base =================
st.subheader("1) Cargar archivo base")
uploaded = st.file_uploader("📂 Sube el Excel (IT/IIT o I/II/III/IV). Soporta 1–4 hojas.", type=["xlsx","xlsm"])
if not uploaded:
    st.info("Sube un archivo para continuar.")
    st.stop()

xls = pd.ExcelFile(uploaded)
sheet_names = xls.sheet_names

# Detectar por nombre; lo que falte se asigna por orden I → II → III → IV
PATTERNS = [
    (r"^(it|i\s*tr|1t|primer|1)\b","I"),
    (r"^(iit|ii\s*tr|2t|seg|segundo|2)\b","II"),
    (r"^(iii|iii\s*tr|3t|terc|tercero|3)\b","III"),
    (r"^(iv|iv\s*tr|4t|cuart|cuarto|4)\b","IV"),
]
def guess_trim(sn: str) -> str:
    s = sn.strip().lower()
    for pat, lab in PATTERNS:
        if re.search(pat, s): return lab
    return ""

mapped, used = {}, set()
for s in sheet_names:
    lab = guess_trim(s)
    if lab and lab not in used:
        mapped[s] = lab; used.add(lab)
for s in sheet_names:
    if s not in mapped:
        for lab in ["I","II","III","IV"]:
            if lab not in used:
                mapped[s] = lab; used.add(lab); break

# Construir DF independiente por trimestre + detectar columnas H..N del archivo original
dfs = {"I": None, "II": None, "III": None, "IV": None}
cols_HN_global = set()
for s, lab in mapped.items():
    if lab not in dfs: continue
    raw = pd.read_excel(xls, sheet_name=s)
    # columnas por posición H..N (0-based 7..13)
    hn = list(raw.columns[7:14]) if raw.shape[1] > 7 else []
    cols_HN_global.update([str(c) for c in hn])

    df = clean_cols(raw)
    df = standardize_delegacion_from_colD(df)
    df = ensure_columns(df, ["Fecha","Delegación","Trimestre","Instituciones"])
    df["Trimestre"] = lab
    df = ensure_row_id(df)
    dfs[lab] = df if dfs[lab] is None else pd.concat([dfs[lab], df], ignore_index=True)

cols_HN_global = [c for c in cols_HN_global if c not in {"Fecha","Delegación","Trimestre","Instituciones"}]

# DFs vacíos con columnas mínimas si faltan
BASE_COLS = ["Fecha","Delegación","Trimestre","Instituciones","Validación PAO"]
for lab in dfs.keys():
    if dfs[lab] is None:
        dfs[lab] = ensure_row_id(pd.DataFrame(columns=BASE_COLS))

# ================= Editor por hoja (independiente) =================
st.subheader("2) Editar por hoja (cada trimestre es independiente)")

def hoja_editor(label: str):
    st.markdown(f"### {label} Trimestre")
    df = dfs[label].copy()

    pao_col = pao_col_name(df)
    if pao_col not in df.columns: df[pao_col] = ""
    yesno_cols = detect_yesno_cols(df, extra_binary_cols=[pao_col])
    df = normalize_yesno(df, yesno_cols)

    df["_Eliminar"] = False
    col_cfg = {
        "_row_id": st.column_config.TextColumn("ID", disabled=True),
        "_Eliminar": st.column_config.CheckboxColumn("Eliminar"),
        "Fecha": st.column_config.DateColumn("Fecha"),
        "Trimestre": st.column_config.SelectboxColumn("Trimestre", options=["I","II","III","IV"], disabled=True),
    }
    for c in yesno_cols:
        if c in df.columns:
            col_cfg[c] = st.column_config.SelectboxColumn(c, options=["","Sí","No"])

    edited = st.data_editor(
        df, num_rows="dynamic", use_container_width=True, height=420,
        column_config=col_cfg, hide_index=True, key=f"ed_{label}"
    )

    c1, c2, c3, c4 = st.columns([1,1,2,2])
    with c1:
        if st.button(f"➕ Agregar fila en {label}", key=f"addrow_{label}", use_container_width=True):
            new = {c: "" for c in edited.columns}
            new["Fecha"] = pd.NaT
            new["Trimestre"] = label
            new["_row_id"] = str(uuid.uuid4())
            new["_Eliminar"] = False
            edited.loc[len(edited)] = new
            st.experimental_rerun()

    with c2:
        new_col = st.text_input("Nueva columna", key=f"newcol_{label}", placeholder="Nombre…")
        if st.button("➕ Agregar columna", key=f"addcol_{label}", use_container_width=True):
            if new_col:
                if new_col in edited.columns:
                    st.warning("Ya existe esa columna.")
                elif new_col in {"_row_id","_Eliminar"}:
                    st.warning("Nombre reservado.")
                else:
                    edited[new_col] = ""
                    st.experimental_rerun()

    with c3:
        if st.button("🗑️ Eliminar seleccionados", key=f"btn_delete_{label}", use_container_width=True):
            to_del = set(edited.loc[edited["_Eliminar"]==True, "_row_id"].astype(str))
            if to_del:
                edited = edited[~edited["_row_id"].astype(str).isin(to_del)]
                st.success(f"Eliminadas {len(to_del)} fila(s).")
                st.experimental_rerun()
            else:
                st.info("Marca 'Eliminar' en al menos una fila.")

    with c4:
        if st.button("💾 Guardar cambios de esta hoja", key=f"save_{label}", use_container_width=True):
            out = edited.drop(columns=["_Eliminar"], errors="ignore")
            out = normalize_yesno(out, yesno_cols)
            dfs[label] = out.copy()
            st.success("Cambios guardados en memoria.")

# Tabs por hoja
t1, t2, t3, t4 = st.tabs(["I Trimestre","II Trimestre","III Trimestre","IV Trimestre"])
with t1: hoja_editor("I")
with t2: hoja_editor("II")
with t3: hoja_editor("III")
with t4: hoja_editor("IV")

# ================= FORMULARIO GLOBAL COMPLETO =================
st.subheader("3) Formulario (completo) para agregar registros")

# Delegaciones sugeridas (de todas las hojas)
all_delegs = sorted(
    set(
        d for lab in dfs
        for d in dfs[lab].get("Delegación", pd.Series(dtype=str)).dropna().astype(str).map(str.strip).tolist()
        if d
    )
)

# Buscar nombres de columnas existentes para tipo y observaciones
def find_col_any(pat: str) -> str | None:
    for lab in dfs:
        for c in dfs[lab].columns:
            if re.fullmatch(pat, c, flags=re.I):
                return c
    return None
col_tipo_any = find_col_any(r"tipo\s*de\s*actividad\.?")
col_obs_any  = find_col_any(r"observaciones?\.?")

# Columna PAO global
def any_df() -> pd.DataFrame:
    for lab in dfs:
        if not dfs[lab].empty: return dfs[lab]
    return pd.concat(dfs.values()).head(0)
pao_global_col = pao_col_name(any_df())

with st.form("form_add_global"):
    a, b, c, d = st.columns(4)
    fecha_new = a.date_input("Fecha", value=date.today())
    trim_new  = b.selectbox("Trimestre", ["I","II","III","IV"], index=2)
    deleg_new = c.selectbox("Delegación", all_delegs + [""] if all_delegs else [""])
    pao_new   = d.selectbox("Validación PAO", ["","Sí","No"], index=0)

    tipo_new = ""
    if col_tipo_any:
        tipos_cat = ["Rendición de cuentas","Seguimiento","Líneas de acción","Informe territorial"]
        tipo_new = st.multiselect("Tipo de actividad (multi)", tipos_cat, default=[])
        tipo_new = "; ".join(tipo_new) if tipo_new else ""

    obs_new  = st.text_area(col_obs_any or "Observaciones", height=120)
    inst_new = st.text_input("Instituciones", "", placeholder="Ingrese instituciones…")

    st.markdown("**Completar columnas H–N**")
    # Determinar si alguna hoja usa estas columnas como Sí/No
    yesno_lookup = {}
    for col in cols_HN_global:
        is_yesno = False
        for lab in dfs:
            if col in dfs[lab].columns and _is_yesno_col(dfs[lab][col]):
                is_yesno = True; break
        yesno_lookup[col] = is_yesno

    valores_hn = {}
    for col in cols_HN_global:
        if yesno_lookup[col]:
            valores_hn[col] = st.selectbox(col, ["","Sí","No"], index=0, key=f"hn_{col}")
        else:
            valores_hn[col] = st.text_input(col, value="", key=f"hn_{col}")

    enviar = st.form_submit_button("➕ Agregar registro")

if enviar:
    # Asegurar columnas en la hoja destino
    for col in [pao_global_col, col_tipo_any, col_obs_any, "Instituciones"] + cols_HN_global:
        if col and col not in dfs[trim_new].columns:
            dfs[trim_new][col] = ""

    nuevo = {
        "_row_id": str(uuid.uuid4()),
        "Fecha": pd.to_datetime(fecha_new),
        "Delegación": deleg_new,
        "Trimestre": trim_new,
        "Instituciones": inst_new,
        pao_global_col: pao_new
    }
    if col_tipo_any: nuevo[col_tipo_any] = tipo_new
    if col_obs_any:  nuevo[col_obs_any]  = obs_new
    for col in cols_HN_global:
        nuevo[col] = valores_hn.get(col, "")

    dfs[trim_new] = pd.concat([dfs[trim_new], pd.DataFrame([nuevo])], ignore_index=True)
    st.success(f"Registro agregado en {trim_new}.")

# ================= Exportar =================
st.subheader("4) Exportar Excel (siempre 4 hojas)")
export_xlsx_force_4_sheets(dfs, filename="seguimiento_trimestres_independiente.xlsx")

st.caption("Editor por hoja + formulario global completo. Cada trimestre es independiente y Delegación se toma de la columna D del archivo cargado.")










