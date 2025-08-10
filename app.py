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
    """
    Crea/actualiza 'Delegación' SIEMPRE desde la columna D (índice 3).
    Elimina otras columnas con nombre parecido a delegación.
    """
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
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).strip().lower()
    if s in {"si", "sí", "s", "yes", "y"}:
        return "Sí"
    if s in {"no", "n"}:
        return "No"
    return ""

def _is_yesno_col(series: pd.Series) -> bool:
    vals = set(_norm_yesno(v) for v in series.dropna().unique())
    return vals.issubset({"Sí", "No"}) and len(vals) <= 2

def detect_yesno_cols(df: pd.DataFrame, extra_binary_cols=None):
    yesno = set(extra_binary_cols or [])
    for c in df.columns:
        if c in {"_row_id", "Fecha", "Delegación", "Trimestre"}:
            continue
        if df[c].dtype == "O" and _is_yesno_col(df[c]):
            yesno.add(c)
    return list(yesno)

def normalize_yesno(df: pd.DataFrame, yesno_cols: list[str]) -> pd.DataFrame:
    df = df.copy()
    for c in yesno_cols:
        if c in df.columns:
            df[c] = df[c].map(_norm_yesno)
    return df

def export_xlsx_force_4_sheets(dfs_by_trim: dict, filename: str):
    """Escribe SIEMPRE I/II/III/IV; si una hoja no tiene filas, se exporta con solo encabezados."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sample = next((d for d in dfs_by_trim.values() if d is not None and not d.empty), None)
        cols = list(sample.columns.drop("_row_id")) if sample is not None else []
        for lab, title in [("I", "I Trimestre"), ("II", "II Trimestre"),
                           ("III", "III Trimestre"), ("IV", "IV Trimestre")]:
            df = dfs_by_trim.get(lab)
            if df is None or df.drop(columns=[c for c in ["_row_id"] if c in df.columns]).empty:
                pd.DataFrame(columns=cols).to_excel(writer, index=False, sheet_name=title[:31])
            else:
                df.drop(columns=[c for c in ["_row_id"] if c in df.columns]) \
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
uploaded = st.file_uploader("📂 Sube el Excel (IT/IIT o I/II/III/IV). Soporta 1–4 hojas.", type=["xlsx", "xlsm"])
if not uploaded:
    st.info("Sube un archivo para continuar.")
    st.stop()

xls = pd.ExcelFile(uploaded)
sheet_names = xls.sheet_names

# Detectar por nombre; lo que falte se asigna por orden I → II → III → IV
PATTERNS = [
    (r"^(it|i\s*tr|1t|primer|1)\b", "I"),
    (r"^(iit|ii\s*tr|2t|seg|segundo|2)\b", "II"),
    (r"^(iii|iii\s*tr|3t|terc|tercero|3)\b", "III"),
    (r"^(iv|iv\s*tr|4t|cuart|cuarto|4)\b", "IV"),
]
def guess_trim(sn: str) -> str:
    s = sn.strip().lower()
    for pat, lab in PATTERNS:
        if re.search(pat, s):
            return lab
    return ""

mapped, used = {}, set()
for s in sheet_names:
    lab = guess_trim(s)
    if lab and lab not in used:
        mapped[s] = lab; used.add(lab)
for s in sheet_names:
    if s not in mapped:
        for lab in ["I", "II", "III", "IV"]:
            if lab not in used:
                mapped[s] = lab; used.add(lab); break

# Construir un DF independiente por trimestre
dfs = {"I": None, "II": None, "III": None, "IV": None}
for s, lab in mapped.items():
    if lab not in dfs:
        continue
    df = pd.read_excel(xls, sheet_name=s)
    df = clean_cols(df)
    df = standardize_delegacion_from_colD(df)
    df = ensure_columns(df, ["Fecha", "Delegación", "Trimestre", "Instituciones"])
    df["Trimestre"] = lab
    df = ensure_row_id(df)
    dfs[lab] = df if dfs[lab] is None else pd.concat([dfs[lab], df], ignore_index=True)

# DFs vacíos con columnas mínimas si faltan
BASE_COLS = ["Fecha", "Delegación", "Trimestre", "Instituciones", "Validación PAO"]
for lab in dfs.keys():
    if dfs[lab] is None:
        dfs[lab] = ensure_row_id(pd.DataFrame(columns=BASE_COLS))

# ================= Editor por hoja (independiente) =================
st.subheader("2) Editar por hoja (cada trimestre es independiente)")

def hoja_editor(label: str):
    st.markdown(f"### {label} Trimestre")
    df = dfs[label].copy()

    # Detectar/normalizar Sí/No (incluir Validación PAO)
    pao_col = pao_col_name(df)
    if pao_col not in df.columns:
        df[pao_col] = ""
    yesno_cols = detect_yesno_cols(df, extra_binary_cols=[pao_col])
    df = normalize_yesno(df, yesno_cols)

    # Editor
    df["_Eliminar"] = False
    col_cfg = {
        "_row_id": st.column_config.TextColumn("ID", disabled=True),
        "_Eliminar": st.column_config.CheckboxColumn("Eliminar"),
        "Fecha": st.column_config.DateColumn("Fecha"),
        "Trimestre": st.column_config.SelectboxColumn("Trimestre", options=["I","II","III","IV"], disabled=True),
    }
    for c in yesno_cols:
        if c in df.columns:
            col_cfg[c] = st.column_config.SelectboxColumn(c, options=["", "Sí", "No"])

    edited = st.data_editor(
        df, num_rows="dynamic", use_container_width=True, height=420,
        column_config=col_cfg, hide_index=True, key=f"ed_{label}"
    )

    # Acciones
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
                elif new_col in {"_row_id", "_Eliminar"}:
                    st.warning("Nombre reservado.")
                else:
                    edited[new_col] = ""
                    st.experimental_rerun()

    with c3:
        if st.button("🗑️ Eliminar seleccionados", key=f"del_{label}", use_container_width=True):
            to_del = set(edited.loc[edited["_Eliminar"] == True, "_row_id"].astype(str))
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

    # Formulario rápido para esta hoja
    with st.expander("➕ Formulario rápido (agrega 1 fila a esta hoja)"):
        a, b, c, d = st.columns(4)
        f_new = a.date_input("Fecha", value=date.today(), key=f"f_{label}")
        # delegaciones sugeridas por lo ya existente
        delegs = sorted([d for d in edited.get("Delegación", pd.Series(dtype=str))
                         .dropna().astype(str).map(str.strip).unique() if d])
        d_new = b.selectbox("Delegación", options=delegs + [""], index=len(delegs) if delegs else 0, key=f"del_{label}")
        pao_new = c.selectbox("Validación PAO", ["", "Sí", "No"], key=f"pao_{label}")
        inst_new = d.text_input("Instituciones", key=f"inst_{label}", placeholder="Ingrese instituciones…")
        # Observaciones (si existe la columna)
        obs_col = next((cn for cn in edited.columns if re.fullmatch(r"observaciones?\.?", cn, re.I)), None)
        if obs_col:
            obs_val = st.text_area(obs_col, key=f"obs_{label}")
        else:
            obs_val = st.text_area("Observaciones", key=f"obs_{label}")
            edited["Observaciones"] = edited.get("Observaciones", "")
            obs_col = "Observaciones"

        if st.button("Agregar registro", key=f"add_form_{label}"):
            nuevo = {
                "_row_id": str(uuid.uuid4()),
                "Fecha": pd.to_datetime(f_new),
                "Delegación": d_new,
                "Trimestre": label,
                "Instituciones": inst_new,
                pao_col_name(edited): pao_new,
                obs_col: obs_val
            }
            # Rellenar demás columnas con vacío
            for ccol in edited.columns:
                if ccol not in nuevo:
                    nuevo[ccol] = "" if ccol != "Fecha" else pd.NaT
            edited.loc[len(edited)] = nuevo
            dfs[label] = edited.drop(columns=["_Eliminar"], errors="ignore")
            st.success("Registro agregado.")
            st.experimental_rerun()

# Tabs por hoja (cada una independiente)
t1, t2, t3, t4 = st.tabs(["I Trimestre", "II Trimestre", "III Trimestre", "IV Trimestre"])
with t1: hoja_editor("I")
with t2: hoja_editor("II")
with t3: hoja_editor("III")
with t4: hoja_editor("IV")

# ================= Exportar =================
st.subheader("3) Exportar Excel (siempre 4 hojas)")
export_xlsx_force_4_sheets(dfs, filename="seguimiento_trimestres_independiente.xlsx")

st.caption("Cada pestaña maneja su propia hoja. Agregar/editar/eliminar en una no afecta a las demás. Delegación siempre proviene de la columna D del archivo cargado.")









