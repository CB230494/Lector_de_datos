import streamlit as st
import pandas as pd
import io
import re
import uuid

st.set_page_config(page_title="Seguimiento por Trimestre (IT/IIT) — Editar y crear III/IV", layout="wide")
st.title("📘 Seguimiento por Trimestre — Lector + Editor + Formulario")

# ------------------------ helpers ------------------------
def clean_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def take_cols_H_to_N(df: pd.DataFrame):
    """Nombres de columnas H..N por posición (Excel H..N → 0-based 7..13)."""
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
    df = df.drop(columns=drop_like)
    return df

def find_col_by_exact(df, pat):
    for c in df.columns:
        if re.fullmatch(pat, c, flags=re.I):
            return c
    return None

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

def ensure_row_id(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega un ID estable por fila para poder reconciliar ediciones."""
    df = df.copy()
    if "_row_id" not in df.columns:
        df["_row_id"] = [str(uuid.uuid4()) for _ in range(len(df))]
    return df

# ------------------------ 1) Cargar archivo base (IT y IIT) ------------------------
st.subheader("1) Cargar archivo base (IT y IIT)")
archivo_base = st.file_uploader("📁 Sube el Excel (contiene IT e IIT)", type=["xlsx", "xlsm"])
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

# Estandarizar Delegación desde D
df_it  = standardize_delegacion_from_colD(df_it)
df_iit = standardize_delegacion_from_colD(df_iit)

# Trimestre
df_it  = add_trimestre(df_it, "I")
df_iit = add_trimestre(df_iit, "II")

# Columnas H..N para usar en formulario/edición
cols_HN_it  = take_cols_H_to_N(df_it)
cols_HN_iit = take_cols_H_to_N(df_iit)
cols_HN = cols_HN_it if len(cols_HN_it) >= len(cols_HN_iit) else cols_HN_iit

# Detectar columnas tipo/observaciones si existen
col_tipo = find_col_by_exact(df_it, r"tipo\s*de\s*actividad\.?") or find_col_by_exact(df_iit, r"tipo\s*de\s*actividad\.?")
col_obs  = find_col_by_exact(df_it, r"observaciones\.?") or find_col_by_exact(df_iit, r"observaciones\.?")

# Consolidado + ID de fila
df_all = pd.concat([df_it, df_iit], ignore_index=True)
df_all = ensure_row_id(df_all)

# ------------------------ 2) Filtros ------------------------
st.subheader("2) Filtros")
delegaciones = sorted([d for d in df_all["Delegación"].dropna().astype(str).map(str.strip).unique() if d])
deleg_sel = st.selectbox("Delegación (columna D)", options=["(Todas)"] + delegaciones, index=0)
trims_sel = st.multiselect("Trimestres", options=["I","II","III","IV"], default=["I","II"])

df_filtrado = df_all.copy()
if deleg_sel != "(Todas)":
    df_filtrado = df_filtrado[df_filtrado["Delegación"] == deleg_sel]
if trims_sel:
    df_filtrado = df_filtrado[df_filtrado["Trimestre"].isin(trims_sel)]

# Columnas principales a mostrar/editar
cols_base = ["Delegación", "Trimestre"] + [c for c in [col_tipo, col_obs] if c]
cols_mostrar = cols_base + [c for c in cols_HN if c not in cols_base]
# Añadir el ID oculto para poder reconciliar
cols_editor = cols_mostrar + ["_row_id"]
df_ed = df_filtrado[cols_editor].copy()

# ------------------------ 3) Editor (editar / agregar / eliminar) ------------------------
st.subheader("3) Editor por delegación (puedes editar, agregar o eliminar filas)")

# Agregar columna de control 'Eliminar' para permitir borrar filas desde el editor
df_ed["Eliminar"] = False

edited = st.data_editor(
    df_ed,
    num_rows="dynamic",  # permite agregar nuevas filas
    use_container_width=True,
    height=420,
    column_config={
        "_row_id": st.column_config.TextColumn("ID (interno)", disabled=True),
        "Eliminar": st.column_config.CheckboxColumn("Eliminar"),
    },
    hide_index=True,
    key="editor",
)

colE1, colE2, colE3 = st.columns(3)
with colE1:
    do_quick_add_iii = st.button("➕ Agregar fila base a III", help="Crea una fila en blanco para el Trimestre III, usando la delegación seleccionada.")
with colE2:
    do_quick_add_iv  = st.button("➕ Agregar fila base a IV", help="Crea una fila en blanco para el Trimestre IV, usando la delegación seleccionada.")
with colE3:
    apply_changes = st.button("💾 Guardar cambios en el consolidado")

# Crear fila base III/IV para la delegación seleccionada
def blank_row_for_trim(trim_label: str):
    base = {k: "" for k in cols_mostrar}
    base["Delegación"] = (deleg_sel if deleg_sel != "(Todas)" else "")
    base["Trimestre"]  = trim_label
    base["_row_id"]    = str(uuid.uuid4())
    return base

if do_quick_add_iii:
    new_row = blank_row_for_trim("III")
    df_all = pd.concat([df_all, pd.DataFrame([new_row])], ignore_index=True)
    st.success("Se agregó una fila base al Trimestre III. (Aparecerá cuando recargues filtros o al exportar)")

if do_quick_add_iv:
    new_row = blank_row_for_trim("IV")
    df_all = pd.concat([df_all, pd.DataFrame([new_row])], ignore_index=True)
    st.success("Se agregó una fila base al Trimestre IV. (Aparecerá cuando recargues filtros o al exportar)")

# Aplicar cambios del editor al consolidado
if apply_changes:
    # Separar filas marcadas para eliminar
    to_delete_ids = set(edited.loc[edited["Eliminar"] == True, "_row_id"].astype(str).tolist())

    # Filas editadas (sin la columna 'Eliminar')
    edited_clean = edited.drop(columns=["Eliminar"]).copy()

    # 1) Eliminar en df_all las filas seleccionadas
    if to_delete_ids:
        df_all = df_all[~df_all["_row_id"].astype(str).isin(to_delete_ids)]

    # 2) Reconcilia: para cada _row_id presente en edited_clean, actualiza columnas visibles
    #    (editor pudo agregar filas nuevas con _row_id vacío → creamos uno)
    for idx, row in edited_clean.iterrows():
        rid = str(row["_row_id"]) if pd.notna(row["_row_id"]) and str(row["_row_id"]).strip() else str(uuid.uuid4())
        # Si es fila nueva (no existe en df_all), la insertamos
        if rid not in set(df_all["_row_id"].astype(str)):
            new_entry = {c: row.get(c, "") for c in cols_mostrar}
            new_entry["_row_id"] = rid
            df_all = pd.concat([df_all, pd.DataFrame([new_entry])], ignore_index=True)
        else:
            # actualizar columnas editables
            mask = df_all["_row_id"].astype(str).eq(rid)
            for c in cols_mostrar:
                if c in edited_clean.columns:
                    df_all.loc[mask, c] = row.get(c, "")

    st.success("Cambios aplicados al consolidado.")

# ------------------------ 4) Formulario rápido (opcional, sigue disponible) ------------------------
st.subheader("4) Formulario rápido para agregar filas")
with st.form("form_add_quick"):
    c1, c2, c3 = st.columns(3)
    trim_new = c1.selectbox("Trimestre", ["I","II","III","IV"], index=2)  # por defecto III
    pao_new  = c2.selectbox("Validación PAO", ["Sí", "No"], index=0)
    deleg_new = c3.selectbox("Delegación", delegaciones if delegaciones else [deleg_sel if deleg_sel!="(Todas)" else ""])

    # Intentar reutilizar nombres de columnas si existen
    tipos_catalogo = ["Rendición de cuentas", "Seguimiento", "Líneas de acción", "Informe territorial"]
    tipo_new = ""
    if col_tipo:
        tipo_new = st.multiselect("Tipo de actividad (multi)", tipos_catalogo, default=[])
        tipo_new = "; ".join(tipo_new) if tipo_new else ""
    obs_new = st.text_area("Observaciones", height=100) if col_obs else ""

    st.markdown("**Completar columnas H–N (nombres reales)**")
    valores_hn = {}
    for col in cols_HN:
        valores_hn[col] = st.text_input(col, value="")

    enviado = st.form_submit_button("➕ Agregar registro")

if enviado:
    nuevo = {"Delegación": deleg_new, "Trimestre": trim_new, "_row_id": str(uuid.uuid4())}
    col_pao = next((c for c in df_all.columns if re.search(r"validaci[oó]n\s*pao", c, re.I)), "Validación PAO")
    nuevo[col_pao] = pao_new
    if col_tipo: nuevo[col_tipo] = tipo_new
    if col_obs:  nuevo[col_obs]  = obs_new
    for col in cols_HN:
        nuevo[col] = valores_hn.get(col, "")
    df_all = pd.concat([df_all, pd.DataFrame([nuevo])], ignore_index=True)
    st.success("Registro agregado al consolidado.")

# ------------------------ 5) Generar Excel (nuevo o actualizar) ------------------------
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

            # Estandarizar 'Delegación' en el archivo anterior si no existe
            if "Delegación" not in old_df.columns:
                if old_df.shape[1] > 3:
                    old_df["Delegación"] = old_df.iloc[:, 3]
                else:
                    old_df["Delegación"] = ""

            # Quitar columnas parecidas a 'delegación'
            drop_like = [c for c in old_df.columns if c != "Delegación" and re.search(r"delegaci[oó]n", str(c), re.I)]
            old_df = old_df.drop(columns=drop_like)

            df_final = pd.concat([old_df, df_all], ignore_index=True)
            st.info(f"Se combinó el archivo anterior ({len(old_df)} filas) con el actual.")
        except Exception as e:
            st.error(f"No se pudo leer el archivo anterior: {e}")

# Quitar duplicados exactos (ignoramos el ID)
if "_row_id" in df_final.columns:
    export_cols = [c for c in df_final.columns if c != "_row_id"]
else:
    export_cols = list(df_final.columns)
df_final = df_final[export_cols].drop_duplicates()

with st.expander("🔎 Vista previa del Excel a generar"):
    st.dataframe(df_final, use_container_width=True, height=420)

# Exportar creando hojas por trimestre (I, II, III, IV) según existan
sheets = {}
for t in ["I","II","III","IV"]:
    parte = df_final[df_final["Trimestre"] == t]
    if not parte.empty:
        sheets[f"{t} Trimestre"] = parte
if not sheets:
    sheets = {"Datos": df_final}

export_xlsx(sheets, filename="seguimiento_trimestres_generado.xlsx")

st.caption(
    "Ahora puedes **editar filas** del filtro por delegación, **agregar** filas nuevas, "
    "y crear registros en **III/IV**. Al exportar, si hay filas de esos trimestres, "
    "se crean las hojas correspondientes automáticamente."
)











