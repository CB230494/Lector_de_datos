# =========================
# 📋 Asistencia – Público + Admin (admin oculto hasta login)
# =========================
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date, time, datetime
from typing import List

st.set_page_config(page_title="Asistencia – Registro y Admin", layout="wide")

# ---------- Backend de datos: Google Sheets ----------
import gspread
try:
    from zoneinfo import ZoneInfo  # Python 3.9+
except Exception:
    ZoneInfo = None

# ⚠️ CONEXIÓN A LA HOJA (actualizado al nuevo sheet que indicaste)
# URL: https://docs.google.com/spreadsheets/d/1vzGRJrlUzaCdhJAQBa6i94RE2QxnKqFvXpch9HF4TO8/edit
SHEET_ID = "1vzGRJrlUzaCdhJAQBa6i94RE2QxnKqFvXpch9HF4TO8"
SHEET_NAME = "Hoja 1"   # cambialo si tu pestaña tiene otro nombre

# Estructura final (8 columnas)
HEADER = ["nombre","cedula","delegacion","cargo","telefono","genero","sexo","edad"]

# Catálogo de delegaciones
DELEGACIONES = [
    "Estrategia Sembremos Seguridad",
    "Carmen","Merced","Hospital","Catedral","San Sebastián","Hatillo",
    "Zapote / San Francisco de dos Rios","Pavas","Uruca / Mata Redonda",
    "Curridabat","Montes de Oca","Goicoechea","Moravia","Tibás","Coronado",
    "Desamparados Norte","Desamparados Sur","Aserrí","Acosta","Alajuelita",
    "Escazu","Santa Ana","Mora","Puriscal","Turrabares",
    "Alajuela Sur","Alajuela Norte","San Ramón","Grecia","San Mateo",
    "Atenas","Naranjo","Palmares","Poas","Orotina","Sarchí",
    "Cartago","Paraíso","La Unión","Jiménez","Turrialba","Alvarado","Oreamuno","El Guarco",
    "Tarrazú","Dota","León Cortéz",
    "Guadalupe","Heredia","Barva","Santo Domingo","Santa Barbara","San Rafael","San Isidro","Belén","Flores","San Pablo",
    "Liberia","Nicoya","Santa Cruz","Bagaces","Carrillo","Cañas","Abangares","Tilarán","Nandayure","Hojancha","La Cruz",
    "Puntarenas","Esparza","Montes de Oro","Quepos","Parrita","Garabito","Paquera","Chomes",
    "Pérez Zeledón","Buenos Aires","Osa",
    "San Carlos Este","San Carlos Oeste","Zarcero","Guatuso","Río Cuarto",
    "Limón","Siquires","Talamanca","Matina",
    "Golfito","Coto Brus","Corredores","Puerto Jiménez",
    "Upala","Los Chiles - Cutris - Pocosol","Sarapiquí","Colorado","Pococí","Guacimo",
]

def _sa_key():
    try:
        sa = st.secrets["gcp_service_account"]
        return sa.get("client_email","") + "|" + sa.get("project_id","")
    except Exception:
        return ""

@st.cache_resource(show_spinner=False)
def _get_ws_cached(sheet_id: str, sheet_name: str, sa_key: str):
    if "gcp_service_account" not in st.secrets:
        raise RuntimeError("Falta el bloque [gcp_service_account] en .streamlit/secrets.toml")
    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
    sh = gc.open_by_key(sheet_id)

    # Obtiene/crea la pestaña
    try:
        ws = sh.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows=2000, cols=len(HEADER))
        ws.update("A1:H1", [HEADER])
        try: ws.freeze(rows=1)
        except: pass

    # Migración: si aún existen id/created_at, elimínalos
    try:
        first_row = [h.strip().lower() for h in ws.row_values(1)]
        if len(first_row) >= 2 and first_row[0] == "id" and first_row[1] == "created_at":
            ws.delete_columns(1, 2)
    except Exception:
        pass

    # Asegura encabezado correcto
    first_row = [h.strip().lower() for h in ws.row_values(1)]
    if first_row != HEADER:
        ws.update("A1:H1", [HEADER])
        try: ws.freeze(rows=1)
        except: pass

    return ws

def _get_ws():
    return _get_ws_cached(SHEET_ID, SHEET_NAME, _sa_key())

def init_db():
    _get_ws()

# ---------- CRUD ----------
def insert_row(row: dict):
    ws = _get_ws()
    telefono = row.get("Teléfono","")
    if telefono and not str(telefono).startswith("'"):  # conserva ceros iniciales
        telefono = "'" + str(telefono)

    payload = [
        row.get("Nombre",""),
        row.get("Cédula de Identidad",""),
        row.get("Delegación",""),
        row.get("Cargo",""),
        telefono,
        row.get("Género",""),
        row.get("Sexo",""),
        row.get("Rango de Edad",""),
    ]
    ws.append_row(payload, value_input_option="USER_ENTERED")

def fetch_all_df(include_rownum=True) -> pd.DataFrame:
    ws = _get_ws()
    values = ws.get_all_values()
    if len(values) < 2:
        cols = ["Nº","Nombre","Cédula de Identidad","Delegación","Cargo","Teléfono","Género","Sexo","Rango de Edad"]
        if include_rownum: cols.insert(1, "rownum")
        return pd.DataFrame(columns=cols)

    header = [h.strip().lower() for h in values[0]]
    data_rows = values[1:]

    name_map = {
        "nombre":"Nombre",
        "cedula":"Cédula de Identidad",
        "delegacion":"Delegación",
        "cargo":"Cargo",
        "telefono":"Teléfono",
        "genero":"Género",
        "sexo":"Sexo",
        "edad":"Rango de Edad",
    }

    records = []
    for idx, row in enumerate(data_rows, start=2):  # fila real en sheet
        rec = {}
        for j, key in enumerate(header):
            if key in name_map:
                rec[name_map[key]] = row[j] if j < len(row) else ""
        rec["rownum"] = idx
        records.append(rec)

    df = pd.DataFrame(records)
    if df.empty:
        cols = ["Nº","Nombre","Cédula de Identidad","Delegación","Cargo","Teléfono","Género","Sexo","Rango de Edad"]
        if include_rownum: cols.insert(1, "rownum")
        return pd.DataFrame(columns=cols)

    cols_order = ["rownum","Nombre","Cédula de Identidad","Delegación","Cargo","Teléfono","Género","Sexo","Rango de Edad"]
    df = df[[c for c in cols_order if c in df.columns]]

    df.insert(0, "Nº", range(1, len(df)+1))
    if not include_rownum and "rownum" in df.columns:
        df = df.drop(columns=["rownum"])
    return df

def update_row_by_rownum(rownum:int, row:dict):
    ws = _get_ws()
    payload = [
        row.get("Nombre",""),
        row.get("Cédula de Identidad",""),
        row.get("Delegación",""),
        row.get("Cargo",""),
        row.get("Teléfono",""),
        row.get("Género",""),
        row.get("Sexo",""),
        row.get("Rango de Edad",""),
    ]
    ws.update(f"A{rownum}:H{rownum}", [payload], value_input_option="USER_ENTERED")

def delete_rows_by_rownums(rownums: List[int]):
    if not rownums:
        return
    ws = _get_ws()
    for r in sorted(rownums, reverse=True):
        ws.delete_rows(r)

def delete_all_rows():
    ws = _get_ws()
    used_rows = len(ws.get_all_values())
    if used_rows >= 2:
        ws.batch_clear([f"A2:H{used_rows}"])

# Inicializa backend
try:
    init_db()
except Exception:
    st.error("Error conectando a Google Sheets. Verifica permisos y secrets.")
    st.stop()

# ---------- Login admin ----------
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False

with st.sidebar:
    st.markdown("### 🔐 Acceso administrador")
    if not st.session_state.is_admin:
        pwd = st.text_input("Contraseña", type="password", placeholder="••••••••")
        if st.button("Ingresar"):
            if pwd == "Sembremos23":
                st.session_state.is_admin = True
                st.success("Acceso concedido.")
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
    else:
        st.success("Sesión de administrador activa")
        if st.button("Cerrar sesión"):
            st.session_state.is_admin = False
            st.rerun()

# ---------- Público ----------
st.markdown("# 📋 Asistencia – Registro")
st.markdown("### ➕ Agregar")
with st.form("form_asistencia_publico", clear_on_submit=True):
    c1, c2, c3 = st.columns([1.2, 1, 1])
    nombre      = c1.text_input("Nombre")
    cedula      = c2.text_input("Cédula de Identidad")

    opciones_deleg = ["— Selecciona una delegación —"] + DELEGACIONES
    sel_deleg = c3.selectbox("Delegación", opciones_deleg, index=0)
    delegacion_sel = "" if sel_deleg == opciones_deleg[0] else sel_deleg

    c4, c5 = st.columns([1, 1])
    cargo    = c4.text_input("Cargo")
    telefono = c5.text_input("Teléfono")

    st.markdown("#### ")
    gcol, scol, ecol = st.columns([1.1, 1.5, 1.5])
    genero = gcol.radio("Género", ["F", "M", "LGBTIQ+"], horizontal=True)
    sexo   = scol.radio("Sexo (Hombre, Mujer o Intersex)", ["H", "M", "I"], horizontal=True)
    edad   = ecol.radio("Rango de Edad", ["18 a 35 años", "36 a 64 años", "65 años o más"], horizontal=True)

    submitted = st.form_submit_button("➕ Agregar", use_container_width=True)
    if submitted:
        if not nombre.strip():
            st.warning("Ingresa al menos el nombre.")
        else:
            fila = {
                "Nombre": nombre.strip(),
                "Cédula de Identidad": cedula.strip(),
                "Delegación": delegacion_sel.strip(),
                "Cargo": cargo.strip(),
                "Teléfono": telefono.strip(),
                "Género": genero,
                "Sexo": sexo,
                "Rango de Edad": edad
            }
            insert_row(fila)
            st.success("Registro guardado.")

st.markdown("### 📥 Registros recibidos")
df_pub = fetch_all_df(include_rownum=False)
if not df_pub.empty:
    st.dataframe(
        df_pub[["Nº","Nombre","Cédula de Identidad","Delegación","Cargo","Teléfono","Género","Sexo","Rango de Edad"]],
        use_container_width=True, hide_index=True
    )
else:
    st.info("Aún no hay registros guardados.")

# ---------- Admin ----------
if st.session_state.is_admin:
    st.markdown("---")
    st.markdown("# 🛠️ Panel del Administrador")

    df_all = fetch_all_df(include_rownum=True)
    if df_all.empty:
        st.info("Aún no hay registros guardados.")
        st.stop()

    # === Multiselección de delegaciones ===
    delegs_existentes = sorted([d for d in df_all["Delegación"].dropna().unique() if str(d).strip()], key=str.casefold)
    sel_filtros = st.multiselect(
        "Filtrar por Delegación",
        options=delegs_existentes,
        default=[],
        help="Vacío = todas. Puedes elegir varias delegaciones."
    )

    if not sel_filtros:
        df_view = df_all.copy().reset_index(drop=True)
    else:
        df_view = df_all[df_all["Delegación"].isin(sel_filtros)].reset_index(drop=True)

    # Encabezado Excel
    st.markdown("### 🧾 Datos de encabezado (Excel)")
    col1, col2 = st.columns([1,1])
    with col1:
        fecha_evento = st.date_input("Fecha", value=date.today())
        lugar = st.text_input("Lugar", value="")
        estrategia = st.text_input("Estrategia o Programa", value="Estrategia Sembremos Seguridad")
    with col2:
        hora_inicio = st.time_input("Hora Inicio", value=time(9,0))
        hora_fin = st.time_input("Hora Finalización", value=time(12,10))
        # 🔹 SIEMPRE EN BLANCO (sin autollenado)
        delegacion_hdr = st.text_input("Dirección / Delegación Policial", value="")
        firmante_nombre = st.text_input("Nombre de quien firma (opcional)", value="")

    st.markdown("### 📝 Anotaciones y Acuerdos (para el Excel)")
    a_col, b_col = st.columns(2)
    anotaciones = a_col.text_area("Anotaciones Generales", height=220, placeholder="Escribe las anotaciones generales…")
    acuerdos    = b_col.text_area("Acuerdos", height=220, placeholder="Escribe los acuerdos…")

    st.markdown("### 👥 Registros y edición")
    if df_view.empty:
        st.info("No hay registros para el filtro seleccionado.")
    else:
        editable = df_view.copy()
        editable["Seleccionar"] = False

        edited = st.data_editor(
            editable[["Nº","Nombre","Cédula de Identidad","Delegación","Cargo","Teléfono",
                      "Género","Sexo","Rango de Edad","Seleccionar"]],
            hide_index=True,
            use_container_width=True,
            column_config={
                "Nº": st.column_config.NumberColumn("Nº", disabled=True),
                "Seleccionar": st.column_config.CheckboxColumn("Seleccionar", help="Marca para eliminar"),
                "Género": st.column_config.SelectboxColumn("Género", options=["F","M","LGBTIQ+"]),
                "Sexo": st.column_config.SelectboxColumn("Sexo", options=["H","M","I"]),
                "Rango de Edad": st.column_config.SelectboxColumn("Rango de Edad",
                    options=["18 a 35 años","36 a 64 años","65 años o más"])
            },
            key="tabla_admin_editable"
        )

        c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.2, 2])
        btn_save = c1.button("💾 Guardar cambios", use_container_width=True)
        btn_delete = c2.button("🗑️ Eliminar seleccionados", use_container_width=True)
        confirm_all = c4.checkbox("Confirmar vaciado total", value=False)
        btn_clear = c3.button("🧹 Vaciar todos", use_container_width=True)

        if btn_save:
            changes = 0
            for idx in edited.index:
                if idx >= len(df_view):
                    continue
                row_orig = df_view.loc[idx]
                row_new  = edited.loc[idx]
                fields = ["Nombre","Cédula de Identidad","Delegación","Cargo","Teléfono","Género","Sexo","Rango de Edad"]
                if any(str(row_orig[f]) != str(row_new[f]) for f in fields):
                    update_row_by_rownum(int(row_orig["rownum"]), {f: row_new[f] for f in fields})
                    changes += 1
            st.success(f"Se guardaron {changes} cambio(s).") if changes else st.info("No hay cambios para guardar.")
            if changes:
                st.rerun()

        if btn_delete:
            idx_sel = edited.index[edited["Seleccionar"] == True].tolist()
            rownums = df_view.iloc[idx_sel]["rownum"].tolist()
            if rownums:
                delete_rows_by_rownums(rownums)
                st.success(f"Eliminadas {len(rownums)} fila(s).")
                st.rerun()
            else:
                st.info("No hay filas seleccionadas para eliminar.")

        if btn_clear:

