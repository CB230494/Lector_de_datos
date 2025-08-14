# =========================
# 📋 Lista de asistencia – Seguimiento de líneas de acción (multiusuario con SQLite)
# =========================
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date
import sqlite3

st.markdown("## 📋 Lista de asistencia – Seguimiento de líneas de acción")

# ---------- DB (SQLite persistente) ----------
DB_PATH = "asistencia.db"

def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False, timeout=30)

def init_db():
    with get_conn() as conn:
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA synchronous=NORMAL;")
        conn.execute("""
        CREATE TABLE IF NOT EXISTS asistencia(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL DEFAULT (datetime('now','localtime')),
            nombre TEXT, cedula TEXT, institucion TEXT,
            cargo TEXT, telefono TEXT,
            genero TEXT, sexo TEXT, edad TEXT
        );
        """)

def insert_row(row):
    with get_conn() as conn:
        conn.execute("""INSERT INTO asistencia
            (nombre, cedula, institucion, cargo, telefono, genero, sexo, edad)
            VALUES (?,?,?,?,?,?,?,?)""",
            (row["Nombre"], row["Cédula de Identidad"], row["Institución"],
             row["Cargo"], row["Teléfono"], row["Género"], row["Sexo"], row["Rango de Edad"])
        )

def fetch_all_df():
    with get_conn() as conn:
        df = pd.read_sql_query("""
            SELECT id, created_at,
                   nombre  AS 'Nombre',
                   cedula  AS 'Cédula de Identidad',
                   institucion AS 'Institución',
                   cargo   AS 'Cargo',
                   telefono AS 'Teléfono',
                   genero  AS 'Género',
                   sexo    AS 'Sexo',
                   edad    AS 'Rango de Edad'
            FROM asistencia
            ORDER BY id ASC
        """, conn)
    if not df.empty:
        df.insert(0, "Nº", range(1, len(df)+1))
    return df

init_db()

# ---------- Formulario sencillo ----------
with st.form("form_asistencia", clear_on_submit=True):
    c1, c2, c3 = st.columns([1.2, 1, 1])
    nombre      = c1.text_input("Nombre")
    cedula      = c2.text_input("Cédula de Identidad")
    institucion = c3.text_input("Institución")

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
                "Institución": institucion.strip(),
                "Cargo": cargo.strip(),
                "Teléfono": telefono.strip(),
                "Género": genero,
                "Sexo": sexo,
                "Rango de Edad": edad
            }
            insert_row(fila)
            st.success("Registro guardado.")

# ---------- Vista de todo lo recibido ----------
st.markdown("### 📥 Registros recibidos")
df_all = fetch_all_df()
st.dataframe(df_all if not df_all.empty else pd.DataFrame(
    columns=["Nº","Nombre","Cédula de Identidad","Institución","Cargo","Teléfono","Género","Sexo","Rango de Edad"]),
    hide_index=True, use_container_width=True
)

# ---------- Generar Excel con el mismo formato de tu plantilla ----------
def build_excel_asistencia(rows_df: pd.DataFrame) -> bytes:
    if rows_df.empty:
        return b""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    except Exception:
        st.error("Falta 'openpyxl' en requirements.txt")
        return b""

    wb = Workbook(); ws = wb.active; ws.title = "Asistencia"
    widths = [5, 28, 18, 24, 20, 16, 12, 12, 12, 12, 12, 12, 14, 14, 14]
    for i, w in enumerate(widths, start=1): ws.column_dimensions[chr(64+i)].width = w

    title_fill = PatternFill("solid", fgColor="1F3B73"); title_font = Font(bold=True, size=14, color="FFFFFF")
    head_fill  = PatternFill("solid", fgColor="DDE7FF"); group_fill = PatternFill("solid", fgColor="B7C6F9")
    head_font  = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin = Side(style="thin", color="000000"); border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Título
    ws.merge_cells("A1:O1")
    c = ws["A1"]; c.value = "Lista de asistencia – Seguimiento de líneas de acción"
    c.fill, c.font, c.alignment = title_fill, title_font, center

    # Encabezados (A2:O3)
    for rng, text in [("A2:A3","Nº"),("B2:B3","Nombre"),("C2:C3","Cédula de Identidad"),
                      ("D2:D3","Institución"),("E2:E3","Cargo"),("F2:F3","Teléfono"),
                      ("G2:I2","Género"),("J2:L2","Sexo (Hombre, Mujer o Intersex)"),("M2:O2","Rango de Edad")]:
        ws.merge_cells(rng); cell = ws[rng.split(":")[0]]
        cell.value = text; cell.alignment = center; cell.font = head_font
        cell.fill = group_fill if rng in ["G2:I2","J2:L2","M2:O2"] else head_fill
    for addr, txt in {"G3":"F","H3":"M","I3":"LGBTIQ+","J3":"H","K3":"M","L3":"I",
                      "M3":"18 a 35 años","N3":"36 a 64 años","O3":"65 años o más"}.items():
        cell = ws[addr]; cell.value = txt; cell.font = head_font; cell.alignment = center; cell.fill = head_fill
    for r in range(2,4):
        for cidx in range(1,16): ws.cell(row=r, column=cidx).border = border_all

    # Cuerpo
    start = 4
    for i, row in rows_df.iterrows():
        rr = start + i
        vals = [row["Nº"], row["Nombre"], row["Cédula de Identidad"], row["Institución"], row["Cargo"], row["Teléfono"]]
        for cidx, v in enumerate(vals, start=1):
            cell = ws.cell(row=rr, column=cidx, value=v); cell.border = border_all
            cell.alignment = center if cidx == 1 else left
        # Marcar X según columnas de texto
        g, s, e = row["Género"], row["Sexo"], row["Rango de Edad"]
        marks = ["X" if g=="F" else "", "X" if g=="M" else "", "X" if g=="LGBTIQ+" else "",
                 "X" if s=="H" else "", "X" if s=="M" else "", "X" if s=="I" else "",
                 "X" if e.startswith("18") else "", "X" if e.startswith("36") else "", "X" if e.startswith("65") else ""]
        for off, v in enumerate(marks, start=7):
            cell = ws.cell(row=rr, column=off, value=v); cell.border = border_all; cell.alignment = center

    ws.freeze_panes = "B4"
    bio = BytesIO(); wb.save(bio); return bio.getvalue()

# Botón de descarga (todos los registros)
if not df_all.empty:
    xls_bytes = build_excel_asistencia(df_all[[
        "Nº","Nombre","Cédula de Identidad","Institución","Cargo","Teléfono","Género","Sexo","Rango de Edad"
    ]])
    st.download_button(
        "⬇️ Descargar Excel (todos los registros)",
        data=xls_bytes,
        file_name=f"Lista_Asistencia_LineasAccion_{date.today():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
else:
    st.info("Aún no hay registros guardados.")






