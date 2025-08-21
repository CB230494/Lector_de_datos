# =========================
# 🛠️ Panel del Administrador – Asistencia (editar, eliminar, Excel oficial)
# =========================
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date, time
import sqlite3
from pathlib import Path

st.set_page_config(page_title="Asistencia - Administrador", layout="wide")
st.markdown("# 🛠️ Panel del Administrador – Asistencia")

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

def fetch_all_df(include_id=True):
    with get_conn() as conn:
        df = pd.read_sql_query("""
            SELECT id,
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
    if not include_id and not df.empty:
        return df.drop(columns=["id"])
    return df

def insert_row(row):
    with get_conn() as conn:
        conn.execute("""INSERT INTO asistencia
            (nombre, cedula, institucion, cargo, telefono, genero, sexo, edad)
            VALUES (?,?,?,?,?,?,?,?)""",
            (row["Nombre"], row["Cédula de Identidad"], row["Institución"],
             row["Cargo"], row["Teléfono"], row["Género"], row["Sexo"], row["Rango de Edad"])
        )

def update_row_by_id(row_id:int, row:dict):
    with get_conn() as conn:
        conn.execute("""
            UPDATE asistencia
               SET nombre=?, cedula=?, institucion=?, cargo=?, telefono=?, genero=?, sexo=?, edad=?
             WHERE id=?""",
            (row["Nombre"], row["Cédula de Identidad"], row["Institución"],
             row["Cargo"], row["Teléfono"], row["Género"], row["Sexo"], row["Rango de Edad"], row_id)
        )

def delete_rows_by_ids(ids):
    if not ids:
        return
    with get_conn() as conn:
        q = ",".join("?" for _ in ids)
        conn.execute(f"DELETE FROM asistencia WHERE id IN ({q})", ids)

def delete_all_rows():
    with get_conn() as conn:
        conn.execute("DELETE FROM asistencia;")

init_db()

# ---------- Tabs (¡esto crea tab_form, tab_tabla y tab_excel!) ----------
tab_form, tab_tabla, tab_excel = st.tabs(
    ["📝 Formulario de Encabezado", "👥 Registros y edición", "⬇️ Excel oficial"]
)

# =========================
# 📝 1) Formulario de Encabezado (para el Excel)
# =========================
with tab_form:
    st.caption("Estos datos rellenan los campos superiores del Excel que se genera.")
    col1, col2 = st.columns([1,1])
    with col1:
        fecha_evento = st.date_input("Fecha", value=date.today())
        lugar = st.text_input("Lugar", value="sesión virtual")
        estrategia = st.text_input("Estrategia o Programa", value="Estrategia Sembremos Seguridad")
    with col2:
        hora_inicio = st.time_input("Hora Inicio", value=time(9,0))
        hora_fin = st.time_input("Hora Finalización", value=time(12,0))
        delegacion = st.text_input("Dirección / Delegación Policial", value="Naranjo")

# =========================
# 👥 2) Registros: ver, editar, eliminar
# =========================
with tab_tabla:
    st.markdown("### 👥 Registros recibidos")
    df_all = fetch_all_df(include_id=True)

    if df_all.empty:
        st.info("Aún no hay registros guardados.")
    else:
        editable = df_all.copy()
        editable["Seleccionar"] = False

        edited = st.data_editor(
            editable[["Nº","Nombre","Cédula de Identidad","Institución","Cargo","Teléfono",
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
                if idx >= len(df_all):
                    continue
                row_orig = df_all.loc[idx]
                row_new = edited.loc[idx]
                fields = ["Nombre","Cédula de Identidad","Institución","Cargo","Teléfono","Género","Sexo","Rango de Edad"]
                if any(str(row_orig[f]) != str(row_new[f]) for f in fields):
                    update_row_by_id(int(row_orig["id"]), {f: row_new[f] for f in fields})
                    changes += 1
            st.success(f"Se guardaron {changes} cambio(s).") if changes else st.info("No hay cambios para guardar.")
            if changes: st.rerun()

        if btn_delete:
            idx_sel = edited.index[edited["Seleccionar"] == True].tolist()
            ids = df_all.iloc[idx_sel]["id"].tolist()
            if ids:
                delete_rows_by_ids(ids)
                st.success(f"Eliminadas {len(ids)} fila(s)."); st.rerun()
            else:
                st.info("No hay filas seleccionadas para eliminar.")

        if btn_clear:
            if confirm_all:
                delete_all_rows(); st.success("Se vaciaron todos los registros."); st.rerun()
            else:
                st.warning("Marca 'Confirmar vaciado total' para continuar.")

# =========================
# ⬇️ 3) Excel oficial (estructura replicada; SIN plantilla)
# =========================
with tab_excel:
    st.markdown("### ⬇️ Descargar Excel oficial (estructura replicada)")
    st.caption("Se recrea la minuta desde cero. Si agregas 'logo_izq.png' y/o 'logo_der.png' en la carpeta, se insertan arriba.")

    df_all = fetch_all_df(include_id=True)

    def _fecha_es(fecha: date) -> str:
        meses = ["enero","febrero","marzo","abril","mayo","junio",
                 "julio","agosto","septiembre","octubre","noviembre","diciembre"]
        return f"{fecha.day} {meses[fecha.month-1]} {fecha.year}"

    def build_excel_official_from_scratch(
        fecha: date,
        lugar: str,
        hora_ini: time,
        hora_fin: time,
        estrategia: str,
        delegacion: str,
        rows_df: pd.DataFrame,
        per_page: int = 16
    ) -> bytes:
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
            from openpyxl.drawing.image import Image as XLImage
        except Exception:
            st.error("Falta 'openpyxl' en requirements.txt")
            return b""

        head_fill  = PatternFill("solid", fgColor="DDE7FF")
        group_fill = PatternFill("solid", fgColor="B7C6F9")
        head_font  = Font(bold=True)
        title_font = Font(bold=True, size=12)
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
        thin = Side(style="thin", color="000000")
        border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

        wb = Workbook()
        ws0 = wb.active
        ws0.title = "Minuta"

        def _setup_sheet(ws):
            widths = {
                "A": 2, "B": 6, "C": 22, "D": 22, "E": 22, "F": 18, "G": 22,
                "H": 20, "I": 16, "J": 6, "K": 6, "L": 10, "M": 6, "N": 6, "O": 6,
                "P": 14, "Q": 14, "R": 14, "S": 16
            }
            for col, w in widths.items(): ws.column_dimensions[col].width = w

            # logos opcionales
            try:
                if Path("logo_izq.png").exists():
                    img = XLImage("logo_izq.png"); img.width *= 0.6; img.height *= 0.6
                    ws.add_image(img, "B2")
                if Path("logo_der.png").exists():
                    img2 = XLImage("logo_der.png"); img2.width *= 0.6; img2.height *= 0.6
                    ws.add_image(img2, "Q2")
            except Exception:
                pass

            ws["B6"].value = f"Fecha: {_fecha_es(fecha)}"; ws["B6"].font = title_font
            ws["E6"].value = f"Lugar:  {lugar}"; ws.merge_cells("E6:I6"); ws["E6"].font = title_font
            ws["J6"].value = f"Hora Inicio: {hora_ini.strftime('%H:%M')}"; ws["J6"].font = title_font
            ws["Q6"].value = f"Hora Finalización: {hora_fin.strftime('%H:%M')}"; ws["Q6"].font = title_font

            ws["B7"].value = f"Estrategia o Programa: {estrategia}"
            ws["B7"].font = title_font; ws.merge_cells("B7:G7")
            ws["H7"].value = "AC... acción, acciones estratégicas, indicadores y metas."
            ws.merge_cells("H7:S8")

            ws["B8"].value = "Dirección / Delegación Policial:"
            ws["E8"].value = delegacion

            # cabeceras de tabla
            ws.merge_cells("B9:E10"); ws["B9"].value = "Nombre"
            ws["F9"].value = "Cédula de Identidad"
            ws["G9"].value = "Institución"
            ws["H9"].value = "Cargo"
            ws["I9"].value = "Teléfono"
            ws.merge_cells("J9:L9"); ws["J9"].value = "Género"
            ws.merge_cells("M9:O9"); ws["M9"].value = "Sexo (Hombre, Mujer o Intersex)"
            ws.merge_cells("P9:R9"); ws["P9"].value = "Rango de Edad"
            ws["S9"].value = "FIRMA"

            for rng in ["B9:E10","J9:L9","M9:O9","P9:R9"]:
                c = ws[rng.split(":")[0]]; c.font = head_font; c.alignment = center; c.fill = group_fill
            for cell in ["F9","G9","H9","I9","S9"]:
                ws[cell].font = head_font; ws[cell].alignment = center; ws[cell].fill = head_fill

            ws["J10"].value, ws["K10"].value, ws["L10"].value = "F", "M", "LGBTIQ+"
            ws["M10"].value, ws["N10"].value, ws["O10"].value = "H", "M", "I"
            ws["P10"].value, ws["Q10"].value, ws["R10"].value = "18 a 35 años", "36 a 64 años", "65 años o más"
            for cell in ["J10","K10","L10","M10","N10","O10","P10","Q10","R10"]:
                ws[cell].font = head_font; ws[cell].alignment = center; ws[cell].fill = head_fill

            for r in range(9, 11):
                for c in range(2, 20): ws.cell(row=r, column=c).border = border_all
            ws.freeze_panes = "C11"

        def _fill_rows(ws, df_slice: pd.DataFrame, start_row: int = 11):
            for i, (_, row) in enumerate(df_slice.iterrows()):
                r = start_row + i
                ws[f"B{r}"].value = i + 1; ws[f"B{r}"].alignment = Alignment(horizontal="center", vertical="center")

                ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)  # C:E
                ws[f"C{r}"].value = str(row["Nombre"] or ""); ws[f"C{r}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

                ws[f"F{r}"].value = str(row["Cédula de Identidad"] or "")
                ws[f"G{r}"].value = str(row["Institución"] or "")
                ws[f"H{r}"].value = str(row["Cargo"] or "")
                ws[f"I{r}"].value = str(row["Teléfono"] or "")

                for col in ["J","K","L","M","N","O","P","Q","R"]: ws[f"{col}{r}"].value = ""

                g = (row["Género"] or "").strip()
                if g == "F": ws[f"J{r}"].value = "X"
                elif g == "M": ws[f"K{r}"].value = "X"
                elif g == "LGBTIQ+": ws[f"L{r}"].value = "X"

                s = (row["Sexo"] or "").strip()
                if s == "H": ws[f"M{r}"].value = "X"
                elif s == "M": ws[f"N{r}"].value = "X"
                elif s == "I": ws[f"O{r}"].value = "X"

                e = (row["Rango de Edad"] or "").strip()
                if e.startswith("18"): ws[f"P{r}"].value = "X"
                elif e.startswith("36"): ws[f"Q{r}"].value = "X"
                elif e.startswith("65"): ws[f"R{r}"].value = "X"

                ws[f"S{r}"].value = "Virtual"
                for c in range(2, 20): ws.cell(row=r, column=c).border = border_all

        total = len(rows_df)
        pages = max(1, (total + per_page - 1) // per_page) if total > 0 else 1

        for p in range(pages):
            ws = wb["Minuta"] if p == 0 else wb.copy_worksheet(wb["Minuta"])
            if p > 0: ws.title = f"Minuta {p+1}"
            _setup_sheet(ws)
            start = p * per_page
            end = min(start + per_page, total)
            df_slice = rows_df.iloc[start:end].reset_index(drop=True) if total > 0 else rows_df.head(0)
            _fill_rows(ws, df_slice)

        bio = BytesIO(); wb.save(bio); return bio.getvalue()

    if st.button("📥 Generar y descargar Excel oficial", use_container_width=True, type="primary"):
        datos = df_all[["Nombre","Cédula de Identidad","Institución","Cargo","Teléfono","Género","Sexo","Rango de Edad"]] if not df_all.empty else df_all
        xls_bytes = build_excel_official_from_scratch(
            fecha_evento, lugar, hora_inicio, hora_fin, estrategia, delegacion, datos
        )
        if xls_bytes:
            st.download_button(
                "⬇️ Descargar Excel (estructura replicada)",
                data=xls_bytes,
                file_name=f"Lista_Asistencia_Oficial_{date.today():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )







