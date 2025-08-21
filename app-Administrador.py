# =========================
# 🛠️ Panel del Administrador – Asistencia (editar, eliminar, Excel oficial)
# =========================
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date, datetime, time
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

# ---------- Tabs ----------
tab_form, tab_tabla, tab_excel = st.tabs(["📝 Formulario de Encabezado", "👥 Registros y edición", "⬇️ Excel oficial"])

# =========================
# 📝 1) Formulario de Encabezado (para el Excel oficial)
# =========================
with tab_form:
    st.caption("Estos datos rellenan los campos superiores de la plantilla oficial (con logos).")
    col1, col2 = st.columns([1,1])
    with col1:
        fecha_evento = st.date_input("Fecha", value=date.today())
        lugar = st.text_input("Lugar", value="sesión virtual")
        estrategia = st.text_input("Estrategia o Programa", value="Estrategia Sembremos Seguridad")
    with col2:
        hora_inicio = st.time_input("Hora Inicio", value=time(9,0))
        hora_fin = st.time_input("Hora Finalización", value=time(12,0))
        delegacion = st.text_input("Dirección / Delegación Policial", value="Naranjo")

    st.info("Cuando descargues el Excel, se aplicarán estos datos en la plantilla.")

# =========================
# 👥 2) Registros: ver, editar, eliminar
# =========================
with tab_tabla:
    st.markdown("### 👥 Registros recibidos")
    df_all = fetch_all_df(include_id=True)

    if df_all.empty:
        st.info("Aún no hay registros guardados.")
    else:
        # Editor editable (excepto Nº)
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
            # Detectar y aplicar cambios campo a campo
            changes = 0
            for idx in edited.index:
                if idx >= len(df_all):
                    continue
                row_orig = df_all.loc[idx]
                row_new = edited.loc[idx]
                # Si cambió algo (excepto Nº y Seleccionar)
                fields = ["Nombre","Cédula de Identidad","Institución","Cargo","Teléfono","Género","Sexo","Rango de Edad"]
                diff = any(str(row_orig[f]) != str(row_new[f]) for f in fields)
                if diff:
                    update_row_by_id(
                        int(row_orig["id"]),
                        {f: row_new[f] for f in fields}
                    )
                    changes += 1
            if changes:
                st.success(f"Se guardaron {changes} cambio(s).")
                st.rerun()
            else:
                st.info("No hay cambios para guardar.")

        if btn_delete:
            idx_sel = edited.index[edited["Seleccionar"] == True].tolist()
            ids = df_all.iloc[idx_sel]["id"].tolist()
            if ids:
                delete_rows_by_ids(ids)
                st.success(f"Eliminadas {len(ids)} fila(s).")
                st.rerun()
            else:
                st.info("No hay filas seleccionadas para eliminar.")

        if btn_clear:
            if confirm_all:
                delete_all_rows()
                st.success("Se vaciaron todos los registros.")
                st.rerun()
            else:
                st.warning("Marca la casilla 'Confirmar vaciado total' para continuar.")

# =========================
# ⬇️ 3) Excel oficial (plantilla con logos)
# =========================
with tab_excel:
    st.markdown("### ⬇️ Descargar Excel oficial con logos (plantilla)")
    st.caption("Se usa la plantilla que compartiste y se rellenan los encabezados y la tabla de asistencia.")
    TEMPLATE_PATH = "LA-2025-NARANJO.xlsx"   # coloca este archivo junto a este .py

    df_all = fetch_all_df(include_id=True)

    def build_excel_from_template(template_path: str,
                                  fecha: date,
                                  lugar: str,
                                  hora_ini: time,
                                  hora_fin: time,
                                  estrategia: str,
                                  delegacion: str,
                                  rows_df: pd.DataFrame) -> bytes:
        """
        Rellena la plantilla 'Minuta':
          - B6: 'Fecha: ...'
          - E6: 'Lugar: ...'
          - J6: 'Hora Inicio: ...'
          - Q6: 'Hora Finalización: ...'
          - B7: 'Estrategia o Programa: ...'
          - E8: delegación
        Tabla (desde fila 11):
          - B: Nº (ya viene impreso en plantilla)
          - C:E -> Nombre (merge por fila)
          - F -> Cédula
          - G -> Institución
          - H -> Cargo
          - I -> Teléfono
          - J/K/L -> Género (F/M/LGBTIQ+)
          - M/N/O -> Sexo (H/M/I)
          - P/Q/R -> Rango de Edad (18-35 / 36-64 / 65+)
          - S -> FIRMA (se deja como está en plantilla)
        Si hay más de 16 registros, duplica la hoja 'Minuta' en páginas.
        """
        try:
            from openpyxl import load_workbook
            from openpyxl.utils import get_column_letter
        except Exception:
            st.error("Falta 'openpyxl' en requirements.txt")
            return b""

        p = Path(template_path)
        if not p.exists():
            st.error(f"No se encontró la plantilla: {template_path}")
            return b""

        wb_master = load_workbook(p, data_only=False)
        # Hoja base
        if "Minuta" not in wb_master.sheetnames:
            st.error("La plantilla no contiene una hoja llamada 'Minuta'.")
            return b""

        # Cantidad de filas por página (según plantilla)
        # Detectamos merges Cxx:Exx para contar filas formateadas
        ws_base = wb_master["Minuta"]
        slots = []
        for r in ws_base.merged_cells.ranges:
            if r.min_col == 3 and r.max_col == 5 and r.min_row == r.max_row:
                slots.append(r.min_row)
        slots = sorted(slots)
        if not slots:
            st.error("No se detectaron filas de tabla en la plantilla (C:E merge por fila).")
            return b""
        start_row = slots[0]  # típicamente 11
        per_page = len(slots) # típicamente 16

        # Función para rellenar una hoja concreta con encabezados + una porción de datos
        def fill_sheet(ws, df_slice: pd.DataFrame):
            # Encabezados
            fecha_txt = fecha.strftime("%-d %B %Y") if hasattr(fecha, "strftime") else str(fecha)
            hora_ini_txt = hora_ini.strftime("%H:%M")
            hora_fin_txt = hora_fin.strftime("%H:%M")
            ws["B6"].value = f"Fecha: {fecha_txt}"
            ws["E6"].value = f"Lugar:  {lugar}"
            ws["J6"].value = f"Hora Inicio: {hora_ini_txt}"
            ws["Q6"].value = f"Hora Finalización: {hora_fin_txt}"
            ws["B7"].value = f"Estrategia o Programa: {estrategia}"
            ws["E8"].value = delegacion

            # Tabla
            # Mapas de columnas
            COL_NUM = "B"
            COL_NOMBRE_L = "C"  # C:E merged
            COL_NOMBRE_R = "E"
            COL_CED = "F"
            COL_INST = "G"
            COL_CARGO = "H"
            COL_TEL = "I"
            # Género
            COL_GEN_F = "J"; COL_GEN_M = "K"; COL_GEN_L = "L"
            # Sexo
            COL_SEX_H = "M"; COL_SEX_M = "N"; COL_SEX_I = "O"
            # Edad
            COL_ED_1 = "P"; COL_ED_2 = "Q"; COL_ED_3 = "R"

            for i, (_, row) in enumerate(df_slice.iterrows()):
                r = start_row + i
                # Nombre (en merge C:E) -> escribimos en la izquierda del merge
                ws[f"{COL_NOMBRE_L}{r}"].value = str(row["Nombre"]) if pd.notna(row["Nombre"]) else ""
                # Ced, Inst, Cargo, Tel
                ws[f"{COL_CED}{r}"].value = str(row["Cédula de Identidad"] or "")
                ws[f"{COL_INST}{r}"].value = str(row["Institución"] or "")
                ws[f"{COL_CARGO}{r}"].value = str(row["Cargo"] or "")
                ws[f"{COL_TEL}{r}"].value = str(row["Teléfono"] or "")
                # Limpiar marcas previas
                for c in [COL_GEN_F, COL_GEN_M, COL_GEN_L, COL_SEX_H, COL_SEX_M, COL_SEX_I, COL_ED_1, COL_ED_2, COL_ED_3]:
                    ws[f"{c}{r}"].value = ""
                # Marcas
                g = (row["Género"] or "").strip()
                if g == "F": ws[f"{COL_GEN_F}{r}"].value = "X"
                elif g == "M": ws[f"{COL_GEN_M}{r}"].value = "X"
                elif g == "LGBTIQ+": ws[f"{COL_GEN_L}{r}"].value = "X"

                s = (row["Sexo"] or "").strip()
                if s == "H": ws[f"{COL_SEX_H}{r}"].value = "X"
                elif s == "M": ws[f"{COL_SEX_M}{r}"].value = "X"
                elif s == "I": ws[f"{COL_SEX_I}{r}"].value = "X"

                e = (row["Rango de Edad"] or "").strip()
                if e.startswith("18"): ws[f"{COL_ED_1}{r}"].value = "X"
                elif e.startswith("36"): ws[f"{COL_ED_2}{r}"].value = "X"
                elif e.startswith("65"): ws[f"{COL_ED_3}{r}"].value = "X"

        # Paginado en copias de la hoja
        total = len(rows_df)
        if total == 0:
            wb_out = wb_master
            bio = BytesIO(); wb_out.save(bio); return bio.getvalue()

        # Creamos un nuevo libro para no modificar el master original en memoria
        from openpyxl import Workbook
        wb_out = Workbook()
        # Quitar hoja por defecto
        default_ws = wb_out.active
        wb_out.remove(default_ws)

        pages = (total + per_page - 1) // per_page
        for p in range(pages):
            start = p * per_page
            end = min(start + per_page, total)
            df_slice = rows_df.iloc[start:end].reset_index(drop=True)
            # Copiar hoja "Minuta"
            ws_copy = wb_master.copy_worksheet(wb_master["Minuta"])
            ws_copy.title = "Minuta" if p == 0 else f"Minuta {p+1}"
            # Rellenar
            fill_sheet(ws_copy, df_slice)
            # Agregar la hoja rellena al libro de salida
            wb_out._add_sheet(ws_copy)  # uso interno; mantiene estilos/merges

        # El master quedará con la original + copias; pero wb_out ya tiene las rellenas.
        bio = BytesIO(); wb_out.save(bio); return bio.getvalue()

    # Botón de descarga (usa plantilla)
    if st.button("📥 Generar y descargar Excel oficial", use_container_width=True, type="primary"):
        xls_bytes = build_excel_from_template(
            TEMPLATE_PATH,
            fecha_evento,
            lugar,
            hora_inicio,
            hora_fin,
            estrategia,
            delegacion,
            df_all[["Nombre","Cédula de Identidad","Institución","Cargo","Teléfono","Género","Sexo","Rango de Edad"]] if not df_all.empty else df_all
        )
        if xls_bytes:
            st.download_button(
                "⬇️ Descargar Excel (plantilla oficial con logos)",
                data=xls_bytes,
                file_name=f"Lista_Asistencia_Oficial_{date.today():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )






