# =========================
# 📋 Lista de asistencia – Seguimiento de líneas de acción
# =========================
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date

st.markdown("## 📋 Lista de asistencia – Seguimiento de líneas de acción")

# Estado
if "asistencia_rows" not in st.session_state:
    st.session_state.asistencia_rows = []

# ------- Formulario simple -------
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

    submitted = st.form_submit_button("➕ Agregar a la lista", use_container_width=True)
    if submitted:
        # Validación mínima
        if not nombre.strip():
            st.warning("Ingresa al menos el nombre.")
        else:
            st.session_state.asistencia_rows.append({
                "Nº": len(st.session_state.asistencia_rows) + 1,
                "Nombre": nombre.strip(),
                "Cédula de Identidad": cedula.strip(),
                "Institución": institucion.strip(),
                "Cargo": cargo.strip(),
                "Teléfono": telefono.strip(),
                "Género": genero,
                "Sexo": sexo,
                "Rango de Edad": edad
            })
            st.success("Registro agregado.")

# ------- Cuadro con lo ingresado -------
if st.session_state.asistencia_rows:
    df_vis = pd.DataFrame(st.session_state.asistencia_rows, columns=[
        "Nº","Nombre","Cédula de Identidad","Institución","Cargo","Teléfono",
        "Género","Sexo","Rango de Edad"
    ])
    st.markdown("### Registros cargados")
    st.dataframe(df_vis, hide_index=True, use_container_width=True)

    cbtn1, cbtn2, _ = st.columns([1,1,6])
    if cbtn1.button("🗑️ Eliminar última fila"):
        st.session_state.asistencia_rows.pop()
        # Reenumerar
        for i, r in enumerate(st.session_state.asistencia_rows, start=1):
            r["Nº"] = i
    if cbtn2.button("🧹 Vaciar lista"):
        st.session_state.asistencia_rows.clear()

# ------- Excel con encabezados combinados y colores (igual a la plantilla) -------
def build_excel_asistencia(rows: list) -> bytes:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    except Exception:
        st.error("Falta el paquete 'openpyxl'. Agrega `openpyxl` a requirements.txt.")
        return b""

    wb = Workbook()
    ws = wb.active
    ws.title = "Asistencia"

    # Anchos de columnas A:O
    widths = [5, 28, 18, 24, 20, 16, 12, 12, 12, 12, 12, 12, 14, 14, 14]
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64+idx)].width = w

    # Estilos
    title_fill = PatternFill("solid", fgColor="1F3B73")
    title_font = Font(bold=True, size=14, color="FFFFFF")
    head_fill  = PatternFill("solid", fgColor="DDE7FF")
    group_fill = PatternFill("solid", fgColor="B7C6F9")
    head_font  = Font(bold=True)
    center     = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left       = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin       = Side(style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Título
    ws.merge_cells("A1:O1")
    c = ws["A1"]; c.value = "Lista de asistencia – Seguimiento de líneas de acción"
    c.fill, c.font, c.alignment = title_fill, title_font, center
    ws.row_dimensions[1].height = 28

    # Encabezados (dos filas)
    merges = [
        ("A2:A3","Nº"), ("B2:B3","Nombre"), ("C2:C3","Cédula de Identidad"),
        ("D2:D3","Institución"), ("E2:E3","Cargo"), ("F2:F3","Teléfono"),
        ("G2:I2","Género"), ("J2:L2","Sexo (Hombre, Mujer o Intersex)"), ("M2:O2","Rango de Edad")
    ]
    for rng, text in merges:
        ws.merge_cells(rng)
        cell = ws[rng.split(":")[0]]
        cell.value = text
        cell.alignment = center
        cell.font = head_font
        cell.fill = group_fill if rng in ["G2:I2","J2:L2","M2:O2"] else head_fill

    subs = {
        "G3":"F", "H3":"M", "I3":"LGBTIQ+",
        "J3":"H", "K3":"M", "L3":"I",
        "M3":"18 a 35 años", "N3":"36 a 64 años", "O3":"65 años o más"
    }
    for addr, text in subs.items():
        cell = ws[addr]; cell.value = text
        cell.font = head_font; cell.alignment = center; cell.fill = head_fill

    ws.row_dimensions[2].height = 24
    ws.row_dimensions[3].height = 24
    for r in range(2, 4):
        for cidx in range(1, 16):
            ws.cell(row=r, column=cidx).border = border_all

    # Cuerpo
    start_row = 4
    for i, r in enumerate(rows, start=0):
        rr = start_row + i
        # Campos texto
        values = [
            r["Nº"], r["Nombre"], r["Cédula de Identidad"], r["Institución"], r["Cargo"], r["Teléfono"]
        ]
        for cidx, v in enumerate(values, start=1):
            cell = ws.cell(row=rr, column=cidx, value=v)
            cell.border = border_all
            cell.alignment = center if cidx == 1 else left

        # Marcas X según selección
        g = r["Género"]; s = r["Sexo"]; e = r["Rango de Edad"]
        marks = [
            "X" if g=="F" else "", "X" if g=="M" else "", "X" if g=="LGBTIQ+" else "",
            "X" if s=="H" else "", "X" if s=="M" else "", "X" if s=="I" else "",
            "X" if e.startswith("18") else "", "X" if e.startswith("36") else "", "X" if e.startswith("65") else ""
        ]
        for off, v in enumerate(marks, start=7):  # columnas G..O
            cell = ws.cell(row=rr, column=off, value=v)
            cell.border = border_all
            cell.alignment = center

        ws.row_dimensions[rr].height = 20

    ws.freeze_panes = "B4"
    bio = BytesIO(); wb.save(bio)
    return bio.getvalue()

# Botón de descarga
if st.session_state.asistencia_rows:
    excel_bytes = build_excel_asistencia(st.session_state.asistencia_rows)
    st.download_button(
        "⬇️ Descargar Excel",
        data=excel_bytes,
        file_name=f"Lista_Asistencia_LineasAccion_{date.today():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
else:
    st.info("Agrega registros al formulario y aquí podrás descargar el Excel.")





