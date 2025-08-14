# =========================
# 📋 Lista de asistencia – Seguimiento de líneas de acción (Excel con estilos)
# =========================
import streamlit as st
from io import BytesIO
from datetime import date

st.markdown("## 📋 Lista de asistencia – Seguimiento de líneas de acción")

filas = st.number_input("Cantidad de filas a generar", min_value=5, max_value=500, value=30, step=5)

def build_excel_asistencia(n_rows: int):
    try:
        # Import dinámico para no romper si aún no está instalado
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
        from openpyxl.utils import get_column_letter
    except Exception as e:
        st.error("Falta el paquete 'openpyxl'. Agrégalo en requirements.txt (línea: openpyxl) y vuelve a ejecutar.")
        return None

    wb = Workbook()
    ws = wb.active
    ws.title = "Asistencia"

    # --- Cols A:O ---
    col_widths = {
        "A": 5, "B": 28, "C": 18, "D": 24, "E": 20, "F": 16,
        "G": 12, "H": 12, "I": 12, "J": 12, "K": 12, "L": 12, "M": 14, "N": 14, "O": 14
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    # --- Estilos ---
    title_fill  = PatternFill("solid", fgColor="1F3B73")  # azul
    title_font  = Font(bold=True, size=14, color="FFFFFF")
    head_fill   = PatternFill("solid", fgColor="DDE7FF")
    group_fill  = PatternFill("solid", fgColor="B7C6F9")
    head_font   = Font(bold=True)
    center      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left        = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="000000")
    border_all  = Border(left=thin, right=thin, top=thin, bottom=thin)

    # --- Título ---
    ws.merge_cells("A1:O1")
    c = ws["A1"]
    c.value = "Lista de asistencia – Seguimiento de líneas de acción"
    c.fill, c.font, c.alignment = title_fill, title_font, center
    ws.row_dimensions[1].height = 28

    # --- Encabezados (dos filas) ---
    # Fila 2 grupos / celdas combinadas
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

    # Fila 3 subcolumnas
    heads_row3 = {
        "G3":"F", "H3":"M", "I3":"LGBTIQ+",
        "J3":"H", "K3":"M", "L3":"I",
        "M3":"18 a 35 años", "N3":"36 a 64 años", "O3":"65 años o más"
    }
    for addr, text in heads_row3.items():
        cell = ws[addr]
        cell.value = text
        cell.font = head_font
        cell.alignment = center
        cell.fill = head_fill

    ws.row_dimensions[2].height = 24
    ws.row_dimensions[3].height = 24

    # Bordes en encabezados A2:O3
    for r in range(2, 4):
        for cidx in range(1, 16):
            ws.cell(row=r, column=cidx).border = border_all

    # --- Cuerpo (n_rows) ---
    start_row = 4
    for i in range(n_rows):
        r = start_row + i
        # Nº
        cell = ws.cell(row=r, column=1, value=i+1)
        cell.alignment = center
        cell.border = border_all
        # Campos texto
        for cidx in [2,3,4,5,6]:
            cell = ws.cell(row=r, column=cidx, value=None)
            cell.alignment = left if cidx != 6 else left
            cell.border = border_all
        # Marcas (G..O)
        for cidx in range(7, 16):
            cell = ws.cell(row=r, column=cidx, value=None)
            cell.alignment = center
            cell.border = border_all
        ws.row_dimensions[r].height = 20

    # Congelar encabezados (hasta fila 3 y columna A)
    ws.freeze_panes = "B4"

    # Guardar a bytes
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()

excel_bytes = build_excel_asistencia(int(filas))
if excel_bytes:
    st.download_button(
        "⬇️ Descargar Excel",
        data=excel_bytes,
        file_name=f"Lista_Asistencia_LineasAccion_{date.today():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Descarga la lista con encabezados combinados, bordes y colores."
    )



