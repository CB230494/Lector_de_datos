# =========================
# 📋 Lista de asistencia – Seguimiento de líneas de acción
# =========================
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date

st.markdown("## 📋 Lista de asistencia – Seguimiento de líneas de acción")

# --- Número de filas ---
filas = st.number_input("Cantidad de filas a generar", min_value=5, max_value=500, value=25, step=5)

# --- Definición de columnas (orden igual a la plantilla) ---
COLS = [
    "Nº", "Nombre", "Cédula de Identidad", "Institución", "Cargo", "Teléfono",
    "Genero_F", "Genero_M", "Genero_LGBTIQ+",
    "Sexo_H", "Sexo_M", "Sexo_I",
    "Edad_18_35", "Edad_36_64", "Edad_65_mas"
]

def make_empty_df(n):
    data = []
    for i in range(1, n+1):
        data.append({
            "Nº": i, "Nombre": "", "Cédula de Identidad": "", "Institución": "", "Cargo": "", "Teléfono": "",
            "Genero_F": False, "Genero_M": False, "Genero_LGBTIQ+": False,
            "Sexo_H": False, "Sexo_M": False, "Sexo_I": False,
            "Edad_18_35": False, "Edad_36_64": False, "Edad_65_mas": False
        })
    return pd.DataFrame(data, columns=COLS)

df_init = make_empty_df(int(filas))

# --- Cabecera "visual" para que se vea como la imagen ---
st.markdown(
    """
    <div style="display:grid;grid-template-columns:50px 1.4fr 1fr 1.2fr .9fr .9fr .5fr .5fr .8fr .5fr .5fr .5fr .9fr .9fr .9fr;gap:2px;font-weight:600;text-align:center;">
      <div></div><div></div><div></div><div></div><div></div><div></div>
      <div style="grid-column:7/10;background:#B7C6F9;border-radius:4px;padding:6px 0;">Género</div>
      <div style="grid-column:10/13;background:#B7C6F9;border-radius:4px;padding:6px 0;">Sexo (Hombre, Mujer o Intersex)</div>
      <div style="grid-column:13/16;background:#B7C6F9;border-radius:4px;padding:6px 0;">Rango de Edad</div>
    </div>
    """,
    unsafe_allow_html=True
)

# --- Editor de datos (formulario) ---
edited_df = st.data_editor(
    df_init,
    hide_index=True,
    use_container_width=True,
    column_config={
        "Nº": st.column_config.NumberColumn("Nº", disabled=True, width="small"),
        "Nombre": st.column_config.TextColumn("Nombre", width="large"),
        "Cédula de Identidad": st.column_config.TextColumn("Cédula de Identidad", width="medium"),
        "Institución": st.column_config.TextColumn("Institución", width="medium"),
        "Cargo": st.column_config.TextColumn("Cargo", width="medium"),
        "Teléfono": st.column_config.TextColumn("Teléfono", width="medium"),

        # Género
        "Genero_F": st.column_config.CheckboxColumn("F", help="Género: Femenino", width="small"),
        "Genero_M": st.column_config.CheckboxColumn("M", help="Género: Masculino", width="small"),
        "Genero_LGBTIQ+": st.column_config.CheckboxColumn("LGBTIQ+", help="Género: LGBTIQ+", width="small"),

        # Sexo
        "Sexo_H": st.column_config.CheckboxColumn("H", help="Sexo: Hombre", width="small"),
        "Sexo_M": st.column_config.CheckboxColumn("M", help="Sexo: Mujer", width="small"),
        "Sexo_I": st.column_config.CheckboxColumn("I", help="Sexo: Intersex", width="small"),

        # Rango de edad
        "Edad_18_35": st.column_config.CheckboxColumn("18 a 35 años", width="medium"),
        "Edad_36_64": st.column_config.CheckboxColumn("36 a 64 años", width="medium"),
        "Edad_65_mas": st.column_config.CheckboxColumn("65 años o más", width="medium"),
    }
)

# --- Generar Excel con encabezados combinados y colores, usando openpyxl ---
def build_excel_asistencia(df: pd.DataFrame) -> bytes:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    except Exception:
        st.error("Falta el paquete 'openpyxl'. Agrega `openpyxl` a requirements.txt y vuelve a ejecutar.")
        return b""

    wb = Workbook()
    ws = wb.active
    ws.title = "Asistencia"

    # Anchos
    widths = [5, 28, 18, 24, 20, 16, 12, 12, 12, 12, 12, 12, 14, 14, 14]
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64+idx)].width = w

    # Estilos
    title_fill  = PatternFill("solid", fgColor="1F3B73")
    title_font  = Font(bold=True, size=14, color="FFFFFF")
    head_fill   = PatternFill("solid", fgColor="DDE7FF")
    group_fill  = PatternFill("solid", fgColor="B7C6F9")
    head_font   = Font(bold=True)
    center      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left        = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin        = Side(style="thin", color="000000")
    border_all  = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Título
    ws.merge_cells("A1:O1")
    c = ws["A1"]
    c.value = "Lista de asistencia – Seguimiento de líneas de acción"
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
        cell = ws[addr]
        cell.value = text
        cell.font = head_font
        cell.alignment = center
        cell.fill = head_fill

    ws.row_dimensions[2].height = 24
    ws.row_dimensions[3].height = 24
    for r in range(2, 4):
        for cidx in range(1, 16):
            ws.cell(row=r, column=cidx).border = border_all

    # Cuerpo
    start_row = 4
    for i, row in df.iterrows():
        r = start_row + i
        vals = [
            row["Nº"], row["Nombre"], row["Cédula de Identidad"], row["Institución"], row["Cargo"], row["Teléfono"],
            "X" if row["Genero_F"] else "", "X" if row["Genero_M"] else "", "X" if row["Genero_LGBTIQ+"] else "",
            "X" if row["Sexo_H"] else "", "X" if row["Sexo_M"] else "", "X" if row["Sexo_I"] else "",
            "X" if row["Edad_18_35"] else "", "X" if row["Edad_36_64"] else "", "X" if row["Edad_65_mas"] else ""
        ]
        for cidx, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=cidx, value=v)
            cell.border = border_all
            cell.alignment = center if cidx in [1,7,8,9,10,11,12,13,14,15] else left
        ws.row_dimensions[r].height = 20

    ws.freeze_panes = "B4"
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()

# --- Botón de descarga ---
excel_bytes = build_excel_asistencia(edited_df[COLS])
st.download_button(
    "⬇️ Descargar Excel",
    data=excel_bytes,
    file_name=f"Lista_Asistencia_LineasAccion_{date.today():%Y%m%d}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="Descarga el Excel con el mismo orden y encabezados combinados."
)





