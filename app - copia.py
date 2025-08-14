# =========================
# 📋 Lista de asistencia – Seguimiento de líneas de acción
# =========================
import streamlit as st
from io import BytesIO
from datetime import date
import xlsxwriter  # (lo usa pandas.ExcelWriter por debajo)

st.markdown("## 📋 Lista de asistencia – Seguimiento de líneas de acción")

filas = st.number_input("Cantidad de filas a generar", min_value=5, max_value=500, value=30, step=5)

def build_excel_asistencia(n_rows: int) -> bytes:
    """
    Genera un archivo Excel con el formato de lista de asistencia:
    Nº | Nombre | Cédula de Identidad | Institución | Cargo | Teléfono |
    Género [F, M, LGBTIQ+] | Sexo (H/M/I) | Rango de edad [18–35, 36–64, 65+]
    """
    output = BytesIO()

    # Abrimos el writer con xlsxwriter para poder dar formato
    import pandas as pd
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Creamos un DF vacío solo para inicializar la hoja
        pd.DataFrame().to_excel(writer, sheet_name="Asistencia", index=False)
        wb  = writer.book
        ws  = writer.sheets["Asistencia"]

        # ----- Formatos -----
        title_fmt = wb.add_format({
            "bold": True, "font_size": 14, "align": "center", "valign": "vcenter",
            "font_color": "white", "bg_color": "#1F3B73"  # azul oscuro elegante
        })
        head_fmt = wb.add_format({
            "bold": True, "align": "center", "valign": "vcenter",
            "text_wrap": True, "border": 1, "bg_color": "#DDE7FF"  # azul claro
        })
        group_fmt = wb.add_format({
            "bold": True, "align": "center", "valign": "vcenter",
            "text_wrap": True, "border": 1, "bg_color": "#B7C6F9"  # tono más intenso
        })
        cell_left = wb.add_format({"border": 1, "align": "left", "valign": "vcenter"})
        cell_center = wb.add_format({"border": 1, "align": "center", "valign": "vcenter"})

        # ----- Columnas (A:O) -----
        # A  Nº
        # B  Nombre
        # C  Cédula de Identidad
        # D  Institución
        # E  Cargo
        # F  Teléfono
        # G  Género F
        # H  Género M
        # I  Género LGBTIQ+
        # J  Sexo H
        # K  Sexo M
        # L  Sexo I
        # M  18 a 35 años
        # N  36 a 64 años
        # O  65 años o más

        # Ancho de columnas
        ws.set_column("A:A", 5)    # Nº
        ws.set_column("B:B", 28)   # Nombre
        ws.set_column("C:C", 18)   # Cédula
        ws.set_column("D:D", 24)   # Institución
        ws.set_column("E:E", 20)   # Cargo
        ws.set_column("F:F", 16)   # Teléfono
        ws.set_column("G:O", 12)   # Columnas de marca (centradas)

        # ----- Título -----
        ws.merge_range("A1:O1", "Lista de asistencia – Seguimiento de líneas de acción", title_fmt)
        ws.set_row(0, 28)

        # ----- Encabezados en dos filas -----
        # Fila 2 (grupos)
        ws.merge_range("A2:A3", "Nº", head_fmt)
        ws.merge_range("B2:B3", "Nombre", head_fmt)
        ws.merge_range("C2:C3", "Cédula de Identidad", head_fmt)
        ws.merge_range("D2:D3", "Institución", head_fmt)
        ws.merge_range("E2:E3", "Cargo", head_fmt)
        ws.merge_range("F2:F3", "Teléfono", head_fmt)

        ws.merge_range("G2:I2", "Género", group_fmt)  # F, M, LGBTIQ+
        ws.merge_range("J2:L2", "Sexo (Hombre, Mujer o Intersex)", group_fmt)  # H, M, I
        ws.merge_range("M2:O2", "Rango de Edad", group_fmt)  # 18–35, 36–64, 65+

        # Fila 3 (subcolumnas)
        ws.write("G3", "F", head_fmt)
        ws.write("H3", "M", head_fmt)
        ws.write("I3", "LGBTIQ+", head_fmt)

        ws.write("J3", "H", head_fmt)
        ws.write("K3", "M", head_fmt)
        ws.write("L3", "I", head_fmt)

        ws.write("M3", "18 a 35 años", head_fmt)
        ws.write("N3", "36 a 64 años", head_fmt)
        ws.write("O3", "65 años o más", head_fmt)

        ws.set_row(1, 26)
        ws.set_row(2, 26)

        # ----- Cuerpo (n_rows) -----
        start_row = 3  # fila 4 en Excel
        for i in range(n_rows):
            r = start_row + i
            # Nº
            ws.write_number(r, 0, i + 1, cell_center)
            # Datos principales
            ws.write_blank(r, 1, None, cell_left)    # Nombre
            ws.write_blank(r, 2, None, cell_left)    # Cédula
            ws.write_blank(r, 3, None, cell_left)    # Institución
            ws.write_blank(r, 4, None, cell_left)    # Cargo
            ws.write_blank(r, 5, None, cell_left)    # Teléfono
            # Marcas (centradas, para colocar X si aplica)
            for c in range(6, 15):  # G..O
                ws.write_blank(r, c, None, cell_center)

        # Congelar encabezados
        ws.freeze_panes(start_row, 1)

    return output.getvalue()

# Botón de descarga
excel_bytes = build_excel_asistencia(int(filas))
st.download_button(
    "⬇️ Descargar Excel",
    data=excel_bytes,
    file_name=f"Lista_Asistencia_LineasAccion_{date.today():%Y%m%d}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)





