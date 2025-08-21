# =========================
# ⬇️ 3) Excel oficial (recreado desde cero, sin plantilla)
# =========================
with tab_excel:
    st.markdown("### ⬇️ Descargar Excel oficial (estructura replicada)")
    st.caption("Se recrea la estructura de la minuta sin usar plantilla externa. Si agregas 'logo_izq.png' y/o 'logo_der.png' en la carpeta, se insertan arriba.")

    df_all = fetch_all_df(include_id=True)

    def build_excel_official_from_scratch(
        fecha: date,
        lugar: str,
        hora_ini,
        hora_fin,
        estrategia: str,
        delegacion: str,
        rows_df: pd.DataFrame,
        per_page: int = 16
    ) -> bytes:
        """
        Crea un libro desde cero con la estructura:
        - Encabezados tipo 'Fecha:', 'Lugar:', 'Hora Inicio:', 'Hora Finalización:', 'Estrategia o Programa:', 'Dirección/Delegación...'
        - Tabla con:
          B: Nº | C:E Nombre (merge por fila) | F Cédula | G Institución | H Cargo | I Teléfono
          J/K/L Género: F/M/LGBTIQ+
          M/N/O Sexo: H/M/I
          P/Q/R Rango de Edad: 18-35 / 36-64 / 65+
          S: FIRMA
        Paginación automática cada 16 filas en hojas 'Minuta', 'Minuta 2', ...
        Inserta logos si existen 'logo_izq.png' y/o 'logo_der.png' en la carpeta.
        """
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
            from openpyxl.drawing.image import Image as XLImage
        except Exception:
            st.error("Falta 'openpyxl' en requirements.txt")
            return b""

        # ===== helpers de estilo =====
        head_fill  = PatternFill("solid", fgColor="DDE7FF")
        group_fill = PatternFill("solid", fgColor="B7C6F9")
        head_font  = Font(bold=True)
        title_font = Font(bold=True, size=12)
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
        thin = Side(style="thin", color="000000")
        border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

        # ===== libro =====
        wb = Workbook()
        ws0 = wb.active
        ws0.title = "Minuta"

        # ===== función para construir una hoja =====
        def _setup_sheet(ws):
            # anchos de columnas aproximados a la referencia
            widths = {
                "A": 2, "B": 6, "C": 22, "D": 22, "E": 22, "F": 18, "G": 22,
                "H": 20, "I": 16, "J": 6, "K": 6, "L": 10, "M": 6, "N": 6, "O": 6,
                "P": 12, "Q": 12, "R": 12, "S": 16
            }
            for col, w in widths.items():
                ws.column_dimensions[col].width = w

            # ===== logos opcionales =====
            try:
                from pathlib import Path
                if Path("logo_izq.png").exists():
                    img = XLImage("logo_izq.png")
                    img.width, img.height = int(img.width*0.6), int(img.height*0.6)
                    ws.add_image(img, "B2")
                if Path("logo_der.png").exists():
                    img2 = XLImage("logo_der.png")
                    img2.width, img2.height = int(img2.width*0.6), int(img2.height*0.6)
                    ws.add_image(img2, "Q2")
            except Exception:
                pass  # si no hay PIL, simplemente sigue sin logos

            # ===== encabezados (líneas 6–8 aprox.) =====
            ws["B6"].value = f"Fecha: {fecha.day} " + fecha.strftime("%B") + f" {fecha.year}"
            ws["E6"].value = f"Lugar:  {lugar}"
            ws.merge_cells("E6:I6")
            ws["J6"].value = f"Hora Inicio: {hora_ini.strftime('%H:%M')}"
            ws["Q6"].value = f"Hora Finalización: {hora_fin.strftime('%H:%M')}"

            ws["B6"].font = title_font
            ws["E6"].font = title_font
            ws["J6"].font = title_font
            ws["Q6"].font = title_font

            ws["B7"].value = f"Estrategia o Programa: {estrategia}"
            ws.merge_cells("B7:G7")
            ws["H7"].value = "AC... acción, acciones estratégicas, indicadores y metas."
            ws.merge_cells("H7:S8")
            ws["B7"].font = title_font

            ws["B8"].value = "Dirección / Delegación Policial:"
            ws["E8"].value = delegacion

            # ===== cabecera de tabla (filas 9–10) =====
            # Bloques
            ws.merge_cells("B9:E10"); ws["B9"].value = "Nombre"
            ws["F9"].value = "Cédula de Identidad"
            ws["G9"].value = "Institución"
            ws["H9"].value = "Cargo"
            ws["I9"].value = "Teléfono"
            ws.merge_cells("J9:L9"); ws["J9"].value = "Género"
            ws.merge_cells("M9:O9"); ws["M9"].value = "Sexo (Hombre, Mujer o Intersex)"
            ws.merge_cells("P9:R9"); ws["P9"].value = "Rango de Edad"
            ws["S9"].value = "FIRMA"

            # Subencabezados fila 10
            ws["J10"].value = "F"; ws["K10"].value = "M"; ws["L10"].value = "LGBTIQ+"
            ws["M10"].value = "H"; ws["N10"].value = "M"; ws["O10"].value = "I"
            ws["P10"].value = "18 a 35 años"; ws["Q10"].value = "36 a 64 años"; ws["R10"].value = "65 años o más"

            # estilos de cabecera
            for rng in ["B9:E10","J9:L9","M9:O9","P9:R9"]:
                top_left = rng.split(":")[0]
                c = ws[top_left]
                c.font = head_font; c.alignment = center; c.fill = group_fill
            for cell in ["F9","G9","H9","I9","S9"]:
                ws[cell].font = head_font; ws[cell].alignment = center; ws[cell].fill = head_fill
            for cell in ["J10","K10","L10","M10","N10","O10","P10","Q10","R10"]:
                ws[cell].font = head_font; ws[cell].alignment = center; ws[cell].fill = head_fill

            # bordes cabecera
            for r in range(9, 11):
                for c in range(2, 20):  # B..S
                    ws.cell(row=r, column=c).border = border_all

        # ===== función para llenar datos en una hoja =====
        def _fill_rows(ws, df_slice: pd.DataFrame, start_row: int = 11):
            for i, (_, row) in enumerate(df_slice.iterrows()):
                r = start_row + i
                # Nº
                ws[f"B{r}"].value = i + 1
                ws[f"B{r}"].alignment = center
                # Nombre (merge por fila C:E)
                ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)  # C:E
                ws[f"C{r}"].value = str(row["Nombre"] or "")
                ws[f"C{r}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

                ws[f"F{r}"].value = str(row["Cédula de Identidad"] or "")
                ws[f"G{r}"].value = str(row["Institución"] or "")
                ws[f"H{r}"].value = str(row["Cargo"] or "")
                ws[f"I{r}"].value = str(row["Teléfono"] or "")

                # limpiar marcas
                for c in ["J","K","L","M","N","O","P","Q","R"]:
                    ws[f"{c}{r}"].value = ""

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

                # Firma (texto de referencia)
                ws[f"S{r}"].value = "Virtual"
                ws[f"S{r}"].alignment = center

                # bordes de la fila
                for c in range(2, 20):  # B..S
                    ws.cell(row=r, column=c).border = border_all

            # congelar paneles para navegar
            ws.freeze_panes = "C11"

        # ===== paginar =====
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

    # Botón de descarga (sin plantilla)
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






