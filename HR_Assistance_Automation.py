import streamlit as st
import pandas as pd
import xlsxwriter
from io import BytesIO
import re

st.set_page_config(page_title="Informe de Asistencias", layout="wide")
st.title("📊 Generador de Informe de Asistencias")
st.markdown("Sube el archivo de asistencias, ingresa el período y genera el informe consolidado en Excel.")

# --- Subir archivo ---

uploaded_file = st.file_uploader("Selecciona el archivo Excel de asistencias (.xlsx)", type=["xlsx"])

if uploaded_file:
    st.success(f"✅ Archivo cargado correctamente: {uploaded_file.name}")
    # el resto del código indentado va aquí...


# Intentar detectar período automáticamente
match = re.search(r"Asistencias?[_\s-]*(\w+)[_\s-]*(\d{4})", uploaded_file.name)
if match:
    mes_detectado, anio_detectado = match.groups()
else:
    mes_detectado, anio_detectado = "", ""

# Permitir ingreso o confirmación manual
mes = st.text_input("Mes del archivo (Ej: Octubre)", value=mes_detectado)
anio = st.text_input("Año del archivo (Ej: 2025)", value=anio_detectado)

if st.button("Generar Informe"):
    if not mes or not anio:
        st.error("⚠️ Debes ingresar el mes y el año del período.")
    else:
        periodo = f"{mes} {anio}"
        st.info(f"📅 Generando informe para el período: **{periodo}** ...")

        # --- Leer archivo base ---
        df = pd.read_excel(uploaded_file)
        df.insert(0, "Periodo", periodo)

        # --- Normalizar horas ---
        def normalizar_hora_str(hora_str):
            try:
                s = str(hora_str).strip()
                if s in ["", "nan", "None"]:
                    return pd.NaT
                parts = s.split(":")
                if len(parts) == 1:
                    h, m, sec = int(parts[0]), 0, 0
                elif len(parts) == 2:
                    h, m, sec = int(parts[0]), int(parts[1]), 0
                else:
                    h, m, sec = int(parts[0]), int(parts[1]), int(parts[2])
                return pd.to_timedelta(f"{h:02d}:{m:02d}:{sec:02d}")
            except:
                return pd.NaT

        for col in ["Hora Entrada", "Hora Salida"]:
            if col in df.columns:
                df[col] = df[col].apply(normalizar_hora_str)

        for col in ["Fecha Entrada", "Fecha Salida"]:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce').dt.normalize()

        for col in ["Retraso (horas)", "Salida Anticipada (horas)"]:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        for col in ["Hora Entrada", "Hora Salida"]:
            df[col] = df[col].apply(lambda x: (x / pd.Timedelta(days=1)) if pd.notna(x) else None)

        # --- Crear Excel en memoria ---
        output = BytesIO()
        output_file = f"Informe_Asistencias_{periodo.replace(' ', '_')}.xlsx"

        def col_idx_to_excel(idx):
            letters = ""
            while idx >= 0:
                letters = chr(idx % 26 + 65) + letters
                idx = idx // 26 - 1
            return letters

        with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format='dd/mm/yyyy', date_format='dd/mm/yyyy') as writer:
            # Hoja Detalle
            df.to_excel(writer, sheet_name="Detalle", index=False)
            workbook = writer.book
            ws_det = writer.sheets["Detalle"]

            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#A697ED', 'border': 1})
            time_fmt = workbook.add_format({'num_format': 'hh:mm:ss'})
            date_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy'})

            n_rows, n_cols = df.shape
            for i, col in enumerate(df.columns):
                ws_det.write(0, i, col, header_fmt)
                if col in ["Fecha Entrada", "Fecha Salida"]:
                    ws_det.set_column(i, i, 12, date_fmt)
                elif col in ["Hora Entrada", "Hora Salida"]:
                    ws_det.set_column(i, i, 10, time_fmt)
                else:
                    max_len = max(len(col), df[col].astype(str).map(len).max()) + 2
                    ws_det.set_column(i, i, max_len)
            ws_det.freeze_panes(1, 0)
            ws_det.autofilter(0, 0, n_rows, n_cols-1)

            # Hoja Resumen
            ws_res = workbook.add_worksheet("Resumen")
            resumen_headers = [
                "Periodo", "Total Registros", "Total Retrasos (h)", "Total Salidas Anticipadas (h)",
                "Turnos Inconsistentes", "% Retrasos", "% Salidas Anticipadas",
                "% Jornadas Completas", "Tiempo Medio Retraso (h)", "Tiempo Medio Salida Anticipada (h)"
            ]
            for j, h in enumerate(resumen_headers):
                ws_res.write(0, j, h, header_fmt)
                ws_res.set_column(j, j, len(h) + 4)

            ws_res.write(1, 0, periodo)

            colmap = {c: col_idx_to_excel(i) for i, c in enumerate(df.columns)}
            total_filas = len(df) + 1
            col_retraso = colmap.get("Retraso (horas)")
            col_salida = colmap.get("Salida Anticipada (horas)")
            col_incons = colmap.get("Inconsistencia de Turno")
            first_col = col_idx_to_excel(0)

            ws_res.write_formula("B2", f"=COUNTA(Detalle!{first_col}2:{first_col}{total_filas})")
            ws_res.write_formula("C2", f"=SUM(Detalle!{col_retraso}2:{col_retraso}{total_filas})")
            ws_res.write_formula("D2", f"=SUM(Detalle!{col_salida}2:{col_salida}{total_filas})")
            ws_res.write_formula("E2", f"=COUNTA(Detalle!{col_incons}2:{col_incons}{total_filas})")
            ws_res.write_formula("F2", f"=IF(B2=0,0,COUNTIF(Detalle!{col_retraso}2:{col_retraso}{total_filas},\">0\")/B2*100)")
            ws_res.write_formula("G2", f"=IF(B2=0,0,COUNTIF(Detalle!{col_salida}2:{col_salida}{total_filas},\">0\")/B2*100)")
            ws_res.write_formula("H2", f"=IF(B2=0,0,COUNTIFS(Detalle!{col_retraso}2:{col_retraso}{total_filas},0,Detalle!{col_salida}2:{col_salida}{total_filas},0)/B2*100)")
            ws_res.write_formula("I2", f"=IFERROR(AVERAGEIF(Detalle!{col_retraso}2:{col_retraso}{total_filas},\">0\"),0)")
            ws_res.write_formula("J2", f"=IFERROR(AVERAGEIF(Detalle!{col_salida}2:{col_salida}{total_filas},\">0\"),0)")

            # ✅ Hoja Incidencias corregida
            ws_inc = workbook.add_worksheet("Incidencias")
            headers_inc = ["Nombre", "Horas Retraso", "Horas Salida Anticipada", "Horas Incidencia"]
            merge_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#A697ED', 'border': 1})
            ws_inc.merge_range(0, 1, 0, 3, f"Periodo: {periodo}", merge_format)

            header_fmt_inc = workbook.add_format({'bold': True, 'bg_color': '#A697ED', 'border': 1})
            for j, h in enumerate(headers_inc):
                ws_inc.write(1, j, h, header_fmt_inc)

            # detectar columna real "Nombre"
            nombre_col_idx = df.columns.get_loc("Nombre")
            nombre_col_excel = col_idx_to_excel(nombre_col_idx)

            nombres = sorted(df["Nombre"].dropna().unique())
            for i, nombre in enumerate(nombres, start=2):
                ws_inc.write(i, 0, nombre)
                fila = i + 1
                ws_inc.write_formula(f"B{fila}", f"=SUMIF(Detalle!${nombre_col_excel}:${nombre_col_excel},A{fila},Detalle!{col_retraso}:{col_retraso})")
                ws_inc.write_formula(f"C{fila}", f"=SUMIF(Detalle!${nombre_col_excel}:${nombre_col_excel},A{fila},Detalle!{col_salida}:{col_salida})")
                ws_inc.write_formula(f"D{fila}", f"=B{fila}+C{fila}")

            ws_inc.freeze_panes(2, 0)
            ws_inc.autofilter(1, 0, len(nombres)+1, len(headers_inc)-1)
            for i, col in enumerate(headers_inc):
                ws_inc.set_column(i, i, max(len(col)+4, 18))

        output.seek(0)
        st.success("✅ Informe generado correctamente.")
        st.download_button(
            label="⬇️ Descargar Informe Excel",
            data=output,
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
