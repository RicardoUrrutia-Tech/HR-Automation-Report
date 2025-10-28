# 📊 Generador Automático de Reportes de Asistencia

Esta aplicación permite generar reportes de asistencia en formato Excel a partir de archivos mensuales exportados desde Power BI o sistemas de control horario.
La herramienta está construida con **Streamlit** y utiliza **Pandas** y **XlsxWriter** para procesar los datos y generar un archivo Excel con tres hojas principales:

* **Detalle:** datos limpios y formateados.
* **Resumen:** métricas globales del periodo.
* **Incidencias:** totales de retrasos y salidas anticipadas por empleado.

---

## 🚀 Características principales

* Carga directa de archivos Excel (`.xlsx`) mediante la interfaz web.
* Detección automática del **mes y año** a partir del nombre del archivo.
* Generación de un **reporte combinado** cuando se suben varios archivos (por ejemplo, Octubre y Noviembre 2025).
* Formato visual coherente con la **gama cromática de Cabify**, basado en tonos púrpura y lavanda.
* Fórmulas dinámicas en Excel para mantener los cálculos actualizados.
* Descarga automática del reporte final en formato `.xlsx`.

---

