{
  "nbformat": 4,
  "nbformat_minor": 0,
  "metadata": {
    "colab": {
      "provenance": [],
      "authorship_tag": "ABX9TyMCCk3oi+Qj3rVLHA4tjXey",
      "include_colab_link": true
    },
    "kernelspec": {
      "name": "python3",
      "display_name": "Python 3"
    },
    "language_info": {
      "name": "python"
    }
  },
  "cells": [
    {
      "cell_type": "markdown",
      "metadata": {
        "id": "view-in-github",
        "colab_type": "text"
      },
      "source": [
        "<a href=\"https://colab.research.google.com/github/RicardoUrrutia-Tech/HR-Automation-Report/blob/main/HR_Assistance_Automation.py\" target=\"_parent\"><img src=\"https://colab.research.google.com/assets/colab-badge.svg\" alt=\"Open In Colab\"/></a>"
      ]
    },
    {
      "cell_type": "code",
      "execution_count": 1,
      "metadata": {
        "colab": {
          "base_uri": "https://localhost:8080/",
          "height": 384
        },
        "id": "sJFkwIS8KCyy",
        "outputId": "2ae32188-73c7-40d3-e9ee-ba45526db729"
      },
      "outputs": [
        {
          "output_type": "error",
          "ename": "ModuleNotFoundError",
          "evalue": "No module named 'streamlit'",
          "traceback": [
            "\u001b[0;31m---------------------------------------------------------------------------\u001b[0m",
            "\u001b[0;31mModuleNotFoundError\u001b[0m                       Traceback (most recent call last)",
            "\u001b[0;32m/tmp/ipython-input-196513018.py\u001b[0m in \u001b[0;36m<cell line: 0>\u001b[0;34m()\u001b[0m\n\u001b[1;32m      1\u001b[0m \u001b[0;31m# app.py\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m----> 2\u001b[0;31m \u001b[0;32mimport\u001b[0m \u001b[0mstreamlit\u001b[0m \u001b[0;32mas\u001b[0m \u001b[0mst\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m      3\u001b[0m \u001b[0;32mimport\u001b[0m \u001b[0mpandas\u001b[0m \u001b[0;32mas\u001b[0m \u001b[0mpd\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m      4\u001b[0m \u001b[0;32mimport\u001b[0m \u001b[0mre\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m      5\u001b[0m \u001b[0;32mimport\u001b[0m \u001b[0mio\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
            "\u001b[0;31mModuleNotFoundError\u001b[0m: No module named 'streamlit'",
            "",
            "\u001b[0;31m---------------------------------------------------------------------------\u001b[0;32m\nNOTE: If your import is failing due to a missing package, you can\nmanually install dependencies using either !pip or !apt.\n\nTo view examples of installing some common dependencies, click the\n\"Open Examples\" button below.\n\u001b[0;31m---------------------------------------------------------------------------\u001b[0m\n"
          ],
          "errorDetails": {
            "actions": [
              {
                "action": "open_url",
                "actionText": "Open Examples",
                "url": "/notebooks/snippets/importing_libraries.ipynb"
              }
            ]
          }
        }
      ],
      "source": [
        "# app.py\n",
        "import streamlit as st\n",
        "import pandas as pd\n",
        "import re\n",
        "import io\n",
        "\n",
        "st.set_page_config(page_title=\"Informe de Asistencias\", layout=\"wide\")\n",
        "\n",
        "st.title(\"📊 Generador de Informe de Asistencias\")\n",
        "\n",
        "# -------------------------\n",
        "# 1️⃣ Subir archivo\n",
        "# -------------------------\n",
        "uploaded_file = st.file_uploader(\"Selecciona el archivo de asistencias (Excel)\", type=[\"xlsx\"])\n",
        "if uploaded_file:\n",
        "    df = pd.read_excel(uploaded_file)\n",
        "\n",
        "    # -------------------------\n",
        "    # 2️⃣ Ingresar período manualmente\n",
        "    # -------------------------\n",
        "    # Intentar detectar automáticamente\n",
        "    match = re.search(r\"Asistencia[_\\s-]*(\\w+)[_\\s-]*(\\d{4})\", uploaded_file.name)\n",
        "    default_periodo = \"\"\n",
        "    if match:\n",
        "        mes, anio = match.groups()\n",
        "        default_periodo = f\"{mes} {anio}\"\n",
        "\n",
        "    st.write(f\"Nombre del archivo: {uploaded_file.name}\")\n",
        "    mes = st.text_input(\"Mes (Ej: Octubre)\", value=default_periodo.split()[0] if default_periodo else \"\")\n",
        "    anio = st.text_input(\"Año (Ej: 2025)\", value=default_periodo.split()[1] if default_periodo else \"\")\n",
        "    periodo = f\"{mes} {anio}\"\n",
        "\n",
        "    # -------------------------\n",
        "    # 3️⃣ Agregar columna Periodo\n",
        "    # -------------------------\n",
        "    df.insert(0, \"Periodo\", periodo)\n",
        "\n",
        "    # -------------------------\n",
        "    # 4️⃣ Normalizar horas y fechas\n",
        "    # -------------------------\n",
        "    def normalizar_hora_str(hora_str):\n",
        "        try:\n",
        "            s = str(hora_str).strip()\n",
        "            if s in [\"\", \"nan\", \"None\"]:\n",
        "                return pd.NaT\n",
        "            parts = s.split(\":\")\n",
        "            if len(parts) == 1:\n",
        "                h, m, sec = int(parts[0]), 0, 0\n",
        "            elif len(parts) == 2:\n",
        "                h, m, sec = int(parts[0]), int(parts[1]), 0\n",
        "            else:\n",
        "                h, m, sec = int(parts[0]), int(parts[1]), int(parts[2])\n",
        "            return pd.to_timedelta(f\"{h:02d}:{m:02d}:{sec:02d}\")\n",
        "        except:\n",
        "            return pd.NaT\n",
        "\n",
        "    for col in [\"Hora Entrada\", \"Hora Salida\"]:\n",
        "        if col in df.columns:\n",
        "            df[col] = df[col].apply(normalizar_hora_str)\n",
        "\n",
        "    for col in [\"Fecha Entrada\", \"Fecha Salida\"]:\n",
        "        if col in df.columns:\n",
        "            df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce').dt.normalize()\n",
        "\n",
        "    for col in [\"Retraso (horas)\", \"Salida Anticipada (horas)\"]:\n",
        "        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)\n",
        "\n",
        "    for col in [\"Hora Entrada\", \"Hora Salida\"]:\n",
        "        df[col] = df[col].apply(lambda x: (x / pd.Timedelta(days=1)) if pd.notna(x) else None)\n",
        "\n",
        "    # -------------------------\n",
        "    # 5️⃣ Descargar Excel procesado\n",
        "    # -------------------------\n",
        "    import xlsxwriter\n",
        "    output = io.BytesIO()\n",
        "    with pd.ExcelWriter(output, engine=\"xlsxwriter\", datetime_format='dd/mm/yyyy', date_format='dd/mm/yyyy') as writer:\n",
        "        df.to_excel(writer, sheet_name=\"Detalle\", index=False)\n",
        "        workbook = writer.book\n",
        "        ws_det = writer.sheets[\"Detalle\"]\n",
        "\n",
        "        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#A697ED', 'border': 1})\n",
        "        for i, col in enumerate(df.columns):\n",
        "            ws_det.write(0, i, col, header_fmt)\n",
        "\n",
        "    output.seek(0)\n",
        "    st.download_button(\n",
        "        label=\"📥 Descargar Excel procesado\",\n",
        "        data=output,\n",
        "        file_name=f\"Informe_Asistencias_{periodo.replace(' ', '_')}.xlsx\",\n",
        "        mime=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\"\n",
        "    )\n",
        "\n",
        "    st.success(\"✅ Archivo listo para descargar\")"
      ]
    }
  ]
}