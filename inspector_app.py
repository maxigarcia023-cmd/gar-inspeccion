import sys, io
import streamlit as st
import pandas as pd
import pytz
try:
    import openpyxl  # <- clave para Excel
    has_openpyxl = True
except Exception as e:
    has_openpyxl = False
    openpyxl = e

st.title("Diagnóstico de entorno – GAR")
st.write("Python:", sys.version)
st.write("pandas:", pd.__version__)
st.write("pytz:", pytz.__version__)
st.write("openpyxl:", openpyxl.__version__ if has_openpyxl else f"NO CARGÓ: {openpyxl}")

# Test rápido de escritura a Excel (sin save(), usando BytesIO)
df = pd.DataFrame({"ok": [1,2,3]})
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="DATOS", index=False)
st.download_button("Descargar Excel de prueba", data=buf.getvalue(),
                   file_name="test.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.success("Si ves este mensaje y el botón de descarga, el entorno de Streamlit está OK.")
