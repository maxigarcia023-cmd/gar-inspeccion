# inspector_app.py
# -*- coding: utf-8 -*-
"""
App ligera para acelerar inspecciones H&S:
- Captura foto desde cámara o archivo.
- Registro guiado de hallazgos (área, no conformidad, G, P, medida, plazo, etc.).
- **Lee tu hoja "TABLA RIESGO"** y aplica **la matriz oficial de tu libro** (sin inventar) para clasificar el riesgo por **(Probabilidad, Gravedad)** y agregar **Acción y cronograma** según la tabla.
- Exporta a Excel (agrega/crea hoja "DATOS" normalizada) para tu libro de informes.
- Genera en el acto un "Resumen informativo" en texto usando una plantilla Jinja2 **editable**.
- Soporta **logo** en el encabezado del resumen (embebido en el .md como data URI).

Requisitos (instalar una vez):
    pip install streamlit pandas openpyxl pillow jinja2 pytz

Ejecución local:
    streamlit run inspector_app.py

NOTA IMPORTANTE:
- La app **no inventa** clasificaciones: toma tu matriz (hoja "TABLA RIESGO") y el bloque "Evaluacion del riesgo / Accion y cronograma" tal como están en tu archivo.
- Si subes tu libro actual, la app no toca tus hojas de presentación ("INFORME Nº1", etc.); solo crea/usa una hoja "DATOS" para registro normalizado.
"""
from __future__ import annotations
import io
import os
import base64
from dataclasses import dataclass, asdict
from datetime import date, datetime
from typing import List, Optional, Dict, Any, Tuple

import pandas as pd
from PIL import Image
from jinja2 import Template
import pytz

try:
    import openpyxl  # noqa: F401
    from openpyxl import load_workbook
except Exception as e:  # pragma: no cover
    raise SystemExit("Falta 'openpyxl'. Instala con: pip install openpyxl")

import streamlit as st

# ------------------------- Configuración básica ------------------------- #
st.set_page_config(page_title="Inspección H&S – Informe + Resumen", layout="wide")
st.title("Inspección H&S – Informe y Resumen")
st.caption("Capturá evidencias, registrá hallazgos y generá el informe/recap al instante, sin inventar datos.")

# Zona horaria del usuario (San Juan, AR)
TZ = pytz.timezone("America/Argentina/San_Juan")
HOY = datetime.now(TZ).date()

# ------------------------- Modelo de datos ------------------------------ #
@dataclass
class Observacion:
    fecha: date
    empresa: str
    ubicacion: str
    area: str
    no_conformidad: str
    descripcion: str
    gravedad: int
    probabilidad: int
    riesgo: int  # GxP
    categoria: Optional[str]
    accion: Optional[str]
    medida: str
    responsable: str
    plazo: Optional[date]
    estado: str  # Pendiente / Cerrada
    normativa: Optional[str]
    foto_nombre: Optional[str]

    def to_row(self) -> Dict[str, Any]:
        d = asdict(self)
        # Normalizar fechas a ISO yyyy-mm-dd para Excel
        d["fecha"] = self.fecha.isoformat()
        d["plazo"] = self.plazo.isoformat() if self.plazo else ""
        return d

CAMPOS_ORDEN = [
    "fecha", "empresa", "ubicacion", "area", "no_conformidad", "descripcion",
    "gravedad", "probabilidad", "riesgo", "categoria", "accion",
    "medida", "responsable", "plazo", "estado", "normativa", "foto_nombre"
]

# ------------------------- Plantilla de resumen (editable) ------------- #
DEFAULT_RESUMEN_TEMPLATE = Template(
    (
        """
{% if logo_data %}![]({{ logo_data }})

{% endif %}
**Fecha:** {{ fecha }}  
**Empresa:** {{ empresa }}  
**Ubicación:** {{ ubicacion }}

**Resumen de hallazgos ({{ total }}):**
{% for area, items in por_area.items() %}
- **Área {{ area }}** ({{ items|length }}):
{% for it in items %}
  - {{ it.no_conformidad }} — Riesgo: {{ it.riesgo }} (G={{ it.gravedad }}, P={{ it.probabilidad }}){% if it.categoria %} — **{{ it.categoria }}**{% endif %}{% if it.medida %}. Medida: {{ it.medida }}{% endif %}{% if it.plazo %} (Plazo: {{ it.plazo }}){% endif %}{% if it.normativa %} [Normativa: {{ it.normativa }}]{% endif %}
    {% if it.accion %}_Acción según tabla:_ {{ it.accion }}{% endif %}
{% endfor %}
{% endfor %}

**Por qué actuar ahora:**
- Reducir probabilidad/consecuencia de incidentes (eléctricos, incendios, resbalones, presión, etc.).
- Evitar paradas no planificadas y costos asociados.
- Cumplir con Ley 19.587 y Decreto 351/79; registrar EPP conforme Res. SRT 299/2011.

**Próximos pasos:**
- Asignar responsables y normalizar plazos vencidos.
- Proveer/registrar EPP donde aplique.
- Programar mantenimiento correctivo/preventivo según hallazgos.
"""
    ).strip()
)

# ------------------------- Utilidades Excel y parsing ------------------- #

def cargar_o_crear_excel_bytes(existing_bytes: Optional[bytes], df_rows: pd.DataFrame) -> bytes:
    """Devuelve un xlsx con la hoja DATOS actualizada, preservando el resto de hojas si subiste un libro."""
    buf = io.BytesIO()
    if existing_bytes:
        wb = load_workbook(io.BytesIO(existing_bytes))
        # Leer hoja DATOS si existe
        try:
            existing_df = pd.read_excel(io.BytesIO(existing_bytes), sheet_name="DATOS")
        except Exception:
            existing_df = pd.DataFrame(columns=CAMPOS_ORDEN)
        full = pd.concat([existing_df, df_rows], ignore_index=True)
        full = full[CAMPOS_ORDEN]
        # Escribir en el mismo workbook y preservar hojas
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            writer.book = wb
            writer.sheets = {ws.title: ws for ws in wb.worksheets}
            full.to_excel(writer, sheet_name="DATOS", index=False)
            writer.save()
        return buf.getvalue()
    else:
        # Crear nuevo libro solo con DATOS
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_rows.to_excel(writer, sheet_name="DATOS", index=False)
            writer.save()
        return buf.getvalue()


def leer_matriz_y_acciones(xls_bytes: Optional[bytes]) -> Tuple[Optional[Dict[Tuple[int, int], str]], Optional[Dict[str, str]]]:
    """Lee tu hoja 'TABLA RIESGO' y devuelve:
    - mapa (P, G) -> categoria (strings tal cual en la hoja)
    - mapa categoria -> accion (desde columna 'Accion y cronograma')
    Si no se encuentra, retorna (None, None).
    """
    if not xls_bytes:
        return None, None
    try:
        df = pd.read_excel(io.BytesIO(xls_bytes), sheet_name="TABLA RIESGO")
    except Exception:
        return None, None

    # Normalizar strings
    df2 = df.copy()
    df2 = df2.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # 1) Ubicar encabezados "Gravedad" (columnas) y "Probabilidad" (filas)
    grav_pos = None
    prob_pos = None
    for i in range(len(df2)):
        for j, col in enumerate(df2.columns):
            val = df2.iloc[i, j]
            if isinstance(val, str):
                s = val.lower()
                if grav_pos is None and "gravedad" in s:
                    grav_pos = (i, j)
                if prob_pos is None and "probabilidad" in s:
                    prob_pos = (i, j)
    if grav_pos is None or prob_pos is None:
        return None, None

    gi, gj = grav_pos
    pi, pj = prob_pos

    # Se espera que justo debajo de "Gravedad" esté la fila de claves 1..4 y a la derecha de "Probabilidad" esté la columna 1..4
    try:
        cols_keys = [int(x) for x in list(df2.iloc[gi+1, gj:gj+4])]
        rows_keys = [int(x) for x in list(df2.iloc[pi:pi+4, pj+1])]
    except Exception:
        return None, None

    # Bloque de matriz de categorias: filas = pi..pi+3, columnas = gj..gj+3
    block = df2.iloc[pi:pi+4, gj:gj+4]
    mapa_pg_to_cat: Dict[Tuple[int, int], str] = {}
    for r_idx, p_val in enumerate(rows_keys):
        for c_idx, g_val in enumerate(cols_keys):
            val = block.iloc[r_idx, c_idx]
            if isinstance(val, str) and val.strip():
                mapa_pg_to_cat[(int(p_val), int(g_val))] = val.strip()

    # 2) Tabla Evaluacion del riesgo -> Accion y cronograma (dos columnas contiguas)
    cat_to_accion: Dict[str, str] = {}
    # Buscar celda con 'Evaluacion del riesgo'
    cat_col = None
    act_col = None
    cat_row = None
    for i in range(len(df2)):
        for j, col in enumerate(df2.columns):
            val = df2.iloc[i, j]
            if isinstance(val, str) and 'evaluacion del riesgo' in val.lower():
                cat_row = i
                cat_col = j
                if j + 1 < len(df2.columns):
                    act_col = j + 1
                break
        if cat_row is not None:
            break
    if cat_row is not None and cat_col is not None and act_col is not None:
        # Las filas siguientes (hasta que se acaben) contienen pares categoria-accion
        for k in range(cat_row + 1, len(df2)):
            cval = df2.iloc[k, cat_col]
            aval = df2.iloc[k, act_col]
            if isinstance(cval, str) and cval.strip():
                cat_to_accion[cval.strip()] = (aval or "").strip() if isinstance(aval, str) else ""
            # cortar si encontramos fila completamente vacía por varias filas seguidas
    if not mapa_pg_to_cat:
        mapa_pg_to_cat = None
    if not cat_to_accion:
        cat_to_accion = None
    return mapa_pg_to_cat, cat_to_accion

# ------------------------- Recursos de logo por defecto ----------------- #
# Logo por defecto (data URI) – se usa si no se sube ninguno por UI
DEFAULT_LOGO_DATA = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAK0AAABGCAYAAAC+LBQCAAAAAXNSR0IArs..."

# ------------------------- Sidebar (metadatos) ------------------------- #
with st.sidebar:
    st.header("Metadatos del informe")
    empresa = st.text_input("Empresa", placeholder="Razón social", value="")
    ubicacion = st.text_input("Ubicación", placeholder="Ciudad/Planta", value="")
    fecha = st.date_input("Fecha de visita", value=HOY)
    st.divider()
    st.subheader("Logo (opcional)")
    logo_file = st.file_uploader("Subir logo (PNG/JPG)", type=["png", "jpg", "jpeg"], key="logo")
    logo_b64_uri: Optional[str] = None
if logo_file is not None:
    ext = os.path.splitext(logo_file.name)[1].lower().strip('.')
    mime = 'image/png' if ext == 'png' else 'image/jpeg'
    logo_bytes = logo_file.getvalue()
    logo_b64_uri = f"data:{mime};base64,{base64.b64encode(logo_bytes).decode()}"
    st.image(logo_bytes, caption="Logo cargado", use_container_width=True)
else:
    # Usa el logo por defecto si está definido
    logo_b64_uri = DEFAULT_LOGO_DATA
    st.divider()
    st.subheader("Plantilla de resumen (Jinja2)")
    plantilla_txt = st.text_area(
        "Podés personalizar el texto. Variables: fecha, empresa, ubicacion, total, por_area (lista de dicts por área), logo_data",
        value=DEFAULT_RESUMEN_TEMPLATE.template,
        height=260,
    )

# ------------------------- Entrada de hallazgos ------------------------ #
st.subheader("Registrar hallazgo")
col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    area = st.text_input("Área", placeholder="Hornos, Empaquetado, Caldera, Rampa, etc.")
    no_conf = st.text_area("No conformidad (título)", placeholder="Ej.: Tablero eléctrico sin tapas y sin señalizar")
    desc = st.text_area("Descripción/Evidencia", placeholder="Detalle objetivo de la condición observada y riesgos asociados")
with col2:
    gravedad = st.number_input("Gravedad (1-4)", min_value=1, max_value=4, value=2)
    prob = st.number_input("Probabilidad (1-4)", min_value=1, max_value=4, value=2)
    riesgo = int(gravedad * prob)
    medida = st.text_area("Medida preventiva", placeholder="Acción concreta: mantenimiento correctivo, señalización, recarga extintor, etc.")
with col3:
    responsable = st.text_input("Responsable", placeholder="Nombre/Cargo")
    plazo = st.date_input("Plazo de corrección (opcional)", value=HOY)
    estado = st.selectbox("Estado", ["Pendiente", "Cerrada"], index=0)
    normativa = st.text_input("Normativa (opcional)", placeholder="Ley 19.587; Dto 351/79; Res. SRT 299/2011; IRAM 3501, etc.")

st.markdown(f"**Riesgo calculado (GxP):** {riesgo}")

# Foto (opcional)
colA, colB = st.columns([1,1])
with colA:
    foto_cam = st.camera_input("Tomar foto (opcional)")
with colB:
    foto_file = st.file_uploader("O subir foto (JPG/PNG)", type=["jpg", "jpeg", "png"], key="foto")

foto_bytes = None
foto_nombre = None
if foto_cam is not None:
    foto_bytes = foto_cam.getvalue()
    foto_nombre = f"foto_{datetime.now(TZ).strftime('%Y%m%d_%H%M%S')}.jpg"
elif foto_file is not None:
    foto_bytes = foto_file.getvalue()
    foto_nombre = foto_file.name

if foto_bytes:
    st.image(foto_bytes, caption=foto_nombre, use_container_width=True)

# Subir libro Excel (opcional)
st.subheader("Libro de informes (Excel)")
libro = st.file_uploader("Subí tu libro .xlsx para anexar (p.ej. 'Informe de prevención ... .xlsx')", type=["xlsx"], key="libro")  # noqa: E501

# Leer matriz y acciones desde la hoja TABLA RIESGO (si hay libro)
matriz_pg = None
acciones_por_categoria = None
if libro is not None:
    matriz_pg, acciones_por_categoria = leer_matriz_y_acciones(libro.getvalue())
    with st.expander("Matriz y acciones detectadas (desde tu hoja 'TABLA RIESGO')"):
        if matriz_pg:
            st.write("Matriz P x G → categoría (muestra):")
            # Mostrar una vista pivot simple
            import numpy as np
            # construir tabla 4x4
            tabla = [[matriz_pg.get((p, g), "") for g in range(1,5)] for p in range(1,5)]
            st.table(pd.DataFrame(tabla, index=["P=1","P=2","P=3","P=4"], columns=["G=1","G=2","G=3","G=4"]))
        else:
            st.warning("No se pudo leer la matriz; se usará solo GxP sin categoría.")
        if acciones_por_categoria:
            st.write("Acción y cronograma por categoría:")
            st.table(pd.DataFrame(sorted(acciones_por_categoria.items()), columns=["Categoría","Acción"]))

# ------------------------- Botón de agregar ---------------------------- #
if "_buffer" not in st.session_state:
    st.session_state._buffer: List[Observacion] = []
if "_fotos" not in st.session_state:
    st.session_state._fotos = {}

if st.button("➕ Agregar hallazgo al informe"):
    if not empresa or not ubicacion:
        st.error("Completá Empresa y Ubicación en la barra lateral.")
    elif not area or not no_conf:
        st.error("Completá Área y No conformidad.")
    else:
        categoria = None
        accion = None
        if matriz_pg:
            categoria = matriz_pg.get((int(prob), int(gravedad)))
            if categoria and acciones_por_categoria:
                accion = acciones_por_categoria.get(categoria)
        o = Observacion(
            fecha=fecha,
            empresa=empresa,
            ubicacion=ubicacion,
            area=area.strip(),
            no_conformidad=no_conf.strip(),
            descripcion=desc.strip(),
            gravedad=int(gravedad),
            probabilidad=int(prob),
            riesgo=int(riesgo),
            categoria=categoria,
            accion=accion,
            medida=medida.strip(),
            responsable=responsable.strip(),
            plazo=plazo if plazo else None,
            estado=estado,
            normativa=normativa.strip() or None,
            foto_nombre=foto_nombre,
        )
        st.session_state._buffer.append(o)
        if foto_bytes and foto_nombre:
            st.session_state._fotos[foto_nombre] = foto_bytes
        st.success("Hallazgo agregado al buffer ✔️")

# Mostrar buffer actual
if st.session_state._buffer:
    st.subheader("Hallazgos cargados")
    data_preview = pd.DataFrame([x.to_row() for x in st.session_state._buffer])
    st.dataframe(data_preview, use_container_width=True)

# ------------------------- Generar archivos ---------------------------- #
st.divider()
colg1, colg2 = st.columns([1,1])
with colg1:
    gen_xlsx = st.button("💾 Generar/actualizar Excel")
with colg2:
    gen_resumen = st.button("📝 Generar resumen (Markdown)")

# Excel
if gen_xlsx:
    if not st.session_state._buffer:
        st.error("No hay hallazgos para exportar.")
    else:
        df_rows = pd.DataFrame([x.to_row() for x in st.session_state._buffer], columns=CAMPOS_ORDEN)
        xlsx_bytes = cargar_o_crear_excel_bytes(libro.getvalue() if libro else None, df_rows)
        st.download_button(
            label="Descargar Excel",
            data=xlsx_bytes,
            file_name=f"informe_prevencion_{HOY.isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# Resumen
if gen_resumen:
    if not st.session_state._buffer:
        st.error("No hay hallazgos para resumir.")
    else:
        # Armar contexto
        por_area: Dict[str, List[Dict[str, Any]]] = {}
        for o in st.session_state._buffer:
            por_area.setdefault(o.area, []).append(o.to_row())
        ctx = {
            "fecha": fecha.isoformat(),
            "empresa": empresa,
            "ubicacion": ubicacion,
            "total": len(st.session_state._buffer),
            "por_area": por_area,
            "logo_data": logo_b64_uri,
        }
        try:
            tpl = Template(plantilla_txt)
        except Exception:
            tpl = DEFAULT_RESUMEN_TEMPLATE
        resumen_md = tpl.render(**ctx)
        st.markdown(resumen_md)
        st.download_button(
            label="Descargar resumen .md",
            data=resumen_md.encode("utf-8"),
            file_name=f"resumen_informativo_{HOY.isoformat()}.md",
            mime="text/markdown",
        )

# ------------------------- Exportar fotos (zip) ------------------------ #
if st.session_state.get("_fotos"):
    import zipfile

    if st.button("📷 Descargar fotos en ZIP"):
        bio = io.BytesIO()
        with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for nombre, contenido in st.session_state._fotos.items():
                zf.writestr(nombre, contenido)
        st.download_button(
            label="Descargar evidencias.zip",
            data=bio.getvalue(),
            file_name=f"evidencias_{HOY.isoformat()}.zip",
            mime="application/zip",
        )

# ------------------------- Ayuda rápida ------------------------------- #
with st.expander("Formato recomendado para uso por chat (para que yo redacte en el acto)"):
    st.code(
        """
OBS:
FECHA: 2025-09-07
EMPRESA: <razón social>
UBICACIÓN: <ciudad/planta>
ÁREA: <sector>
NO CONFORMIDAD: <título breve>
DESCRIPCIÓN: <hecho objetivo>
GRAVEDAD: <1-4>
PROBABILIDAD: <1-4>
MEDIDA: <acción preventiva>
RESPONSABLE: <nombre/cargo>
PLAZO: <aaaa-mm-dd>
FOTO: <archivo.jpg> (opcional)
NORMATIVA: <opcional>
        """.strip(),
        language="yaml",
    )
    st.write("Con ese bloque, te devuelvo inmediatamente el Informe (fila Excel) y el Resumen.")
streamlit
pandas
openpyxl
Pillow
jinja2
pytz

