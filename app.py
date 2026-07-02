"""
Inventario de Chips SIM - App interna de control de inventario
================================================================
Captura fotos con la cámara del celular/laptop, detecta el ICCID por
código de barras/QR (pyzbar) o por OCR (pytesseract), evita duplicados,
muestra la lista en tiempo real y permite exportar a Excel.

Ejecutar con:
    streamlit run app.py
"""

import io
import re
import sqlite3
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image

# ----------------------------------------------------------------------
# Dependencias opcionales: la app debe seguir funcionando (modo manual)
# aunque falten librerías de sistema (zbar / tesseract) en el equipo.
# ----------------------------------------------------------------------
try:
    from pyzbar.pyzbar import decode as zbar_decode
    ZBAR_AVAILABLE = True
except Exception:
    ZBAR_AVAILABLE = False

try:
    import cv2
    CV2_AVAILABLE = True
except Exception:
    CV2_AVAILABLE = False

try:
    import pytesseract
    TESSERACT_AVAILABLE = True
except Exception:
    TESSERACT_AVAILABLE = False

DB_PATH = "inventario_sim.db"
# ICCID: empieza en 89 (industria de telecom), 19-20 dígitos en total
ICCID_REGEX = re.compile(r"89\d{16,18}")


# ----------------------------------------------------------------------
# Base de datos (SQLite) - persiste entre sesiones y reinicios de la app
# ----------------------------------------------------------------------
@st.cache_resource
def get_connection():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS chips (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            iccid TEXT UNIQUE NOT NULL,
            fecha_hora TEXT NOT NULL,
            metodo TEXT
        )
        """
    )
    conn.commit()
    return conn


def get_all(conn: sqlite3.Connection) -> pd.DataFrame:
    return pd.read_sql_query(
        "SELECT iccid AS ICCID, fecha_hora AS 'Fecha y Hora de Captura', "
        "metodo AS 'Metodo' FROM chips ORDER BY id DESC",
        conn,
    )


def iccid_exists(conn: sqlite3.Connection, iccid: str) -> bool:
    cur = conn.execute("SELECT 1 FROM chips WHERE iccid = ?", (iccid,))
    return cur.fetchone() is not None


def insert_chip(conn: sqlite3.Connection, iccid: str, metodo: str) -> None:
    conn.execute(
        "INSERT INTO chips (iccid, fecha_hora, metodo) VALUES (?, ?, ?)",
        (iccid, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), metodo),
    )
    conn.commit()


def delete_chip(conn: sqlite3.Connection, iccid: str) -> None:
    conn.execute("DELETE FROM chips WHERE iccid = ?", (iccid,))
    conn.commit()


# ----------------------------------------------------------------------
# Detección del código: 1) barcode/QR  2) OCR  3) manual
# ----------------------------------------------------------------------
def _extract_iccid_from_digits(digits: str) -> str | None:
    m = ICCID_REGEX.search(digits)
    if m:
        return m.group(0)
    if len(digits) >= 15:
        return digits
    return None


def try_barcode(image_pil: Image.Image) -> str | None:
    if not ZBAR_AVAILABLE:
        return None
    img = np.array(image_pil.convert("L"))
    for result in zbar_decode(img):
        data = result.data.decode("utf-8", errors="ignore")
        digits = re.sub(r"\D", "", data)
        code = _extract_iccid_from_digits(digits)
        if code:
            return code
    return None


def _preprocess_for_ocr(image_pil: Image.Image):
    img = np.array(image_pil.convert("L"))
    if CV2_AVAILABLE:
        img = cv2.GaussianBlur(img, (3, 3), 0)
        _, img = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return img


def try_ocr(image_pil: Image.Image) -> str | None:
    if not TESSERACT_AVAILABLE:
        return None
    processed = _preprocess_for_ocr(image_pil)
    config = "--psm 6 -c tessedit_char_whitelist=0123456789"
    try:
        text = pytesseract.image_to_string(processed, config=config)
    except Exception:
        return None
    digits = re.sub(r"\D", "", text)
    m = ICCID_REGEX.search(digits)
    return m.group(0) if m else None


def detect_code(image_pil: Image.Image):
    code = try_barcode(image_pil)
    if code:
        return code, "Código de barras/QR"
    code = try_ocr(image_pil)
    if code:
        return code, "OCR"
    return None, "Manual"


# ----------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------
st.set_page_config(page_title="Inventario Chips SIM", page_icon="📶", layout="centered")

conn = get_connection()

if "detected_code" not in st.session_state:
    st.session_state.detected_code = ""
if "detected_metodo" not in st.session_state:
    st.session_state.detected_metodo = "Manual"
if "input_key" not in st.session_state:
    st.session_state.input_key = 0

st.title("📶 Inventario de Chips SIM")

total = conn.execute("SELECT COUNT(*) FROM chips").fetchone()[0]
st.metric("Total de chips registrados", total)

if not ZBAR_AVAILABLE or not TESSERACT_AVAILABLE:
    faltantes = []
    if not ZBAR_AVAILABLE:
        faltantes.append("lectura de código de barras/QR (zbar)")
    if not TESSERACT_AVAILABLE:
        faltantes.append("OCR (tesseract)")
    st.warning(
        "⚠️ No están disponibles: " + " y ".join(faltantes) +
        ". Puedes seguir usando la app con captura/edición manual del ICCID. "
        "Revisa el README para instalar las dependencias de sistema."
    )

st.subheader("1. Escanear")
img_file = st.camera_input("Toma una foto del ICCID / código de barras del chip")

if img_file is not None:
    image = Image.open(img_file)
    with st.spinner("Analizando imagen..."):
        code, metodo = detect_code(image)
    if code:
        st.success(f"Código detectado ({metodo}): {code}")
    else:
        st.warning("No se detectó el código automáticamente. Ingrésalo manualmente abajo.")
    st.session_state.detected_code = code or ""
    st.session_state.detected_metodo = metodo

iccid_input = st.text_input(
    "ICCID (verifica o corrige antes de guardar)",
    value=st.session_state.detected_code,
    max_chars=22,
    key=f"iccid_input_{st.session_state.input_key}",
)

col1, col2 = st.columns(2)
with col1:
    guardar = st.button("✅ Guardar registro", use_container_width=True, type="primary")
with col2:
    limpiar = st.button("🔄 Limpiar", use_container_width=True)

if guardar:
    clean = re.sub(r"\D", "", iccid_input)
    if not clean:
        st.error("Ingresa un ICCID válido.")
    elif len(clean) < 15:
        st.error("El ICCID parece incompleto (muy corto). Verifica la foto o el texto.")
    elif iccid_exists(conn, clean):
        st.error(f"⚠️ Este chip ya fue escaneado antes: {clean}")
    else:
        insert_chip(conn, clean, st.session_state.detected_metodo)
        st.success(f"Chip guardado: {clean}")
        st.session_state.detected_code = ""
        st.session_state.input_key += 1
        st.rerun()

if limpiar:
    st.session_state.detected_code = ""
    st.session_state.input_key += 1
    st.rerun()

st.divider()
st.subheader("2. Chips escaneados")

df = get_all(conn)
st.dataframe(df, use_container_width=True, hide_index=True)

st.divider()
st.subheader("3. Exportar")

if len(df) > 0:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Inventario SIM")
    buffer.seek(0)
    st.download_button(
        label="⬇️ Descargar Excel",
        data=buffer,
        file_name=f"inventario_sim_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
else:
    st.info("Aún no hay chips escaneados.")

with st.expander("🗑️ Eliminar un registro (correcciones)"):
    if len(df) > 0:
        to_delete = st.selectbox("Selecciona el ICCID a eliminar", df["ICCID"].tolist())
        if st.button("Eliminar registro seleccionado"):
            delete_chip(conn, to_delete)
            st.rerun()
    else:
        st.caption("No hay registros para eliminar.")

st.caption(
    "Los datos se guardan localmente en `inventario_sim.db` en el equipo donde corre la app."
)
