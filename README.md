# Inventario de Chips SIM

App interna (Streamlit) para escanear ICCID de chips SIM con la cámara, evitar duplicados y exportar el inventario a Excel. Pensada para uso en campo con 500-1000 chips.

## Cómo funciona

1. **Escanear**: toma una foto con la cámara del celular o laptop.
2. La app intenta leer el código en este orden:
   - Código de barras / QR (librería `pyzbar`)
   - OCR de los dígitos impresos (librería `pytesseract`)
   - Si ninguno detecta nada, puedes escribir el ICCID a mano.
3. El código detectado aparece en un campo editable — siempre revisa/corrige antes de guardar.
4. Al guardar, si el ICCID ya existe se muestra una alerta y no se duplica.
5. La tabla y el contador se actualizan en tiempo real.
6. Botón para descargar todo el inventario en Excel (`.xlsx`) con columnas ICCID, Fecha y Hora de Captura, y Método.
7. Los datos quedan guardados en un archivo local `inventario_sim.db` (SQLite), así que no se pierden si cierras la app o el navegador.

## Requisitos previos

- Python 3.10 o superior
- Dos dependencias de **sistema** (no se instalan con pip):
  - **Tesseract OCR** (para lectura de texto)
  - **ZBar** (para lectura de código de barras/QR)

Si no instalas estas dos, la app sigue funcionando en modo 100% manual (escribes el ICCID tú mismo).

### Instalar dependencias de sistema

**Windows**
- Tesseract: descarga el instalador desde https://github.com/UB-Mannheim/tesseract/wiki e instálalo (ruta típica `C:\Program Files\Tesseract-OCR\tesseract.exe`). Si `pytesseract` no lo encuentra automáticamente, agrega al inicio de `app.py`:
  ```python
  pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
  ```
- ZBar: `pyzbar` en Windows ya incluye las DLL necesarias. Si falla, instala el "Visual C++ Redistributable 2013" de Microsoft.

**macOS**
```bash
brew install tesseract zbar
```

**Linux (Debian/Ubuntu)**
```bash
sudo apt-get update
sudo apt-get install -y tesseract-ocr libzbar0
```

## Instalación y ejecución local

```bash
# 1. Crear entorno virtual (recomendado)
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

# 2. Instalar dependencias Python
pip install -r requirements.txt

# 3. Ejecutar la app
streamlit run app.py
```

Streamlit abrirá el navegador en `http://localhost:8501`.

### Usar desde el celular en la misma red local

1. Corre la app en tu laptop con:
   ```bash
   streamlit run app.py --server.address 0.0.0.0
   ```
2. Averigua la IP local de la laptop (`ipconfig` en Windows / `ifconfig` o `ip a` en Mac/Linux).
3. En el celular (conectado al mismo Wi-Fi), abre `http://<IP-de-la-laptop>:8501`.
4. El navegador pedirá permiso de cámara — acéptalo. Usa Chrome o Safari actualizados.

> Nota: los navegadores solo permiten acceso a la cámara en `localhost` o en conexiones HTTPS. Si accedes por IP local en HTTP, algunos navegadores (Chrome en Android) lo permiten para redes locales, pero si tienes problemas, usa el despliegue en la nube (HTTPS) descrito abajo.

## Despliegue en Streamlit Community Cloud (recomendado para uso en campo)

1. Sube esta carpeta (`app.py`, `requirements.txt`, `packages.txt`) a un repositorio de GitHub.
2. Entra a https://share.streamlit.io y conecta tu cuenta de GitHub.
3. Selecciona "New app", elige el repo, la rama y `app.py` como archivo principal.
4. Streamlit Cloud instalará automáticamente `requirements.txt` y, gracias al archivo `packages.txt`, también `tesseract-ocr` y `libzbar0`.
5. Obtendrás una URL pública HTTPS — se puede abrir desde cualquier celular con cámara sin instalar nada.

**Importante sobre persistencia en la nube**: Streamlit Community Cloud no garantiza almacenamiento permanente del archivo `inventario_sim.db` entre reinicios del contenedor. Para uso en campo con 500-1000 chips en una sola jornada esto normalmente no es problema, pero si necesitas persistencia a largo plazo en la nube, exporta el Excel frecuentemente o considera conectar una base de datos externa (por ejemplo, una hoja de Google Sheets o una base Postgres) — puedo ayudarte a adaptarlo si lo necesitas.

## Estructura de archivos

```
app.py             # Aplicación principal
requirements.txt   # Dependencias Python
packages.txt        # Dependencias de sistema para Streamlit Cloud
inventario_sim.db   # Base de datos local (se crea automáticamente al usar la app)
```

## Solución de problemas

- **"No están disponibles: lectura de código de barras..."**: falta instalar `tesseract-ocr` y/o `libzbar0` en el sistema (ver arriba). La app sigue funcionando en modo manual.
- **La cámara no abre**: revisa permisos del navegador y que estés en `localhost` o HTTPS.
- **El OCR lee mal los números**: acerca la cámara, evita reflejos/sombras, y siempre revisa el campo antes de guardar — está diseñado para corrección manual rápida.
- **Quiero borrar toda la base y empezar de cero**: cierra la app y elimina el archivo `inventario_sim.db`.
