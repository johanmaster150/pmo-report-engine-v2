# JAGA TECH Report Engine V2 📊🤖
### Sistema de Unificación de Informes PMO — Alcaldía de Medellín

Este proyecto es una herramienta web local diseñada para unificar, dar formato y enriquecer los informes mensuales de avance de los diferentes equipos y proyectos de la PMO de la Alcaldía de Medellín, con un enfoque particular en la **Versión V2 (Consolidación Completa con Autoajuste, Índice e Inteligencia Artificial)**.

---

## 🚀 Versiones Soportadas en la Aplicación
El sistema incluye diferentes interfaces y lógicas según la necesidad del reporte:
1. **Versiones 1 y 2 (Panel V2 - `/v2`):** Es la versión recomendada y más avanzada. Realiza:
   - Unión de archivos Word de insumos respetando formatos y estilos de la plantilla maestra.
   - Autoajuste de todas las tablas dentro de los documentos.
   - Detección dinámica y renumeración de productos (secciones 1.1 a 1.7, 2.1 a 2.3, y 3.1) y sus tablas asociadas.
   - Generación de un **Índice (Tabla de Contenido) personalizado y dinámico** basado en los títulos encontrados.
   - Análisis inteligente de texto mediante Inteligencia Artificial (a través de **OpenRouter** con el modelo `arcee-ai/trinity-large-preview:free`) para generar automáticamente un **Resumen Ejecutivo** y un **Informe Predictivo de Avance**.
   - Extracción automática de enlaces (URLs de evidencias) y descarga de archivos asociados en lote.
2. **Versiones 3 y 4 (Formato Estricto - `/`):** Diseñadas para consolidaciones que exigen un apego estricto a plantillas sin variaciones de estilo.
3. **Versión V5 (Linux Server):** Ubicada en la carpeta `linux_server/`, diseñada para ejecutarse dentro de contenedores **Docker** bajo sistemas operativos basados en Linux (remplazando la automatización de MS Word con LibreOffice headless y Python puro).

---

## 📌 Requerimientos del Sistema
Para la ejecución nativa en Windows (versión V2 con automatización de Word):
- **Sistema Operativo:** Windows 10/11.
- **Software Requerido:** Microsoft Office (Word) instalado localmente (ya que el backend levanta una instancia COM de Word para procesar el diseño original de manera nativa).
- **Lenguaje:** Python 3.10 o superior.

### Librerías principales (entorno virtual):
- `Flask` (Servidor Web)
- `pywin32` (Automatización COM de Word)
- `python-docx` (Manipulación de documentos)
- `requests` (Llamadas a la API de OpenRouter)
- `waitress` (Servidor WSGI de producción)

---

## 🛠️ Paso a Paso para Instalación y Ejecución

### 1. Clonar el repositorio
```bash
git clone https://github.com/johanmaster150/pmo-report-engine-v2.git
cd pmo-report-engine-v2
```

### 2. Configurar el Entorno Virtual y dependencias
Crea y activa un entorno virtual de Python, luego instala las dependencias necesarias:
```powershell
python -m venv .venv
.venv\Scripts\activate
pip install flask requests python-docx pypiwin32 waitress
```

### 3. Configurar la clave de la IA (OpenRouter)
Establece tu API Key de OpenRouter como variable de entorno (necesaria para el Resumen Ejecutivo e Informe Predictivo de la V2):
```powershell
$env:OPENROUTER_API_KEY="tu_api_key_aqui"
```

### 4. Iniciar la aplicación
Ejecuta el script de inicio:
```powershell
python run_app.py
```
El script buscará automáticamente el puerto 5000 e iniciará la aplicación, abriendo tu navegador en `http://127.0.0.1:5000/v2`.

---

## 📦 Empaquetado (Crear Ejecutable `.exe` para Windows)
Si realizas modificaciones al código y deseas generar un nuevo ejecutable portable para que los usuarios finales no necesiten tener Python instalado, puedes compilarlo usando **PyInstaller**:

1. Asegúrate de tener PyInstaller instalado en tu entorno virtual:
   ```bash
   pip install pyinstaller
   ```
2. Ejecuta el comando de compilación:
   ```powershell
   Remove-Item -Recurse -Force dist, build -ErrorAction SilentlyContinue
   pyinstaller --noconfirm --onedir --windowed --add-data "templates;templates" --add-data "static;static" run_app.py --name InformeMensualAlcaldiaApp
   ```
3. El resultado compilado estará disponible en la carpeta: `dist/InformeMensualAlcaldiaApp/`.

---

## 📖 Documentación para Desarrolladores
Para un desglose técnico completo de cómo funciona el motor por debajo (Win32COM, asincronismo mediante polling, etc.), consulta el archivo [DOCUMENTACION.md](DOCUMENTACION.md).
