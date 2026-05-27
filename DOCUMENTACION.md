# Documentación: Sistema de Unificación de Informes PMO

Este documento explica cómo está estructurada la aplicación web local para la unificación de los informes mensuales de la Alcaldía de Medellín, cómo funciona el código por debajo y los pasos necesarios si deseas realizar modificaciones y volver a generar el ejecutable (empaquetado).

---

## 1. ¿Qué copiar para mover el proyecto a otro equipo?

Dependiendo de tu necesidad, debes copiar una u otra carpeta:

**A. Solo para EJECUTAR la aplicación en otro PC (Usuario final):**
Debes copiar TODA la carpeta generada por el compilador, ubicada en:
`dist/InformeMensualAlcaldiaApp/`
(Todo el contenido de esa carpeta es necesario, ya que incluye las librerías de Python empaquetadas y el `InformeMensualAlcaldiaApp.exe`). El otro PC debe tener Windows y Microsoft Office (Word) instalado.

**B. Para MODIFICAR el código fuente en otro PC (Modo Programador):**
Debes copiar toda la carpeta raíz **`Web App`** (que contiene los archivos `.py`, `templates`, `static`, etc.). Ignora o elimina las carpetas `build` y `dist` para que pese menos al copiarla, ya que las volverás a generar al compilar.

---

## 2. Estructura del Código Fuente

*   **`run_app.py`**: Es el archivo de entrada (Entry point) principal del ejecutable. Su función es encontrar un puerto libre, iniciar el servidor web usando `waitress` (para producción), y abrir automáticamente el navegador web apuntando a la dirección local correcta (ej: `http://127.0.0.1:5000`).
*   **`app.py`**: Contiene toda la lógica principal (Backend) de la aplicación, incluyendo:
    *   Rutas de Flask (`/`, `/v2`, `/api/generate`, `/download/...`).
    *   Lógica de procesamiento de Word usando **Win32com** (Unión, inserción de páginas, bordes a tablas, renombrado, autoajuste).
    *   Lógica de llamadas a la IA Generativa (OpenRouter) para el Informe Predictivo y el Resumen Ejecutivo.
    *   Lógica de descarga y validación de las URLs extraídas del Word.
*   **Carpeta `templates/`**: Contiene los archivos HTML de la interfaz gráfica de usuario.
    *   `index.html`: (Versiones 3 y 4 - Formato Estricto).
    *   `index_v2.html`: (Versiones 1 y 2 - Consolidación Completa con Autoajuste, Índice e IA).
*   **Carpeta `static/`**: Contiene los archivos estáticos como CSS, JS, e Imágenes (logos) usados por las plantillas HTML.
*   **Carpeta `uploads/`**: Directorio temporal donde se guardan temporalmente las partes o insumos subidos al sistema.
*   **Carpeta `outputs/`**: Directorio donde se guardan los archivos fusionados resultantes (Word, PDF, CSV de URLs).

---

## 3. Funcionamiento Técnico Relevante

### Interacción con Microsoft Word (Win32COM)
El corazón del reporteador no es una simple lectura de texto. Por la complejidad visual solicitada, se construyó utilizando automatización de Office (`win32com.client`).
**Importante:** 
*   Esto significa que `app.py` levanta una instancia invisible de MS Word en el fondo.
*   Esto hace que el sistema sea **exclusivo para el Sistema Operativo Windows** que tenga Microsoft Office (Word) instalado. No funcionará en Linux (Ubuntu) originariamente por limitaciones de las Macros de COM.

### Tareas Asíncronas (Progreso en vivo)
Como la consolidación e inteligencia artificial tardan cierto tiempo, los endpoints principales de guardado devuelven un identificador (`task_id`). La interfaz web (Javascript) utiliza este ID para hacer consultas regulares ('Long Polling') al Backend pidiendo actualizaciones (`update_progress`), que es lo que mueve la barra de carga en pantalla.

### Inteligencia Artificial
La aplicación extrae el texto (sin formatos) de cada reporte, lo concatena hasta un límite seguro y se lo envía a **OpenRouter** para el modelo de IA: `arcee-ai/trinity-large-preview:free`. La petición está construida dentro de las funciones `generate_executive_summary` y `generate_predictive_report` de `app.py`. En aquellas funciones está definido el *Prompt* general. Si luego requieres usar un modelo pago o distinto (como OpenAI u otro endpoint), ese es el lugar exacto.

---

## 4. ¿Cómo hacer modificaciones en el futuro?

Si vas a agregar un nuevo botón a la interfaz, modificar el CSS, o alterar los Prompts de la IA:

1.  Abre los archivos pertinentes (ej: `templates/index_v2.html` o `app.py`) en tu editor de código preferido (ej. VS Code).
2.  Realiza tus ajustes.
3.  **Para probar:** Ejecuta con Python simple: `python run_app.py`  (OJO: Tu entorno de Python debe tener todas las librerías del proyecto. Idealmente hacer `pip install flask requests python-docx pypiwin32 waitress`).
4.  Comprueba en tu navegador que los cambios hayan surtido efecto.
5.  Cierra la consola de ejecución (mata el proceso python).

## 5. Volver a "Empaquetar" (Crear nuevo EXE)

Una vez que comprobaste que tus cambios en el código fuente funcionan perfectamente, necesitas reemplazar tu ejecutable viejo.

Abre la consola de Windows (PowerShell/CMD) en la carpeta `Web App` y presiona ENTER para borrar cachés, e iniciar nuevamente a PyInstaller con la siguiente sintaxis:

```powershell
# 1. Elimina distribuciones viejas (Solo PowerShell)
Remove-Item -Recurse -Force dist, build -ErrorAction SilentlyContinue

# 2. Compilar con PyInstaller
pyinstaller --noconfirm --onedir --windowed --add-data "templates;templates" --add-data "static;static" run_app.py --name InformeMensualAlcaldiaApp
```

### Explicación del comando de Empaquetado:
*   `--onedir`: Genera una carpeta en lugar de solo un gigante .exe (esto soluciona problemas de carga y lentitud con antivirus).
*   `--windowed`: Indica que NO debe mostrar la consola negra al usuario final.
*   `--add-data`: Incluye las carpetas HTML, CSS y logos de forma interna y las enlaza virtualmente en Windows para que PyInstaller sepa llevarse los archivos web con él.
*   `--name`: Nombre final del fichero (`InformeMensualAlcaldiaApp.exe`).

Tras un par de minutos, si todo sale con "Exit code: 0", encontrarás tu nueva aplicación lista para compartir en `dist/InformeMensualAlcaldiaApp/`.
