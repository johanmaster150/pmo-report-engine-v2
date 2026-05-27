# ==========================================
# APP LINUX — Flask Application (Cross-Platform)
# ==========================================
# Self-contained Flask app for Linux/ARM deployment.
# No Windows dependencies (win32com, pythoncom).

import os
import re
import sys
import uuid
import threading
import traceback

import requests
import docx
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename

from engine_linux import (
    compile_unified_v5_linux,
    build_ai_docx,
    read_document_text,
    convert_docx_to_format,
    ensure_docx,
    KNOWN_PROD_IDS,
)

# ============================================
# APP SETUP
# ============================================

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB max

# ============================================
# GLOBAL TASK STATE
# ============================================

global_tasks = {}


def update_progress(task_id, progress, status):
    if task_id and task_id in global_tasks:
        global_tasks[task_id]['progress'] = progress
        global_tasks[task_id]['status'] = status


# ============================================
# UTILITY: CTI PATH CHECK
# ============================================

def check_cti_path(url):
    """Check if a URL belongs to the CTI strategic projects path."""
    from urllib.parse import unquote
    decoded_url = unquote(url).lower()
    cti_indicators = [
        'proyectos estratégicos cti', 'proyectos estrategicos cti',
        'proyectos%20estrat', '03productos', '03 productos',
        'documentos'
    ]
    matches = sum(1 for ind in cti_indicators if ind in decoded_url)
    if matches >= 2:
        return "✅ Sí - En ruta CTI"
    elif matches == 1:
        return "⚠️ Posible - Ruta parcial CTI"
    return "❌ No"


# ============================================
# AI FUNCTIONS (GROK 4.1 FAST + REASONING)
# ============================================

GROK_API_KEY = os.environ.get("OPENROUTER_API_KEY", "")
GROK_MODEL = "x-ai/grok-4.1-fast"


def _call_grok(prompt, task_id=None, status_msg="IA Grok: Procesando..."):
    """Call Grok 4.1 Fast via OpenRouter with reasoning mode."""
    update_progress(task_id, 15, status_msg)
    try:
        response = requests.post(
            url="https://openrouter.ai/api/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {GROK_API_KEY}",
                "Content-Type": "application/json"
            },
            json={
                "model": GROK_MODEL,
                "reasoning": {"effort": "high"},
                "messages": [{"role": "user", "content": prompt}]
            },
            timeout=180
        )
        if response.status_code == 200:
            data = response.json()
            content = data['choices'][0]['message']['content'].strip()
            return content
        else:
            err = f"API Grok Falló ({response.status_code}): {response.text[:500]}"
            if task_id:
                global_tasks[task_id]['error_ai'] = err
            print(err)
            return None
    except Exception as e:
        err = f"Error IA Grok: {e}"
        if task_id:
            global_tasks[task_id]['error_ai'] = err
        print(err)
        return None


def _read_project_texts(project_files, task_id=None, start_pct=2, end_pct=12):
    """Read text from project files for AI context."""
    combined_text = ""
    total = len(project_files)
    for i, file_path in enumerate(project_files):
        pct = start_pct + int((end_pct - start_pct) * (i / total))
        update_progress(task_id, pct, f"IA Grok: Leyendo contexto de {os.path.basename(file_path)}...")
        text = read_document_text(file_path)
        combined_text += f"\n--- Proyecto: {os.path.basename(file_path)} ---\n"
        combined_text += text
    return combined_text


def _read_unified_document_text(unified_docx_path, task_id=None, start_pct=2, end_pct=12):
    """Read text from the unified document for complete AI analysis."""
    update_progress(task_id, start_pct, "IA Grok: Leyendo documento unificado para análisis...")
    text = read_document_text(unified_docx_path, max_chars=80000)
    update_progress(task_id, end_pct, f"IA Grok: Documento unificado leído...")
    return text


def generate_executive_summary_v5(project_files, task_id=None, unified_docx_path=None):
    """Generate executive summary + audit with Grok 4.1 Fast."""
    update_progress(task_id, 2, "IA Grok: Iniciando análisis del documento unificado...")

    if unified_docx_path and os.path.exists(unified_docx_path):
        combined_text = _read_unified_document_text(unified_docx_path, task_id, 2, 10)
    else:
        combined_text = _read_project_texts(project_files, task_id, 2, 10)

    products_list = ", ".join(KNOWN_PROD_IDS)

    prompt = (
        "Eres el Gerente Senior de la PMO (Oficina de Gestión de Proyectos) de la Alcaldía de Medellín, "
        "con experiencia en auditoría de proyectos de innovación, ciencia, tecnología e industria 4.0.\n\n"
        "A continuación recibirás el texto completo del INFORME MENSUAL UNIFICADO que consolida "
        "los reportes de avance de todos los equipos técnicos del mes.\n\n"
        "DEBES generar un documento con la siguiente estructura EXACTA:\n\n"
        "===================================================\n"
        "PARTE 1: RESUMEN EJECUTIVO GENERAL\n"
        "===================================================\n"
        "Escribe un resumen ejecutivo consolidado de 2 a 3 páginas que incluya:\n"
        "- Visión general del estado del portafolio de proyectos\n"
        "- Logros principales del mes a nivel agregado\n"
        "- Principales riesgos y alertas identificados transversalmente\n"
        "- Estado general de cumplimiento de metas\n"
        "- Conclusiones y recomendaciones estratégicas para la alta dirección\n\n"
        "===================================================\n"
        "PARTE 2: AUDITORÍA GENERAL DEL INFORME\n"
        "===================================================\n"
        "Realiza una auditoría transversal del informe completo evaluando:\n"
        "- Coherencia entre los diferentes productos y proyectos\n"
        "- Consistencia en formatos, métricas y forma de reportar\n"
        "- Vacíos de información detectados a nivel global\n"
        "- Duplicidades o contradicciones entre secciones\n"
        "- Calidad general de la evidencia presentada\n"
        "- Cumplimiento de estándares de reporte esperados\n\n"
        "===================================================\n"
        f"PARTE 3: AUDITORÍA DETALLADA POR PRODUCTO ({products_list})\n"
        "===================================================\n"
        f"Para CADA UNO de los 12 productos ({products_list}), genera una sección separada con:\n\n"
        "## Producto X.Y: [Nombre del producto]\n\n"
        "a) RESUMEN DEL AVANCE\nb) HALLAZGOS POSITIVOS\nc) OBSERVACIONES CRÍTICAS\n"
        "d) CALIFICACIÓN SEMÁFORO (VERDE/AMARILLO/ROJO)\ne) RECOMENDACIONES ESPECÍFICAS\n\n"
        "===================================================\n"
        "INSTRUCCIONES DE FORMATO:\n"
        "===================================================\n"
        "- Escribe en español de Colombia corporativo, formal y directo.\n"
        "- Usa los numerales 1.1, 1.2... 3.1 como títulos claros.\n"
        "- NO uses asteriscos dobles **. Usa texto plano bien espaciado.\n"
        f"- Los 12 productos son: {products_list}\n\n"
        f"DOCUMENTO UNIFICADO COMPLETO:\n{combined_text[:80000]}"
    )

    update_progress(task_id, 15, "IA Grok: Enviando documento completo a Grok 4.1 Fast (reasoning)...")
    ai_text = _call_grok(prompt, task_id, "IA Grok: Procesando auditoría ejecutiva con reasoning...")

    if not ai_text:
        return None

    update_progress(task_id, 35, "IA Grok: Estructurando reporte de auditoría...")

    output_docx = "00_Resumen_y_Auditoria_IA_V5.docx"
    output_pdf = "00_Resumen_y_Auditoria_IA_V5.pdf"
    docx_path = os.path.join(OUTPUT_FOLDER, output_docx)
    pdf_path = os.path.join(OUTPUT_FOLDER, output_pdf)

    build_ai_docx(ai_text, 'Resumen Ejecutivo y Auditoría por Producto (IA Grok V5)', docx_path)

    update_progress(task_id, 38, "IA Grok: Convirtiendo auditoría a PDF...")
    convert_docx_to_format(docx_path, 'pdf', OUTPUT_FOLDER)

    print("¡Auditoría IA Grok V5 completada!")
    update_progress(task_id, 40, "IA Grok: ¡Auditoría construida! Pasando al ensamblador...")
    return {"docx": output_docx, "pdf": output_pdf}


def generate_predictive_report_v5(project_files, task_id=None, unified_docx_path=None):
    """Generate predictive report with Grok 4.1 Fast."""
    update_progress(task_id, 20, "IA Grok Predictiva: Preparando análisis del documento unificado...")

    if unified_docx_path and os.path.exists(unified_docx_path):
        combined_text = _read_unified_document_text(unified_docx_path, task_id, 20, 25)
    else:
        combined_text = _read_project_texts(project_files, task_id, 20, 25)

    products_list = ", ".join(KNOWN_PROD_IDS)

    prompt = (
        "Eres un consultor senior de PMO experto en proyectos de "
        "innovación, ciencia, tecnología e industria 4.0 para la Alcaldía de Medellín.\n\n"
        "A continuación recibirás el texto completo del INFORME MENSUAL UNIFICADO de avances.\n\n"
        "DEBES generar un INFORME PREDICTIVO Y DE RECOMENDACIONES con esta estructura EXACTA:\n\n"
        "PARTE 1: ANÁLISIS PREDICTIVO GENERAL DEL PORTAFOLIO\n"
        f"PARTE 2: ANÁLISIS PREDICTIVO POR PRODUCTO ({products_list})\n"
        "PARTE 3: MATRIZ DE PRIORIDADES\n\n"
        "- Escribe en español de Colombia corporativo.\n"
        "- NO uses asteriscos dobles **.\n"
        f"- Los 12 productos son: {products_list}\n\n"
        f"DOCUMENTO UNIFICADO COMPLETO:\n{combined_text[:80000]}"
    )

    update_progress(task_id, 27, "IA Grok Predictiva: Enviando a Grok 4.1 Fast (reasoning)...")
    ai_text = _call_grok(prompt, task_id, "IA Grok Predictiva: Procesando informe con reasoning...")

    if not ai_text:
        return None

    update_progress(task_id, 33, "IA Grok Predictiva: Construyendo informe predictivo...")

    output_docx = "01_Informe_Predictivo_IA_V5.docx"
    output_pdf = "01_Informe_Predictivo_IA_V5.pdf"
    docx_path = os.path.join(OUTPUT_FOLDER, output_docx)
    pdf_path = os.path.join(OUTPUT_FOLDER, output_pdf)

    build_ai_docx(ai_text, 'Informe Predictivo y Recomendaciones por Producto (IA Grok V5)', docx_path)

    update_progress(task_id, 36, "IA Grok Predictiva: Convirtiendo a PDF...")
    convert_docx_to_format(docx_path, 'pdf', OUTPUT_FOLDER)

    print("¡Informe Predictivo Grok V5 completado!")
    update_progress(task_id, 38, "IA Grok Predictiva: ¡Informe construido!")
    return {"docx": output_docx, "pdf": output_pdf}


# ============================================
# WORKER THREAD
# ============================================

def worker_compilation_v5(task_id, master_path, project_files, output_path,
                          font_family, generate_ai, mode, export_format):
    try:
        # STEP 1: Compile the unified document
        saved_files = compile_unified_v5_linux(
            master_path, project_files, output_path,
            font_family, export_format, task_id,
            update_fn=update_progress,
            check_cti_fn=check_cti_path,
            output_folder=OUTPUT_FOLDER
        )

        # STEP 2: Generate AI reports on the unified document
        ai_filenames = None
        ai_predictive_filenames = None
        if generate_ai and saved_files:
            unified_docx_path = output_path if os.path.exists(output_path) else None
            update_progress(task_id, 2, "IA Grok: Preparando análisis del documento unificado...")
            ai_filenames = generate_executive_summary_v5(
                project_files, task_id, unified_docx_path=unified_docx_path
            )
            ai_predictive_filenames = generate_predictive_report_v5(
                project_files, task_id, unified_docx_path=unified_docx_path
            )

        if saved_files:
            download_urls = {}
            if "docx" in saved_files:
                download_urls["docx"] = f"/download/{saved_files['docx']}"
            if "pdf" in saved_files:
                download_urls["pdf"] = f"/download/{saved_files['pdf']}"
            if "doc" in saved_files:
                download_urls["doc"] = f"/download/{saved_files['doc']}"
            if "csv" in saved_files:
                download_urls["csv"] = f"/download/{saved_files['csv']}"
            if "links_docx" in saved_files:
                download_urls["links_docx"] = f"/download/{saved_files['links_docx']}"

            global_tasks[task_id]['downloadUrls'] = download_urls

            if ai_filenames:
                global_tasks[task_id]['aiDownloadUrls'] = {
                    fmt: f"/download/{fname}" for fmt, fname in ai_filenames.items()
                }
            if ai_predictive_filenames:
                global_tasks[task_id]['aiPredictiveUrls'] = {
                    fmt: f"/download/{fname}" for fmt, fname in ai_predictive_filenames.items()
                }

            global_tasks[task_id]['success'] = True
        else:
            global_tasks[task_id]['error'] = "Falló la compilación V5 del documento."
    except Exception as e:
        global_tasks[task_id]['error'] = f"Excepción crítica durante worker V5: {e}"
        traceback.print_exc()
    finally:
        global_tasks[task_id]['completed'] = True


# ============================================
# FLASK ROUTES
# ============================================

@app.route("/", methods=["GET"])
@app.route("/v5", methods=["GET"])
def index_v5():
    return render_template("index_linux.html")


@app.route("/api/progress/<task_id>")
def get_progress(task_id):
    if task_id in global_tasks:
        return jsonify(global_tasks[task_id])
    return jsonify({"error": "Task not found"}), 404


@app.route("/api/generate_v5", methods=["POST"])
def generate_report_v5():
    print("Received async generation request V5 (Linux)")

    task_id = str(uuid.uuid4())

    global_tasks[task_id] = {
        "progress": 0,
        "status": "Iniciando motor V5 Linux...",
        "completed": False,
        "success": False,
        "error": None,
        "error_ai": None,
        "downloadUrl": None,
        "aiDownloadUrl": None
    }

    try:
        if 'masterTemplate' not in request.files:
            return jsonify({"error": "No has subido la plantilla maestra"}), 400

        master_file = request.files['masterTemplate']
        original_ext = os.path.splitext(master_file.filename)[1].lower()
        master_save_name = f'{task_id}_master_template{original_ext}'
        master_path = os.path.join(UPLOAD_FOLDER, master_save_name)
        master_file.save(master_path)

        # Convert .doc to .docx if needed
        if original_ext == '.doc':
            converted = ensure_docx(master_path, UPLOAD_FOLDER)
            if converted:
                master_path = converted

        project_files = []
        ordered_keys = request.form.getlist('projectIds[]')
        for pk in ordered_keys:
            proj_file = request.files.get(f'projectFile_{pk}')
            if proj_file:
                safe_name = secure_filename(proj_file.filename)
                save_path = os.path.join(UPLOAD_FOLDER, f'{task_id}_{safe_name}')
                proj_file.save(save_path)

                # Convert .doc to .docx if needed
                if save_path.lower().endswith('.doc'):
                    converted = ensure_docx(save_path, UPLOAD_FOLDER)
                    if converted:
                        save_path = converted

                project_files.append(save_path)

        if not project_files:
            return jsonify({"error": "No se subieron archivos de proyecto válidos"}), 400

        font_family = request.form.get("fontFamily", "default")
        mode = request.form.get("mode", "unified")
        generate_ai = request.form.get("generateAI") == "true"
        export_format = request.form.get("exportFormat", "word")

        output_filename = f"Reporte_Mensual_V5_{task_id[:8]}.docx"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        thread = threading.Thread(
            target=worker_compilation_v5,
            args=(task_id, master_path, project_files, output_path,
                  font_family, generate_ai, mode, export_format)
        )
        thread.start()

        return jsonify({"taskId": task_id, "status": "Task Started V5 (Linux)"})

    except Exception as e:
        print(f"Error procesando peticion POST V5 Linux: {e}")
        return jsonify({"error": f"Error al procesar archivos: {str(e)}"}), 500


@app.route("/download/<filename>")
def download_file(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    return jsonify({"error": "File not found"}), 404


# ============================================
# HEALTH CHECK
# ============================================

@app.route("/health")
def health():
    """Health check endpoint for monitoring."""
    import shutil
    lo_available = shutil.which("libreoffice") is not None
    return jsonify({
        "status": "ok",
        "platform": sys.platform,
        "libreoffice": lo_available,
        "python": sys.version,
        "engine": "V5 Linux (python-docx + LibreOffice Headless)"
    })
