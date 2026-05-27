# JAGA TECH Report Engine V5 — Linux Edition

Motor de compilación de informes mensuales compatible con Linux ARM (Oracle Linux).

## Requisitos

- **Docker** (recomendado) o Python 3.9+ con LibreOffice instalado
- Para ARM64: Docker con soporte multi-arch (Oracle Linux, Ubuntu ARM, etc.)

## Despliegue con Docker (Recomendado)

```bash
# 1. Copiar esta carpeta al servidor
scp -r linux_server/ user@servidor:/opt/report-engine/

# 2. En el servidor, construir y ejecutar
cd /opt/report-engine
docker compose up -d --build

# 3. Verificar estado
docker compose logs -f
curl http://localhost:8080/health
```

## Despliegue sin Docker

```bash
# 1. Instalar LibreOffice
sudo dnf install -y libreoffice-writer libreoffice-headless liberation-sans-fonts liberation-serif-fonts

# 2. Instalar dependencias Python
pip3 install -r requirements.txt

# 3. Ejecutar (desarrollo)
python3 run_server.py

# 4. Ejecutar (producción con gunicorn)
gunicorn --bind 0.0.0.0:8080 --workers 2 --threads 4 --timeout 300 app_linux:app
```

## Reverse Proxy (Caddy)

Si ya tienes Caddy configurado en tu servidor:

```caddyfile
reportengine.tudominio.com {
    reverse_proxy localhost:8080
}
```

## Archivos Soportados

| Formato | Entrada | Salida |
|---------|---------|--------|
| .docx   | ✅      | ✅     |
| .doc     | ✅ (auto-convierte via LibreOffice) | ✅ |
| .pdf     | —       | ✅     |

## Endpoints

| Endpoint | Método | Descripción |
|----------|--------|-------------|
| `/` o `/v5` | GET | Interfaz web principal |
| `/api/generate_v5` | POST | Iniciar compilación |
| `/api/progress/<task_id>` | GET | Progreso de tarea |
| `/download/<filename>` | GET | Descargar archivo |
| `/health` | GET | Health check |

## Diferencias con la versión Windows

| Característica | Windows (V5) | Linux (V5) |
|---|---|---|
| Motor de Word | Win32 COM (Microsoft Word) | python-docx (pure Python) |
| PDF Export | Word COM SaveAs PDF | LibreOffice headless |
| .doc Support | Nativo (Word COM) | LibreOffice conversion |
| Servidor | Flask dev / PyInstaller | gunicorn / Docker |
| IA Grok | ✅ Idéntico | ✅ Idéntico |
| Puerto default | 5000 | 8080 |

## Arquitectura

```
┌─────────────────┐     ┌──────────────────┐     ┌───────────────┐
│  Browser (UI)   │────▶│  Flask / gunicorn │────▶│  engine_linux │
│  index_linux    │     │  app_linux.py     │     │  python-docx  │
└─────────────────┘     └──────────────────┘     └───────┬───────┘
                                │                         │
                                ▼                         ▼
                        ┌───────────────┐        ┌────────────────┐
                        │ OpenRouter API │        │  LibreOffice   │
                        │ (Grok 4.1)    │        │  --headless    │
                        └───────────────┘        └────────────────┘
```
