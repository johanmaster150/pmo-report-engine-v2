#!/usr/bin/env python3
# ==========================================
# RUN SERVER — Linux Launcher
# ==========================================

import os
import sys
import traceback

from app_linux import app

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    host = os.environ.get('HOST', '0.0.0.0')
    debug = os.environ.get('DEBUG', 'false').lower() == 'true'

    try:
        print("=" * 60)
        print("  JAGA TECH Report Engine V5 — Linux Edition")
        print("=" * 60)
        print(f"  Motor:    Pipeline 3 Fases (python-docx + LibreOffice)")
        print(f"  IA:       x-ai/grok-4.1-fast via OpenRouter (Reasoning)")
        print(f"  URL:      http://{host}:{port}/")
        print(f"  Health:   http://{host}:{port}/health")
        print(f"  Platform: {sys.platform}")
        print("=" * 60)

        app.run(host=host, port=port, debug=debug, use_reloader=False)

    except Exception as e:
        print("\n" + "=" * 50)
        print("ERROR CRITICO AL INICIAR EL SERVIDOR:")
        print("=" * 50)
        print(f"Detalle del error: {e}")
        traceback.print_exc()
        print("\nPOSIBLES CAUSAS:")
        print(f"1. El puerto {port} ya esta siendo usado.")
        print("2. Otra instancia del programa esta abierta.")
        print("=" * 50)
        sys.exit(1)
