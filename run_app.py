import os
import sys
from threading import Timer
import webbrowser
from app import app
import traceback

def open_browser():
    try:
        webbrowser.open_new("http://127.0.0.1:5000/v2")
    except Exception:
        pass

if __name__ == '__main__':
    try:
        print("Iniciando JAGA TECH Report Engine V2...")
        Timer(1.5, open_browser).start()
        app.run(host="0.0.0.0", port=5000, debug=False, use_reloader=False)
    except Exception as e:
        print("\n" + "="*50)
        print("❌ ERROR CRÍTICO AL INICIAR EL SERVIDOR:")
        print("="*50)
        print(f"Detalle del error: {e}")
        traceback.print_exc()
        print("\n💡 POSIBLES CAUSAS:")
        print("1. El puerto 5000 ya está siendo usado por otra aplicación.")
        print("2. Tienes otra instancia de este programa abierta (una ventana negra igual a esta). Ciérrala primero.")
        print("="*50)
        input("Presiona la tecla ENTER para salir...")
        sys.exit(1)
