# -*- coding: utf-8 -*-
"""desktop.py
============================================================
Lanzador de escritorio para Derma Essenza.
Inicia el servidor FastAPI en un hilo de fondo, abre el navegador
predeterminado del usuario y muestra una pequeña ventana de control
hecha en Tkinter (nativo de Python) para detener o abrir el panel.
"""
import os
import sys
import tempfile
import threading
import time
import uvicorn
import webbrowser
import tkinter as tk
from tkinter import messagebox

# 1. Configurar directorios persistentes y temporales antes de importar nada
HOME_DIR = os.path.expanduser('~')

# DATA_DIR: guarda backups y reportes en la carpeta personal de usuario para persistencia
os.environ['DATA_DIR'] = os.environ.get('DATA_DIR', os.path.join(HOME_DIR, 'DermaEssenzaCRM'))

# TMP_DIR: directorio de descargas temporales y simulación seguro para cualquier OS
os.environ['TMP_DIR'] = os.environ.get('TMP_DIR', os.path.join(tempfile.gettempdir(), 'derma_essenza_tmp'))

# MAESTRO_PATH: por defecto apunta a Descargas
if 'MAESTRO_PATH' not in os.environ:
    os.environ['MAESTRO_PATH'] = os.path.join(HOME_DIR, 'Downloads', 'BD DATA DERMA ESSENZA.xlsx')

# Asegurar directorios creados
os.makedirs(os.environ['DATA_DIR'], exist_ok=True)
os.makedirs(os.environ['TMP_DIR'], exist_ok=True)


# 2. Iniciar el servidor FastAPI en segundo plano
def iniciar_servidor():
    # Importación tardía para usar las variables de entorno correctas
    from app import app
    uvicorn.run(app, host='127.0.0.1', port=8011, log_level='info')


def abrir_navegador():
    webbrowser.open('http://127.0.0.1:8011')


def salir():
    if messagebox.askokcancel("Cerrar", "¿Deseas cerrar el servidor de Derma Essenza?"):
        root.destroy()
        sys.exit(0)


if __name__ == '__main__':
    # Hilo para el servidor FastAPI
    server_thread = threading.Thread(target=iniciar_servidor, daemon=True)
    server_thread.start()

    # Esperar a que el servidor se levante
    time.sleep(1.2)

    # Abrir el navegador automáticamente al inicio
    abrir_navegador()

    # Interfaz Tkinter simple
    root = tk.Tk()
    root.title("Derma Essenza CRM")
    root.geometry("400x200")
    root.resizable(False, False)

    # Centrar la ventana en la pantalla
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width / 2) - (400 / 2)
    y = (screen_height / 2) - (200 / 2)
    root.geometry(f"400x200+{int(x)}+{int(y)}")

    # Elementos de la interfaz
    label_titulo = tk.Label(root, text="Derma Essenza - Servidor Local", font=("Arial", 14, "bold"))
    label_titulo.pack(pady=15)

    label_estado = tk.Label(root, text="El sistema está corriendo en http://127.0.0.1:8011\ny sincronizando con Google Drive.", font=("Arial", 10))
    label_estado.pack(pady=10)

    btn_abrir = tk.Button(root, text="Abrir en Navegador", command=abrir_navegador, width=18, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
    btn_abrir.pack(side=tk.LEFT, padx=25, pady=15)

    btn_salir = tk.Button(root, text="Detener y Salir", command=salir, width=18, bg="#f44336", fg="white", font=("Arial", 10, "bold"))
    btn_salir.pack(side=tk.RIGHT, padx=25, pady=15)

    # Manejar el botón de cerrar (X) de la ventana
    root.protocol("WM_DELETE_WINDOW", salir)

    # Bucle principal de Tkinter
    root.mainloop()
