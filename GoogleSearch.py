

import pandas as pd
import logging
from datetime import datetime
import time
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from utils import initialize_driver, setup_logging, enviar_texto_a_input, esperar_y_clicar, elemento_visible

# Declarar variables globales
fichero_entry = None
root = None

def seleccionar_fichero():
    global fichero_entry
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
    )
    if file_path:
        fichero_entry.delete(0, tk.END)
        fichero_entry.insert(0, file_path)

def iniciar_programa():
    global fichero_entry
    file_path = fichero_entry.get()
    if file_path.endswith(('.xlsx', '.xls')):
        # Logging
        directorio_base = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        current_time = datetime.now()
        setup_logging(current_time)

        # Leemos Excel
        logging.info("Información", f"Fichero seleccionado: {file_path}")
        df = pd.read_excel(file_path)
        logging.info("Excel leido en DataFrame")

        # Agregar la columna 'Lucky Search'
        df['Lucky Search'] = ''
        logging.info("Nueva columna en DataFrame")

        # Mostrar el DataFrame resultante
        print(df)

        # Inicializamos driver Chrome
        driver = initialize_driver()
        url = 'https://www.google.com/'
        logging.info("Driver inicializado")

        # Entramos en la web de google por primera vez
        driver.get(url)
        driver.maximize_window()

        # Verificar si se visualiza la seleccion de busqueda predeterminada
        es_visible = elemento_visible(driver, By.ID, "actionButton")
        # Selecciona el botón "Google" como predeterminado
        if es_visible:
            esperar_y_clicar(driver, 'name', "1", 1, 'Clicando en la opción Google')
            esperar_y_clicar(driver, 'id', "actionButton", 1, 'Clicando en la opción Establecer Predeterminado')
        else:
            logging.info("La seleccion de busqueda predeterminada no es visible.")

        # Verificar si se visualiza el inicio de sesión
        es_visible = elemento_visible(driver, By.ID, "W0wltc")
        # Selecciona el botón "Rechazar Todo"
        if es_visible:
            esperar_y_clicar(driver, 'id', "W0wltc", 1, 'Clicando en la opción Rechazar Todo')
        else:
            logging.info("El inicio de sesión de Google no es visible.")

        for index, row in df.iterrows():
            # Entramos en la web de google
            driver.get(url)

            # Recogemos valor a buscar
            search_value = row['Search']

            # Copiamos texto en casilla y damos a buscar
            enviar_texto_a_input(driver, 'name', 'q', 1, search_value, f'Valor introducido: {search_value}.')

            # Pinchamos en Voy a tener Suerte
            esperar_y_clicar(driver, 'name', 'btnI', 1, 'Clicando en el botón Voy a tener Suerte')

            # Guardamos el resultado en el dataFrame
            df.at[index, 'Lucky Search'] = driver.current_url

        # Guardar el DataFrame actualizado 
        timestamp = current_time.strftime('%Y%m%d__%H%M%S')
        output_file_final = os.path.join(directorio_base, f'Search_{timestamp}.xlsx')
        df.to_excel(output_file_final, index=False)

        logging.info("Proceso finalizado")
        messagebox.showinfo("FIN","Proceso completado")

    else:
        messagebox.showerror("Error", "Por favor selecciona un fichero Excel válido.")

def main():
    try:
        global fichero_entry, root

        # Crear la ventana de Tkinter
        root = tk.Tk()
        root.title("Selector de Fichero Excel")
        root.geometry("500x200")

        # Crear un cuadro de texto para mostrar el fichero seleccionado
        fichero_entry = tk.Entry(root, width=50)
        fichero_entry.pack(pady=20)

        # Botón para seleccionar el fichero
        boton_seleccionar = tk.Button(root, text="Seleccionar Fichero", command=seleccionar_fichero)
        boton_seleccionar.pack(pady=10)

        # Botón para iniciar el programa
        boton_iniciar = tk.Button(root, text="Iniciar Programa", command=iniciar_programa)
        boton_iniciar.pack(pady=10)

        # Ejecutar la ventana de Tkinter
        root.mainloop()

    except Exception as e:
        logging.exception("Error inesperado en el script principal.")
        error_message = f"Ha ocurrido un error: {str(e)}"
        messagebox.showerror("Error", error_message)
        exit(1)

if __name__ == "__main__":
    main()
