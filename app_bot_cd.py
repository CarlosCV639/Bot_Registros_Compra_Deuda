import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import win32com.client as win32
import pandas as pd
import pyautogui
import pyperclip
import logging
import time
import cv2
import os
import pygetwindow as gw

# Ruta fija donde guardar logs
log_dir = r"\\ruta\para_generar\logs"
os.makedirs(log_dir, exist_ok=True)

# Nombre de archivo con ruta completa
fecha_actual = datetime.now().strftime("%Y-%m-%d")
log_filename = os.path.join(log_dir, f"log_proceso_{fecha_actual}.txt")

# Configurar logging
if not logging.getLogger().hasHandlers():
    logging.basicConfig(
        filename=log_filename,
        filemode='w',
        format='%(asctime)s - %(levelname)s - %(message)s',
        level=logging.INFO
    )

# Define una función para obtener rutas absolutas
def resource_path(relative_path):
    """Obtén la ruta absoluta de un recurso."""
    base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)

def wait_for_image(image_path, timeout=10, confidence=0.8):
    """Espera hasta que una imagen sea visible en pantalla."""
    start_time = time.time()
    while True:
        location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
        if location:
            return location  # Imagen encontrada
        if time.time() - start_time > timeout:
            raise TimeoutError(f"No se encontró la imagen: {image_path}")
        time.sleep(1)  # Pequeña pausa entre búsquedas

def ensure_window_active_2(window_title):
    """Asegúrate de que la ventana esté activa."""
    windows = gw.getWindowsWithTitle(window_title)
    if windows:
        window = windows[0]
        window.restore()  # Restaura la ventana si está minimizada
        window.minimize()  # Minimiza y luego maximiza para forzar la activación
        window.maximize()
        time.sleep(1)  # Breve pausa
    else:
        raise Exception(f"No se encontró una ventana con el título: {window_title}")

def upload_file():
    """Permite al usuario seleccionar el archivo Excel."""
    global file_path
    file_path = filedialog.askopenfilename(
        title="Seleccionar archivo CONSOLIDADO CD",
        filetypes=[("Archivos Excel", "*.xlsx")]
    )
    if file_path:
        lbl_file.config(text=f"Archivo seleccionado: {os.path.basename(file_path)}")

def process_file():
    """Ejecuta el proceso principal del bot."""
    if not file_path:
        messagebox.showwarning("Advertencia", "Por favor, selecciona un archivo antes de continuar.")
        return

    logging.info(f"Iniciando procesamiento del archivo: {file_path}")

    try:
        # Leer archivo Excel
        df = pd.read_excel(file_path)
        df['IMP.SOL'] = df['IMP.SOL'].astype(float).round(2)

        ensure_window_active_2("S.A.T (Interbank)")

        for index, row in df.iterrows():
            try:
                contrato = row['CONTRATO']
                contrato = str(contrato)

                logging.info(f"Procesando contrato: {contrato}")

                monto = row['IMP.SOL']
                monto = str(monto)

                plazo = row['PL']
                plazo = str(plazo)

                intereses = row['TEM']
                intereses = str(intereses)

                #SAT --> Consulta Operaciones
                pyautogui.moveTo(51, 134, duration = 0.5)
                pyautogui.click()
                pyautogui.moveTo(125, 256, duration = 0.25)
                pyautogui.click()
                pyautogui.moveTo(330, 293, duration = 0.25)
                pyautogui.click()
                time.sleep(2.5)

                #Click en C.Cuotas
                pyautogui.moveTo(1452, 256, duration = 0.5)
                pyautogui.click()

                #Escribir Número de Contrato
                pyautogui.moveTo(647, 347, duration = 0.25)
                pyautogui.click()
                pyautogui.write(contrato)
                pyautogui.hotkey('enter')
                time.sleep(3.5)

                #Click en Opciones --> Usar imagen
                try:
                    location1 = wait_for_image(resource_path('Imagenes_SAT/opciones.png'), timeout=10)
                    pyautogui.click(location1, clicks = 3)
                    df.at[index, 'ESTADO'] = 'OK'
                except Exception as e:
                    df.at[index, 'ESTADO'] = 'REVISAR'
                    print(f"Error al encontrar la imagen 'opciones.png': {e}")
                    logging.error(f"Error al procesar contrato {contrato}: {e}")
                    continue

                #Click en Alta C.Cuotas
                pyautogui.moveTo(1345, 270, duration = 0.25)
                pyautogui.click()
                time.sleep(4.5)

                #Seleccionar Moneda
                pyautogui.moveTo(1250, 251, duration = 0.5)
                pyautogui.click()
                pyautogui.moveTo(1206, 264, duration = 0.25)
                pyautogui.click()

                #Seleccionar Tipo de línea
                pyautogui.moveTo(1406, 552, duration = 0.25)
                pyautogui.click()
                time.sleep(2)
                try:
                    location2 = wait_for_image(resource_path('Imagenes_SAT/tipo_linea.png'), timeout=10)
                    pyautogui.click(location2)
                    df.at[index, 'ESTADO'] = 'OK'
                except Exception as e:
                    time.sleep(1)
                    location7 = wait_for_image(resource_path('Imagenes_SAT/cerrar.png'), timeout=10)
                    pyautogui.click(location7)
                    df.at[index, 'ESTADO'] = 'REVISAR'
                    print(f"Error al encontrar la imagen '0006.png': {e}")
                    logging.error(f"Error al procesar contrato {contrato}: {e}")
                    continue
                time.sleep(0.25)

                #Click en Interés
                pyautogui.moveTo(599, 497, duration = 0.5)
                pyautogui.click()

                #Seleccionar Tipo de Cuota
                pyautogui.moveTo(252, 552, duration = 0.25)
                pyautogui.click()
                pyautogui.moveTo(165, 593, duration = 0.25)
                pyautogui.click()

                #Escribir Importe a financiar
                pyautogui.moveTo(181, 497, duration = 0.25)
                pyautogui.click()
                pyautogui.write(monto)

                #Escribir N° Cuotas
                pyautogui.moveTo(653, 552, duration = 0.25)
                pyautogui.click()
                pyautogui.write(plazo)
                
                #Escribir el %Interes
                pyautogui.moveTo(1191, 497, duration = 0.25)
                pyautogui.click()
                pyautogui.write(intereses)
                time.sleep(0.25)

                #Click en Simular --> Usar imagen
                try:
                    location3 = wait_for_image(resource_path('Imagenes_SAT/simular.png'), timeout=10)
                    pyautogui.click(location3)
                    df.at[index, 'ESTADO'] = 'OK'
                except Exception as e:
                    df.at[index, 'ESTADO'] = 'REVISAR'
                    print(f"Error al encontrar la imagen 'simular.png': {e}")
                    logging.error(f"Error al procesar contrato {contrato}: {e}")
                    continue
                time.sleep(1)
                pyautogui.hotkey('Enter')
                time.sleep(5)

                #Click en Confirmar
                try:
                    location4 = wait_for_image(resource_path('Imagenes_SAT/confirmar.png'), timeout=10)
                    pyautogui.click(location4)
                    df.at[index, 'ESTADO'] = 'OK'
                except Exception as e:
                    df.at[index, 'ESTADO'] = 'REVISAR'
                    print(f"Error al encontrar la imagen 'confirmar.png': {e}")
                    logging.error(f"Error al procesar contrato {contrato}: {e}")
                    continue
                time.sleep(1)
                pyautogui.hotkey('enter')
                time.sleep(5)
                try:
                    location5 = wait_for_image(resource_path('Imagenes_SAT/aceptar.png'), timeout=10)
                    pyautogui.click(location5)
                    df.at[index, 'ESTADO'] = 'OK'
                except Exception as e:
                    time.sleep(5)
                    pyautogui.hotkey('enter')
                    df.at[index, 'ESTADO'] = 'REVISAR'
                    print(f"Error al encontrar la imagen 'aceptar.png': {e}")
                    logging.error(f"Error al procesar contrato {contrato}: {e}")
                    continue
                time.sleep(2.5)

                #Click en Opciones
                try:
                    location6 = wait_for_image(resource_path('Imagenes_SAT/opciones.png'), timeout=10)
                    pyautogui.click(location6, clicks = 3)
                    df.at[index, 'ESTADO'] = 'OK'
                except Exception as e:
                    df.at[index, 'ESTADO'] = 'REVISAR'
                    print(f"Error al encontrar la imagen 'opciones.png': {e}")
                    logging.error(f"Error al procesar contrato {contrato}: {e}")
                    continue
                
                #Click en Consulta Contratos
                pyautogui.moveTo(1353, 316, duration = 0.25)
                pyautogui.click()
                time.sleep(2)

                #Copia y pega NOMBRE en df
                pyautogui.moveTo(42, 286, duration = 1)
                pyautogui.click(clicks = 3)
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(0.5)
                nombre_copiado = pyperclip.paste()
                df.at[index, 'NOMBRE'] = nombre_copiado
                time.sleep(1.5)
                logging.info(f"Contrato {contrato} procesado correctamente.")
            
            except Exception as global_exception:
                # Manejo general para errores inesperados, excepto pyautogui.FailSafeException
                df.at[index, 'ESTADO'] = 'REVISAR'
                print(f"Error inesperado al procesar el contrato {contrato}: {global_exception}")
                logging.error(f"Error inesperado con contrato {contrato}: {global_exception}")

        # Ruta personalizada donde quieres guardar automáticamente el archivo resultado
        ruta_resultado_definida = r"\\rutas\resultados_bot"
        os.makedirs(ruta_resultado_definida, exist_ok=True)  # Crea la carpeta si no existe
        
        # Guardar resultados
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            title="Guardar archivo de resultados"
        )

        total = len(df)
        ok = df['ESTADO'].value_counts().get('OK', 0)
        logging.info(f"Success: El ARPI procesó {ok}/{total} filas el día {datetime.now().strftime('%d/%m/%Y')}")

        # Procesamiento final de columnas
        df['FECHA'] = pd.to_datetime(df['FECHA']).dt.date
        df['CLIENTE'] = df['CLIENTE'].astype(str).str.zfill(10)
        df['CONTRATO'] = df['CONTRATO'].astype(str).str.zfill(12)
        
        # Guardado automático en la ruta personalizada
        fecha_actual_nombre = datetime.now().strftime("%d.%m")  # Formato: día.mes
        nombre_archivo = f"CONSOLIDADO {fecha_actual_nombre}.xlsx"
        ruta_guardado_automatico = os.path.join(ruta_resultado_definida, nombre_archivo)
        df.to_excel(ruta_guardado_automatico, index=False)

        # Luego dejar que el usuario elija otra ubicación si desea
        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Éxito", f"Procesamiento completado. Archivo guardado como:\n{save_path}")
        else:
            messagebox.showinfo("Cancelado", "Guardado cancelado por el usuario. El archivo 'resultado.xlsx' se guardó en la ruta compartida.")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al procesar el archivo: {e}")
        lbl_status.config(text="Estado: Error durante el procesamiento.")


def download_file():
    """Permite al usuario guardar el archivo resultado.xlsx en un lugar específico."""
    try:
        if not os.path.exists("resultado.xlsx"):
            messagebox.showwarning("Advertencia", "No se encontró el archivo de resultados para descargar.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Guardar archivo resultado",
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")]
        )
        if save_path:
            os.rename("resultado.xlsx", save_path)
            messagebox.showinfo("Éxito", f"Archivo guardado como {save_path}.")
        else:
            messagebox.showinfo("Cancelado", "Descarga cancelada.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al guardar el archivo: {e}")


def enviar_correo():
    """Envía el correo con el archivo generado y logs."""
    try:

        hoy = datetime.now()
        fecha_ddmm = hoy.strftime('%d.%m')
        fecha_yyyy_mm_dd = hoy.strftime('%Y-%m-%d')

        ruta_excel = fr"\\rutas\resultados_bot\CONSOLIDADO {fecha_ddmm}.xlsx"
        ruta_txt_1 = r"\\rutas\compra-deuda-auto-logs\data-cruzada-log.txt"
        ruta_txt_2 = fr"\\rutas\compra-deuda-auto-logs\log_proceso_{fecha_yyyy_mm_dd}.txt"

        if not os.path.exists(ruta_excel):
            messagebox.showerror("Error", f"No se encontró el archivo Excel:\n{ruta_excel}")
            return

        # Leer logs
        with open(ruta_txt_1, "r", encoding="latin-1") as file1:
            lineas_txt1 = file1.readlines()
            texto1 = lineas_txt1[-1].strip() if lineas_txt1 else "[Archivo vacío]"

        with open(ruta_txt_2, "r", encoding="latin-1") as file2:
            lineas = file2.readlines()
            ultima_linea = lineas[-1] if lineas else ""
            indice = ultima_linea.find("Success")
            texto2 = ultima_linea[indice:].strip() if indice != -1 else "[No se encontró 'Success']"

        cuerpo = f"""Buen día estimados,

Envío el detalle de la base de CD trabajada el {fecha_ddmm} con su respectiva revisión:

- {texto1}
- {texto2}

Saludos,
Carlos Cuba
"""

        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.To = "elcorreo@github.com"
        mail.CC = "suscorreos@github.com"
        mail.Subject = f"Consolidado CD - {fecha_ddmm}"
        mail.Body = cuerpo
        mail.Attachments.Add(ruta_excel)

        mail.Display()  # Cambia a mail.Send() si quieres que se envíe automáticamente

        messagebox.showinfo("Correo listo", "Correo generado exitosamente. Revísalo antes de enviarlo.")
        lbl_status.config(text="Estado: Correo generado")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al generar el correo:\n{e}")
        lbl_status.config(text="Estado: Error al generar el correo")


# Crear ventana principal con ttkbootstrap
theme = "superhero"  # Puedes cambiarlo por otros temas como 'darkly', 'cosmo', etc.
root = ttk.Window(themename=theme)
root.title("ARPI para Registros CD")
root.geometry("620x520")

frame_main = ttk.Frame(root, padding=10)
frame_main.pack(fill=BOTH, expand=True)

lbl_title = ttk.Label(frame_main, text="ARPI para Procesar CD", font=("Arial", 16, "bold"))
lbl_title.pack(pady=10)

btn_upload = ttk.Button(frame_main, text="Subir Archivo CD", bootstyle=PRIMARY, command=upload_file)
btn_upload.pack(pady=5)

lbl_file = ttk.Label(frame_main, text="Archivo seleccionado: Ninguno", bootstyle="info")
lbl_file.pack(pady=5)

btn_process = ttk.Button(frame_main, text="Procesar Archivo", bootstyle=SUCCESS, command=process_file)
btn_process.pack(pady=5)

btn_download = ttk.Button(frame_main, text="Descargar Resultados", bootstyle=WARNING, command=download_file)
btn_download.pack(pady=5)

btn_email = ttk.Button(frame_main, text="Enviar Correo", bootstyle=INFO, command=enviar_correo, width=13)
btn_email.pack(pady=5)

lbl_status = ttk.Label(frame_main, text="Estado: Esperando acción del usuario", bootstyle="success")
lbl_status.pack(pady=10)

frame_notes = ttk.LabelFrame(frame_main, text="Consideraciones importantes", bootstyle="danger")
frame_notes.pack(fill=BOTH, expand=True, pady=10)

notes = [
    "1. Asegúrate de que tu archivo tenga la columna 'CONTRATO'.",
    "2. Mantén abierta la ventana 'S.A.T (Interbank)' con zoom de 90%.",
    "3. Para detener el bot, mueve el cursor a la esquina superior derecha.",
]
for note in notes:
    ttk.Label(frame_notes, text=note, bootstyle="white").pack(anchor="w", padx=5, pady=2)

root.mainloop()
