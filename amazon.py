import tkinter as tk
from tkinter import messagebox, Canvas, Scrollbar
from PIL import Image, ImageTk
import requests
from io import BytesIO
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import xlwt
from xlwt import Workbook
import pyodbc

ruta_carpeta_actual = os.path.dirname(os.path.abspath(__file__))
ruta_chromedriver = os.path.join(ruta_carpeta_actual, "chromedriver.exe")

# XPath inicial y número máximo a probar
xpath_base = '//*[@id="search"]/div[1]/div[1]/div/span[1]/div[1]/div[{0}]/div/div/span/div/div'

max_xpaths = 10

productos = []  # Definir productos a nivel global

def buscar_productos():
    global productos
    nombre_articulo = entrada_articulo.get()

    driver = webdriver.Chrome(executable_path=ruta_chromedriver)
    driver.get("https://www.amazon.com")
    time.sleep(12)  # Asegúrate de que la página haya cargado

    search_bar = driver.find_element(By.ID, "twotabsearchtextbox")
    search_bar.send_keys(nombre_articulo)
    search_bar.send_keys(Keys.RETURN)
    time.sleep(5)

    productos = []
    intentos = 0
    for i in range(70, 44, -1):  # Modificado el límite inferior a 44
        if len(productos) >= max_xpaths:
            break
        xpath = xpath_base.format(i)
        try:
            producto = {}
            imagen_elemento = driver.find_element(By.XPATH, xpath + "/div[1]/span/a/div/img")
            nombre_elemento = driver.find_element(By.XPATH, xpath + "/div[2]/div[2]/h2/a/span")
            material_elemento = driver.find_element(By.XPATH, xpath + "/div[2]/div[2]/div/div")
            precio_elemento = driver.find_element(By.XPATH, xpath + "/div[2]/div[4]/div/div[1]/a/span/span[2]")
            producto["imagen"] = imagen_elemento.get_attribute("src")  # Obtenemos la URL de la imagen
            producto["nombre"] = nombre_elemento.text
            producto["material"] = material_elemento.text
            producto["precio"] = precio_elemento.text
            productos.append(producto)
        except Exception as e:
            print(f"No se pudo extraer el producto con XPath {xpath}: {e}")
        finally:
            intentos += 1

    for widget in frame.winfo_children():
        widget.destroy()

    for i, producto in enumerate(productos):
        # Cargar la imagen desde la URL
        response = requests.get(producto["imagen"])
        image = Image.open(BytesIO(response.content))
        image = image.resize((100, 100), Image.ANTIALIAS)
        imagen = ImageTk.PhotoImage(image)

        # Mostrar la imagen en un Label
        label_imagen = tk.Label(frame, image=imagen)
        label_imagen.image = imagen  # Mantener una referencia para evitar que se borre
        label_imagen.grid(row=i, column=0, padx=5, pady=5)

        # Formatear el precio para mostrarlo correctamente
        precio_original = producto["precio"].replace(',', '').replace('$', '')
        precio_numerico = ''.join(filter(str.isdigit, precio_original))

        if len(precio_numerico) > 2:
            precio_formateado = f"{precio_numerico[:-2]}.{precio_numerico[-2:]}"
        else:
            precio_formateado = f"0.{precio_numerico.zfill(2)}"
        
        # Mostrar la información del producto en un Label
        label_info = tk.Label(frame, text=f"Producto {i+1}:\n{producto['nombre']}\nMaterial: {producto['material']}\nPrecio: ${precio_formateado}", wraplength=200, justify="left")
        label_info.grid(row=i, column=1, padx=5, pady=5, sticky='w')

    driver.quit()

def guardar_en_bd():
    global productos
    try:
        # Conexión utilizando autenticación de Windows
        conexion = pyodbc.connect('DRIVER={SQL Server};SERVER=localhost\\SQLEXPRESS;DATABASE=ProductosDB;Trusted_Connection=yes;')
        cursor = conexion.cursor()

        # Borrar todos los datos existentes en la tabla Productos
        cursor.execute("DELETE FROM Productos")

        for producto in productos:
            # Formatear el precio para guardarlo correctamente
            precio_original = producto["precio"].replace(',', '').replace('$', '')
            precio_numerico = ''.join(filter(str.isdigit, precio_original))
            if len(precio_numerico) > 2:
                precio_formateado = f"{precio_numerico[:-2]}.{precio_numerico[-2:]}"
            else:
                precio_formateado = f"0.{precio_numerico.zfill(2)}"
            cursor.execute("INSERT INTO Productos (Nombre, Precio, ImagenURL) VALUES (?, ?, ?)", producto["nombre"], precio_formateado, producto["imagen"])

        conexion.commit()

        # Verificar que los datos se han insertado correctamente
        cursor.execute("SELECT COUNT(*) FROM Productos")
        count = cursor.fetchone()[0]
        conexion.close()

        messagebox.showinfo("Información", f"Productos guardados en la base de datos. Total de registros: {count}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar en la base de datos: {e}")

def generar_reporte():
    global productos
    try:
        wb = Workbook()
        sheet = wb.add_sheet('Productos')

        for i, producto in enumerate(productos):
            # Formatear el precio para el reporte
            precio_original = producto["precio"].replace(',', '').replace('$', '')
            precio_numerico = ''.join(filter(str.isdigit, precio_original))
            if len(precio_numerico) > 2:
                precio_formateado = f"{precio_numerico[:-2]}.{precio_numerico[-2:]}"
            else:
                precio_formateado = f"0.{precio_numerico.zfill(2)}"
            sheet.write(i, 0, producto["nombre"])
            sheet.write(i, 1, precio_formateado)
            sheet.write(i, 2, producto["imagen"])

        ruta_reporte = os.path.join(ruta_carpeta_actual, 'productos.xls')
        wb.save(ruta_reporte)
        messagebox.showinfo("Información", f"Reporte generado correctamente en {ruta_reporte}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo generar el reporte: {e}")

def salir_del_programa():
    raiz.destroy()

def validar_credenciales(usuario, contrasena):
    ruta_carpeta = os.path.dirname(os.path.abspath(__file__))  # Obtener la carpeta del script actual
    ruta_archivo = os.path.join(ruta_carpeta, "usuarios.txt")
    try:
        with open(ruta_archivo, "r") as file:
            for line in file:
                usuario_archivo, contrasena_archivo = line.strip().split(":")
                if usuario == usuario_archivo and contrasena == contrasena_archivo:
                    return True
    except FileNotFoundError:
        messagebox.showerror("Error", "El archivo 'usuarios.txt' no se ha encontrado.")
    except Exception as e:
        messagebox.showerror("Error", f"Error al leer el archivo 'usuarios.txt': {e}")
    return False

# Función para mostrar la imagen de Amazon
def mostrar_imagen_amazon():
    ruta_imagen_amazon = os.path.join(os.path.dirname(os.path.abspath(__file__)), "amazon.png")
    imagen_amazon = Image.open(ruta_imagen_amazon)
    imagen_amazon = imagen_amazon.resize((300, 200), Image.ANTIALIAS)  # Redimensionar la imagen al doble de tamaño
    foto_amazon = ImageTk.PhotoImage(imagen_amazon)
    etiqueta_imagen_amazon = tk.Label(marco_login, image=foto_amazon)
    etiqueta_imagen_amazon.image = foto_amazon
    etiqueta_imagen_amazon.pack()

def iniciar_sesion():
    usuario = entrada_usuario.get()
    contrasena = entrada_contrasena.get()
    
    if validar_credenciales(usuario, contrasena):
        messagebox.showinfo("Inicio de sesión", "¡Inicio de sesión exitoso!")
        marco_login.pack_forget()  # Ocultar pantalla de login
        marco_principal.pack(fill='both', expand=True)  # Mostrar pantalla principal
    else:
        messagebox.showerror("Inicio de sesión", "Usuario o contraseña incorrectos")

# Crear la ventana principal
raiz = tk.Tk()
raiz.title("Buscador Amazon")
raiz.geometry("400x500")  # Tamaño inicial ajustado

# Frame para el login
marco_login = tk.Frame(raiz)
marco_login.pack(fill='both', expand=True)

# Mostrar la imagen de Amazon en la parte superior
mostrar_imagen_amazon()

tk.Label(marco_login, text="Usuario").pack(pady=5)
entrada_usuario = tk.Entry(marco_login)
entrada_usuario.pack(pady=5)
tk.Label(marco_login, text="Contraseña").pack(pady=5)
entrada_contrasena = tk.Entry(marco_login, show="*")
entrada_contrasena.pack(pady=5)
tk.Button(marco_login, text="Iniciar sesión", command=iniciar_sesion).pack(pady=10)

# Frame principal
marco_principal = tk.Frame(raiz)

tk.Label(marco_principal, text="Buscar artículo en Amazon").pack(pady=5)
entrada_articulo = tk.Entry(marco_principal)
entrada_articulo.pack(pady=5)
tk.Button(marco_principal, text="Buscar", command=buscar_productos).pack(pady=5)

frame_canvas = tk.Frame(marco_principal)
frame_canvas.pack(fill='both', expand=True, pady=10)

# Crear un Canvas para mostrar productos con barra de desplazamiento
canvas = Canvas(frame_canvas)
scrollbar = Scrollbar(frame_canvas, orient="vertical", command=canvas.yview)
scrollable_frame = tk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(
        scrollregion=canvas.bbox("all")
    )
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

frame = scrollable_frame

tk.Button(marco_principal, text="Guardar en BD", command=guardar_en_bd).pack(pady=5)
tk.Button(marco_principal, text="Generar Reporte", command=generar_reporte).pack(pady=5)
tk.Button(marco_principal, text="Salir", command=salir_del_programa).pack(pady=5)

# Ejecutar la aplicación
raiz.mainloop()
