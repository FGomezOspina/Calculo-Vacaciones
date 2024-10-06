import tkinter as tk
from tkinter import messagebox, ttk
from openpyxl import Workbook
from datetime import datetime, timedelta
import holidays
import sqlite3
import os
import sys


# Obtener la ruta del directorio del archivo ejecutable
if getattr(sys, 'frozen', False):
    # Si estamos en un ejecutable
    directorio_base = sys._MEIPASS
else:
    # Si estamos en el entorno de desarrollo
    directorio_base = os.path.dirname(os.path.abspath(__file__))


# Construir la ruta de la base de datos
db_path = os.path.join(directorio_base, "vacaciones_empleados.db")

# Conexión a la base de datos (creada si no existe)
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

# Crear la tabla de empleados en la base de datos si no existe
cursor.execute('''CREATE TABLE IF NOT EXISTS empleados (
                    id INTEGER PRIMARY KEY,
                    nombre TEXT,
                    fecha_ingreso TEXT,
                    fecha_actual TEXT,
                    fecha_inicio_vacaciones TEXT,
                    fecha_fin_vacaciones TEXT,
                    dias_disponibles INTEGER
                )''')
conn.commit()

# Función para calcular días hábiles
def calcular_dias_habiles(fecha_inicio, dias, pais='CO'):
    co_holidays = holidays.CountryHoliday(pais, years=fecha_inicio.year)
    current_day = fecha_inicio
    habiles = 0

    while habiles < dias:
        current_day += timedelta(days=1)
        if current_day.weekday() < 5 and current_day not in co_holidays:
            habiles += 1
    return current_day

# Función para mostrar u ocultar el campo de días de vacaciones
def toggle_dias_vacaciones():
    if var_sesiones.get() == 1:
        dias_vacaciones_label.grid(row=6, column=0, padx=10, pady=10)
        dias_vacaciones_entry.grid(row=6, column=1, padx=10, pady=10)
    else:
        dias_vacaciones_label.grid_forget()
        dias_vacaciones_entry.grid_forget()

# Función para autocompletar el nombre y otros campos al ingresar el ID
def autocompletar_empleado(event):
    id_empleado = id_entry.get()
    cursor.execute("SELECT * FROM empleados WHERE id = ?", (id_empleado,))
    empleado = cursor.fetchone()

    if empleado:
        nombre_entry.delete(0, tk.END)
        nombre_entry.insert(0, empleado[1])
        fecha_ingreso_entry.delete(0, tk.END)
        fecha_deseada_entry.delete(0, tk.END)
        dias_vacaciones_entry.delete(0, tk.END)
    else:
        nombre_entry.delete(0, tk.END)
        fecha_ingreso_entry.delete(0, tk.END)
        fecha_deseada_entry.delete(0, tk.END)
        dias_vacaciones_entry.delete(0, tk.END)

# Función para agregar o actualizar empleado en la base de datos
def agregar_empleado():
    id_empleado = id_entry.get()
    nombre = nombre_entry.get()
    fecha_ingreso = fecha_ingreso_entry.get()
    fecha_actual = datetime.now().strftime('%Y-%m-%d')

    if not id_empleado or not nombre or not fecha_ingreso or not fecha_deseada_entry.get():
        messagebox.showerror("Error", "Por favor, completa todos los campos.")
        return

    try:
        fecha_deseada = datetime.strptime(fecha_deseada_entry.get(), '%Y-%m-%d')
    except ValueError:
        messagebox.showerror("Error", "La fecha deseada de inicio de vacaciones debe tener el formato YYYY-MM-DD.")
        return

    cursor.execute("SELECT id FROM empleados WHERE id = ?", (id_empleado,))
    resultado = cursor.fetchone()

    if tipo_trabajador.get() == "Poscosecha":
        if var_sesiones.get() == 1:
            try:
                dias_tomados = int(dias_vacaciones_entry.get())
            except ValueError:
                messagebox.showerror("Error", "Por favor, ingresa un número válido de días de vacaciones.")
                return

            dias_disponibles = 15 - dias_tomados
            fecha_vacaciones_fin = calcular_dias_habiles(fecha_deseada, dias_tomados).strftime('%Y-%m-%d')
        else:
            dias_tomados = 15
            dias_disponibles = 0
            fecha_vacaciones_fin = calcular_dias_habiles(fecha_deseada, 15).strftime('%Y-%m-%d')
    else:
        try:
            dias_tomados = int(dias_vacaciones_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Por favor, ingresa un número válido de días de vacaciones.")
            return

        dias_disponibles = 15 - dias_tomados
        fecha_vacaciones_fin = calcular_dias_habiles(fecha_deseada, dias_tomados).strftime('%Y-%m-%d')

    fecha_inicio_vacaciones = fecha_deseada_entry.get()

    if resultado:
        cursor.execute("""UPDATE empleados SET nombre = ?, fecha_ingreso = ?, fecha_actual = ?, 
                        fecha_inicio_vacaciones = ?, fecha_fin_vacaciones = ?, dias_disponibles = ? WHERE id = ?""",
                       (nombre, fecha_ingreso, fecha_actual, fecha_inicio_vacaciones, fecha_vacaciones_fin, dias_disponibles, id_empleado))
        messagebox.showinfo("Éxito", "Datos del empleado actualizados correctamente.")

        for item in tabla.get_children():
            if tabla.item(item, 'values')[0] == id_empleado:
                tabla.item(item, values=(id_empleado, nombre, fecha_ingreso, fecha_actual, fecha_inicio_vacaciones, fecha_vacaciones_fin, dias_disponibles))
                break
    else:
        cursor.execute("""INSERT INTO empleados (id, nombre, fecha_ingreso, fecha_actual, 
                        fecha_inicio_vacaciones, fecha_fin_vacaciones, dias_disponibles) 
                        VALUES (?, ?, ?, ?, ?, ?, ?)""", 
                       (id_empleado, nombre, fecha_ingreso, fecha_actual, fecha_inicio_vacaciones, fecha_vacaciones_fin, dias_disponibles))
        messagebox.showinfo("Éxito", "Empleado agregado correctamente.")
        tabla.insert("", tk.END, values=(id_empleado, nombre, fecha_ingreso, fecha_actual, fecha_inicio_vacaciones, fecha_vacaciones_fin, dias_disponibles))

    conn.commit()

    id_entry.delete(0, tk.END)
    nombre_entry.delete(0, tk.END)
    fecha_ingreso_entry.delete(0, tk.END)
    fecha_deseada_entry.delete(0, tk.END)
    dias_vacaciones_entry.delete(0, tk.END)

# Función para eliminar empleado de la tabla y de la base de datos
def eliminar_empleado():
    selected_item = tabla.selection()
    if selected_item:
        empleado = tabla.item(selected_item, 'values')
        id_empleado = empleado[0]

        cursor.execute("DELETE FROM empleados WHERE id = ?", (id_empleado,))
        conn.commit()

        tabla.delete(selected_item)
        messagebox.showinfo("Eliminar", f"Empleado {empleado[1]} eliminado correctamente.")
    else:
        messagebox.showwarning("Error", "Seleccione un empleado para eliminar.")

# Función para generar el archivo Excel a partir de la base de datos
def generar_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Vacaciones"

    # Encabezados
    ws.append(["ID", "Nombre", "Fecha de Ingreso", "Fecha Actual", "Fecha Inicio de Vacaciones", "Fecha Fin de Vacaciones", "Días Disponibles"])

    # Extraer todos los empleados de la base de datos
    cursor.execute("SELECT * FROM empleados")
    empleados = cursor.fetchall()

    for empleado in empleados:
        ws.append(empleado)

    # Guardar el archivo Excel
    archivo_excel = "vacaciones_empleados.xlsx"
    wb.save(archivo_excel)
    messagebox.showinfo("Éxito", "Archivo Excel generado con éxito.")

    # Abrir el archivo Excel generado
    os.system("open " + archivo_excel)  # Para macOS


# Crear la interfaz gráfica
root = tk.Tk()
root.title("Cálculo de Vacaciones")

tk.Label(root, text="ID del empleado:").grid(row=0, column=0, padx=10, pady=10)
id_entry = tk.Entry(root)
id_entry.grid(row=0, column=1, padx=10, pady=10)
id_entry.bind("<KeyRelease>", autocompletar_empleado)

tk.Label(root, text="Nombre del empleado:").grid(row=1, column=0, padx=10, pady=10)
nombre_entry = tk.Entry(root)
nombre_entry.grid(row=1, column=1, padx=10, pady=10)

tk.Label(root, text="Fecha de ingreso (YYYY-MM-DD):").grid(row=2, column=0, padx=10, pady=10)
fecha_ingreso_entry = tk.Entry(root)
fecha_ingreso_entry.grid(row=2, column=1, padx=10, pady=10)

tk.Label(root, text="Tipo de trabajador:").grid(row=3, column=0, padx=10, pady=10)
tipo_trabajador = tk.StringVar()
opciones = ["Poscosecha", "Administrativo", "Mantenimiento"]
tipo_trabajador_menu = ttk.Combobox(root, textvariable=tipo_trabajador, values=opciones)
tipo_trabajador_menu.grid(row=3, column=1, padx=10, pady=10)

tk.Label(root, text="Fecha deseada de inicio de vacaciones (YYYY-MM-DD):").grid(row=4, column=0, padx=10, pady=10)
fecha_deseada_entry = tk.Entry(root)
fecha_deseada_entry.grid(row=4, column=1, padx=10, pady=10)


# Campo para seleccionar si se toma menos de 15 días
var_sesiones = tk.IntVar()
check_sesiones = tk.Checkbutton(root, text="¿Desea tomar menos de 15 días?", variable=var_sesiones, command=toggle_dias_vacaciones)
check_sesiones.grid(row=5, columnspan=2, pady=10)

# Etiqueta y campo para días de vacaciones a tomar (inicialmente ocultos)
dias_vacaciones_label = tk.Label(root, text="Días de vacaciones a tomar:")
dias_vacaciones_entry = tk.Entry(root)

# Botón para agregar empleado
tk.Button(root, text="Agregar empleado con vacaciones", command=agregar_empleado).grid(row=7, columnspan=2, pady=10)

# Crear la tabla para mostrar empleados
tabla = ttk.Treeview(root, columns=("ID", "Nombre", "Fecha de Ingreso", "Fecha Actual", "Fecha Inicio de Vacaciones", "Fecha Fin de Vacaciones", "Días Disponibles"), show="headings")
tabla.heading("ID", text="ID")
tabla.heading("Nombre", text="Nombre")
tabla.heading("Fecha de Ingreso", text="Fecha de Ingreso")
tabla.heading("Fecha Actual", text="Fecha Actual")
tabla.heading("Fecha Inicio de Vacaciones", text="Fecha Inicio de Vacaciones")
tabla.heading("Fecha Fin de Vacaciones", text="Fecha Fin de Vacaciones")
tabla.heading("Días Disponibles", text="Días Disponibles")

# Hacemos que la tabla ocupe varias filas y columnas con grid
tabla.grid(row=8, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

# Botón para eliminar empleado
tk.Button(root, text="Eliminar empleado", command=eliminar_empleado).grid(row=9, columnspan=2, pady=10)

# Botón para generar Excel
btn_generar_excel = tk.Button(root, text="Generar Excel", command=generar_excel)
btn_generar_excel.grid(row=10, columnspan=2, pady=10)

# Hacer que las columnas y filas de la tabla se expandan correctamente
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
root.grid_rowconfigure(8, weight=1)

# Cargar empleados existentes desde la base de datos y mostrarlos en la tabla
def cargar_empleados():
    cursor.execute("SELECT * FROM empleados")
    empleados = cursor.fetchall()
    for empleado in empleados:
        tabla.insert("", tk.END, values=empleado)

# Ejecutar la función para cargar los empleados al iniciar la aplicación
cargar_empleados()

# Iniciar el bucle principal de la interfaz gráfica
root.mainloop()

# Cerrar la conexión a la base de datos al salir
conn.close()
