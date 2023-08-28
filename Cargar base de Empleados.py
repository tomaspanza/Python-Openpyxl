import openpyxl
from openpyxl.styles import Font
import os

# Verificar si el archivo Excel ya existe
if os.path.exists("Empleados.xlsx"):
    workbook = openpyxl.load_workbook("Empleados.xlsx")
    sheet = workbook.active
else:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Empleados_Mañana"
    headers = [
        "Nombre",
        "Edad",
        "Area",
        "Horas diarias",
        "Pago por hora",
        "Dias trabajados",
        "Sueldo mensual",
    ]
    sheet.append(headers)

# Solicitar al usuario que ingrese los datos
nombre = input("Ingrese el nombre: ")
edad = int(input("Ingrese la edad: "))
area = input("Ingrese el área: ")
horas_diarias = float(input("Ingrese las horas diarias: "))
pago_por_hora = float(input("Ingrese el pago por hora: "))
dias_trabajados = int(input("Ingrese los días trabajados: "))
sueldo_mensual = horas_diarias * pago_por_hora * dias_trabajados

data = [
    nombre,
    edad,
    area,
    horas_diarias,
    pago_por_hora,
    dias_trabajados,
    sueldo_mensual,
]
sheet.append(data)

# Aplicar formato a los encabezados en negrita
for cell in sheet[1]:
    cell.font = Font(bold=True)

# Guardar el libro de trabajo en un archivo
workbook.save("Empleados.xlsx")

print("Datos agregados al archivo 'Empleados.xlsx' con éxito.")
