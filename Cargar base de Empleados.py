import openpyxl
from openpyxl.styles import Font
import os


# Main menu
def mostrar_menu():
    print("\nMenú:")
    print("1. Agregar Empleados")
    print("2. Modificar Empleados")
    print("3. Borrar Empleados")
    print("4. Salir")


# Guardado de cambios en el excel
def guardar_archivo():
    workbook.save("Empleados.xlsx")
    print("Cambios guardados en 'Empleados.xlsx'.")


# Solicitar al usuario que ingrese los datos
def agregar_empleado():
    numero_legajo = input("Ingrese el número de legajo: ")
    nombre = input("Ingrese el nombre: ")
    edad = int(input("Ingrese la edad: "))
    area = input("Ingrese el área: ")
    horas_diarias = float(input("Ingrese las horas diarias: "))
    pago_por_hora = float(input("Ingrese el pago por hora: "))
    dias_trabajados = int(input("Ingrese los días trabajados: "))
    sueldo_mensual = horas_diarias * pago_por_hora * dias_trabajados

    data = [
        numero_legajo,
        nombre,
        edad,
        area,
        horas_diarias,
        pago_por_hora,
        dias_trabajados,
        sueldo_mensual,
    ]
    sheet.append(data)
    guardar_archivo()
    print("Empleado se agrego con éxito!!")


mostrar_menu()


# Modificar empleado


def modificar_empleado():
    pass


# Borrar empleado


def borrar_empleado():
    pass


# Verificar si el archivo Excel ya existe
if os.path.exists("datos.xlsx"):
    workbook = openpyxl.load_workbook("Empleados.xlsx")
    sheet = workbook.active
else:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Empleados_Mañana"
    headers = [
        "Numero de Legajo",
        "Nombre",
        "Edad",
        "Area",
        "Horas diarias",
        "Pago por hora",
        "Dias trabajados",
        "Sueldo mensual",
    ]
    sheet.append(headers)
    guardar_archivo()

mostrar_menu()


# Aplicar formato a los encabezados en negrita
for cell in sheet[1]:
    cell.font = Font(bold=True)


while True:
    opcion = input("\nIngrese la opción deseada: ")

    if opcion == "1":
        agregar_empleado()
    elif opcion == "2":
        modificar_empleado()
    elif opcion == "3":
        borrar_empleado()
    elif opcion == "4":
        print("Muchas gracias por usar el sistema!")
        break
    else:
        print("Opción no válida. Por favor, elija una opción del menú.")
