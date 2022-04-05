########################################
#
# Autor: Ernesto Lomar
# Contacto: marioernestolomar@protonmail.ch
#
########################################

from openpyxl import Workbook
import random

libro = Workbook()

hoja = libro.active
hoja["A1"] = "Genero"
hoja["B1"] = "Num Empleos"
hoja["C1"] = "T.A"
hoja["D1"] = "H.T"
hoja["E1"] = "Sueldo"

arr_sexo = ["M","F"]
arr_cant_empleos = ["1","2","3 o más"]
arr_trabajo_actual = ["Si", "No"]
arr_hora_trabajo = ["6am-12pm", "12pm-6am","9am-6pm"]
arr_suedo = ["$0-$500", "$500-$1000", "$1000-$1500", "$1500 o más"]

print("Añadiendo datos al archivo Excel...")

for i in range(50):
  i+=1
  hoja[f"A{i+1}"] = arr_sexo[random.randint(0, 1)]
  hoja[f"B{i+1}"] = arr_cant_empleos[random.randint(0, 2)]
  trabajo = arr_trabajo_actual[random.randint(0, 1)]
  hoja[f"C{i+1}"] = trabajo
  if trabajo == "Si":
    hoja[f"D{i+1}"] = arr_hora_trabajo[random.randint(0, 2)]
    hoja[f"E{i+1}"] = arr_suedo[random.randint(0, 3)]

libro.save('Estadistica.xlsx')

print("Se guardo correctamente el archivo de Excel.")
