import random
from openpyxl import Workbook
from openpyxl.styles import Font
import subprocess

bombo1 = ["Japón", "Corea del Sur", "Australia", "China", "Irán", "Indonesia", "Tailandia", "Vietnam", "Irak", "Emiratos Árabes Unidos"]
bombo2 = ["Siria", "Malasia", "Qatar", "Arabia Saudita", "Bahrein", "Kuwait", "Omán", "Jordania", "Palestina", "Yemen"]
bombo3 = ["Brunéi", "Laos", "Mongolia", "Macao", "Hong Kong", "Islas Marianas del Norte", "Timor Oriental", "Birmania", "Camboya", "Tayikistán"]
bombo4 = ["Taiwán", "Turkmenistán", "Bangladés", "Bután", "India", "Maldivas", "Nepal", "Sri Lanka", "Filipinas", "Singapur"]
bombo5 = ["Kirguistán", "Guam", "Líbano", "Afganistán"]

grupos = {
    'A': [],
    'B': [],
    'C': [],
    'D': [],
    'E': [],
    'F': [],
    'G': [],
    'H': [],
    'I': [],
    'J': []
}

random.shuffle(bombo1)
for i, equipo in enumerate(bombo1):
    grupo = chr(ord('A') + i)
    grupos[grupo].append(equipo)

random.shuffle(bombo2)
for i, equipo in enumerate(bombo2):
    grupo = chr(ord('J') - i)
    grupos[grupo].append(equipo)

random.shuffle(bombo3)
for i, equipo in enumerate(bombo3):
    grupo = chr(ord('A') + i)
    grupos[grupo].append(equipo)

random.shuffle(bombo4)
for i, equipo in enumerate(bombo4):
    grupo = chr(ord('J') - i)
    grupos[grupo].append(equipo)

random.shuffle(bombo5)
for i, equipo in enumerate(bombo5):
    grupo = chr(ord('A') + i)
    grupos[grupo].append(equipo)

for grupo, lista in grupos.items():
    print("Grupo " + grupo)
    for i, pais in enumerate(lista, start=1):
        print(f'{i}. {pais}')

# Crear un nuevo libro de Excel
libro = Workbook()

# Seleccionar la hoja activa
hoja = libro.active

for grupo, paises in grupos.items():
    celda = hoja[f'{grupo}1']
    celda.value = f'Grupo {grupo}'
    celda.font = Font(bold=True)
    columna = ord(grupo) - ord('A') + 1
    fila = 2
    for pais in paises:
        celda = hoja.cell(row=fila, column=columna)
        celda.value = pais
        fila += 1

# Guardar el libro de Excel
libro.save("sorteoAsia.xlsx")
subprocess.run(f'start "" "sorteoAsia.xlsx"', shell=True)