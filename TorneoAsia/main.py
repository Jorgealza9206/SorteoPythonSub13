import random
#import xlsxwriter

# Define los equipos en cada bombo
bombo1 = ["Uzbekistán", "Irán", "Corea del Sur", "Japón"]
bombo2 = ["Australia", "Emiratos Árabes Unidos", "China", "Irak"]
bombo3 = ["Malasia", "Tailandia", "Vietnam", "Siria"]
bombo4 = ["Jordania", "Qatar", "Tayikistán", "Indonesia"]

# Define los grupos
grupos = {}
for i in range(1, 5):
    grupos[f"Grupo {chr(64+i)}"] = []

print(grupos)

#Realizamos el sorteo del primer bombo

#A1 Anfitrión
equipo = bombo1[0]
grupos[f"Grupo {chr(65)}"].append(equipo)
bombo1.remove(equipo)

for i in range(3):
    equipo = random.choice(bombo1)
    grupos[f"Grupo {chr(66 + i)}"].append(equipo)
    bombo1.remove(equipo)

print(grupos)

bombos = [bombo2,bombo3,bombo4]

#Realizamos el sorteo
for i in range(3):
    for j in range(4):
        equipo = random.choice(bombos[i])
        grupos[f"Grupo {chr(65 + j)}"].append(equipo)
        bombos[i].remove(equipo)

print(grupos)

try:
    ruta = "../AsiáticoUzbekistán.xlsx"
    archivo = xlsxwriter.Workbook(ruta)

    hoja = archivo.add_worksheet()
    hoja.title = "Asiático Uzbekistán 2"

    fila = 0
    columna = 0

    for grupo, paises in grupos.items():
        hoja.write(fila, columna, grupo)
        for pais in paises:
            fila += 1
            hoja.write(fila, columna, pais)
        columna += 1
        fila = 0

    archivo.close()

except:
    print("Descarga la librería de excel")

