import numpy
import openpyxl


excel_document = openpyxl.load_workbook('Muestra.xlsx')
sheet = excel_document['Hoja 1']

cant_personas_renun = 0
noproductivo = 0
infeliz = 0
sinambicion = 0
vivelejos = 0
tienefamilia = 0

for persona in sheet.iter_cols(min_col=2, max_col=9):
    value = persona[9].value
    if value == 1: #Si la persona renunció
        cant_personas_renun += 1
        line = 0
        for preguntas in persona:
            line += 1
            if line == 2:
                noproductivo += preguntas.value
            elif line == 3:
                infeliz += preguntas.value
            elif line == 4:
                sinambicion += preguntas.value
            elif line == 5:
                vivelejos += preguntas.value
            elif line == 6:
                tienefamilia += preguntas.value

            #print('%s: cell.value=%s' % (cell, cell.value))


porc_noproductivo = noproductivo * 100 / cant_personas_renun
porc_infeliz = infeliz * 100 / cant_personas_renun
porc_sinambicion = sinambicion * 100 / cant_personas_renun
porc_vivelejos = vivelejos * 100 / cant_personas_renun
porc_tienefamilia = tienefamilia * 100 / cant_personas_renun


print('No Productivos: %s' % (porc_noproductivo))
print('No es Feliz: %s' % (porc_infeliz))
print('No tiene ambición: %s' % (porc_sinambicion))
print('Vive lejos: %s' % (porc_vivelejos))
print('Tiene Familia: %s' % (porc_tienefamilia))


noproductivo = noproductivo / cant_personas_renun
infeliz = infeliz / cant_personas_renun
sinambicion = sinambicion / cant_personas_renun
vivelejos = vivelejos / cant_personas_renun
tienefamilia = tienefamilia / cant_personas_renun


opciones = ['No Productivos', 'No es Feliz', 'No tiene ambición', 'Vive lejos', 'Tiene Familia']
probabilidades = [noproductivo, infeliz, sinambicion, vivelejos, tienefamilia]

eleccion = numpy.random.choice(opciones, p=probabilidades)
print()
print('Es muy probable que un empleado se vaya de la empresa porque: %s' % (eleccion))
