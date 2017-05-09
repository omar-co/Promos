import pandas as pd
import time

def cabecera_resumen_mot(row, format):
    worksheet.write('A' + str(row), 'Categoría', format)
    worksheet.write('B' + str(row), 'Cadena', format)
    worksheet.write('C' + str(row), '1.Si 2.No', format)
    worksheet.write('D' + str(row), '1.Si 2.Si', format)
    worksheet.write('E' + str(row), 'No El gerente no me permite ejecutar', format)
    worksheet.write('F' + str(row), 'No ha caido la promoción', format)
    worksheet.write('G' + str(row), 'No he ejecutado', format)
    worksheet.write('I' + str(row), 'No puedo señalizar', format)
    worksheet.write('J' + str(row), 'No tengo suficiente producto', format)
    worksheet.write('K' + str(row), 'Si', format)
    worksheet.write('L' + str(row), 'Categoría', format)
    worksheet.write('M' + str(row), 'Categoría', format)
    worksheet.write('N' + str(row), 'Categoría', format)
    worksheet.write('O' + str(row), 'Categoría', format)
    worksheet.write('P' + str(row), 'Categoría', format)
    worksheet.write('Q' + str(row), 'Categoría', format)
    worksheet.write('R' + str(row), 'Categoría', format)
    worksheet.write('S' + str(row), 'Categoría', format)
    worksheet.write('T' + str(row), 'Categoría', format)
    worksheet.write('U' + str(row), 'Ejecutado', format)
    worksheet.write('V' + str(row), 'NO Ejecutado', format)
    worksheet.write('W' + str(row), '% Ejecución', format)
    worksheet.write('X' + str(row), '% No ha caido la promoción', format)
    worksheet.write('Y' + str(row), '% No tengo suficiente producto', format)
    worksheet.write('Z' + str(row), '% No puedo señalizar', format)
    worksheet.write('AA' + str(row), '% No he ejecutado', format)

""" Ingresa Planes Promocionales"""

"""file_belleza = input('Ingrese el nombre del file con el Plan Promocional de Belleza.')
file_farmacia = input('Ingrese el nombre del file con el Plan Promocional de Farmacia.')
file_quimicos = input('Ingrese el nombre del file con el Plan Promocional de Quimicos.')"""

""" Ingresa lecturas de iniciativas"""

"""iniciativa_belleza_file = input('Ingrese el nombre del file con la lectura de Belleza.')
iniciativa_farmacia_file = input('Ingrese el nombre del file con la lectura de Farmacia.')
iniciativa_quimicos_file = input('Ingrese el nombre del file con la lectura de Quimicos.')"""

"""Ingresa Catálogo de Stores"""

# stores_file = input('Ingrese el nombre del file con el Catálogo de Tiendas')

start_time = time.time()

print('Leyendo Catálogo de Tiendas')
stores_file = 'stores_clean.xlsx'

""" Inicia lectura de stores """
stores = pd.read_excel(stores_file, 0)
print('Catálogo de Tiendas Cargado en memoria')
""" Inicia lectura de preguntas por categoria """

"""belleza = pd.read_excel(file_belleza, 0, 2)
farmacia = pd.read_excel(file_farmacia, 0, 2)
quimicos = pd.read_excel(file_quimicos, 0, 2) """

print('Leyendo Planes Promocionales')
belleza = pd.read_excel('Plan Promocional Belleza Mayo 2017.xlsx', 4, 2)
farmacia = pd.read_excel('Plan Promocional Farmacia Mayo 2017.xlsx', 0, 2)
quimicos = pd.read_excel('Plan Promocional Quimicos Mayo 2017.xlsx', 4, 2)
print('Planes Promocionales cargados en memoria')

print('Consolidando Planes Promocionales')
frames = [belleza, farmacia, quimicos]

preguntas = pd.concat(frames)

preguntas = preguntas[['Categoría', 'NOMBRE DE LA INICIATIVA']]

# iniciativas = pd.read_excel('iniciativas.xlsx', 0)

""" Inicia lectura de iniciativas """
"""iniciativa_belleza = pd.read_csv(iniciativa_belleza_file, encoding='mbcs')
iniciativa_farmacia = pd.read_csv(iniciativa_farmacia_file, encoding='mbcs')
iniciativa_quimicos = pd.read_csv(iniciativa_quimicos_file, encoding='mbcs') """

print('Leyendo Iniciativas')
iniciativa_belleza = pd.read_csv('lecturas_belleza.csv', encoding='mbcs')
iniciativa_farmacia = pd.read_csv('lecturas_farmacia.csv', encoding='mbcs')
iniciativa_quimicos = pd.read_csv('lecturas_quimicos.csv', encoding='mbcs')
print('Iniciativas vargadas en memoria')

print('Consolidando Iniciativas')
iniciativas = pd.concat([iniciativa_belleza, iniciativa_farmacia, iniciativa_quimicos])

iniciativas = iniciativas[
    ['Cadena', 'Región', 'Zona', 'Tienda ID', 'Nombre Tienda', 'Fecha Captura', 'Grupo Categorias',
     'Iniciativa', 'Apoyo', 'Respuesta Opc.Multiple/Texto Abierto']]

print('Exrtrayendo ID de preguntas')
preguntas_id = iniciativas['Apoyo'].str.extract("(?P<PreguntaID>\w{5}\d{1,3})(?P<Pregunta>¿.+)", expand=False)

print('Vinculando ID de preguntas con iniciativas')
iniciativas = iniciativas.join(preguntas_id)

preguntas = preguntas.set_index('NOMBRE DE LA INICIATIVA')

iniciativas = iniciativas.join(preguntas, on='Pregunta')

#  iniciativas.to_excel('prueba1.xlsx', 'demo')  # solo para testing

stores = stores[['Stores ID', '# Sucursal Cliente', 'Cadena', 'Formato', 'Nombre Tienda', 'Estatus de Tienda', 'Canal']]

stores = stores.drop_duplicates()

print('Vincualando iniciativas con Catálogo de Tiendas')
iniciativas = iniciativas.join(stores.set_index('Stores ID'), on='Tienda ID', rsuffix='_stores')

columnas_iniciativas = iniciativas.columns

print('Eliminando lecturas duplicadas')
iniciativas = iniciativas.sort_values(by='Fecha Captura')
print('Tomando últimos valores de lecturas')
iniciativas = iniciativas.drop_duplicates(['Nombre Tienda', 'Pregunta'], keep='last').values

iniciativas = pd.DataFrame(iniciativas, columns=columnas_iniciativas)

data = iniciativas[['Canal', 'Categoría', 'Cadena', 'Región', 'Zona', 'Tienda ID', 'Nombre Tienda', 'Fecha Captura',
                    'Iniciativa', 'Pregunta', 'Respuesta Opc.Multiple/Texto Abierto']]

print('Agrupando lectura de iniciativas')
order = data.groupby(['Canal', 'Categoría', 'Cadena', 'Pregunta', 'Respuesta Opc.Multiple/Texto Abierto'],
                     as_index=False).count()

order = order[['Canal', 'Categoría', 'Cadena', 'Pregunta', 'Respuesta Opc.Multiple/Texto Abierto', 'Zona']]

print('Ordenando por Categorías')
categorias = order.pivot_table(index=['Canal', 'Categoría', 'Cadena', 'Pregunta'],
                               columns='Respuesta Opc.Multiple/Texto Abierto', values='Zona', fill_value=0).reset_index(
    ['Canal', 'Categoría', 'Cadena', 'Pregunta'])

print('Calculando Ejecutados')
categorias = categorias.assign(Ejecutado=(categorias['Si'] + categorias['Si pero el precio no corresponde']))

print('Calculando NO Ejecutados')
categorias = categorias.assign(No_Ejecutado=(
    categorias['No ha caido la promoción'] + categorias['No he ejecutado'] + categorias['No puedo señalizar'] +
    categorias['No tengo suficiente producto']))

print('Calculando Porcentajes de Ejecución')
categorias = categorias.assign(
    Porcentaje_Ejecución=(categorias['Ejecutado'] / (categorias['Ejecutado'] + categorias['No_Ejecutado'])))

categorias = categorias.rename(columns={'No_Ejecutado': 'No Ejecutado', 'Porcentaje_Ejecución': '% Ejecución'})

# para categoria y preguntas
# falta hacer pivot para tener orden adecuado
# order[order['Categoría'] == 'Detergentes'].to_excel('Detergentes.xlsx','Detergentes')

print('Calculando Reporte MOT')
reporte_mot = data.groupby(['Categoría', 'Cadena', 'Respuesta Opc.Multiple/Texto Abierto'], as_index=False).count()

reporte_mot = reporte_mot.pivot_table(index=['Categoría', 'Cadena'], columns='Respuesta Opc.Multiple/Texto Abierto',
                                      values='Zona', fill_value=0)

reporte_mot = reporte_mot.reset_index()

reporte_mot = reporte_mot.assign(Ejecutado=(reporte_mot['Si'] + reporte_mot['Si pero el precio no corresponde']))

reporte_mot = reporte_mot.assign(No_Ejecutado=(
    reporte_mot['No ha caido la promoción'] + reporte_mot['No he ejecutado'] + reporte_mot['No puedo señalizar'] +
    reporte_mot['No tengo suficiente producto']))

reporte_mot = reporte_mot.assign(
    Porcentaje_Ejecución=(reporte_mot['Ejecutado'] / (reporte_mot['Ejecutado'] + reporte_mot['No_Ejecutado'])))

reporte_mot = reporte_mot.assign(
    Por_No_ha_caido_la_promocion=(reporte_mot['No ha caido la promoción'] / reporte_mot['No_Ejecutado']))

reporte_mot = reporte_mot.assign(
    Por_No_tengo_suficiente_producto=(reporte_mot['No tengo suficiente producto'] / reporte_mot['No_Ejecutado']))

reporte_mot = reporte_mot.assign(
    Por_No_puedo_senalizar=(reporte_mot['No puedo señalizar'] / reporte_mot['No_Ejecutado']))

reporte_mot = reporte_mot.assign(Por_No_he_ejecutado=(reporte_mot['No he ejecutado'] / reporte_mot['No_Ejecutado']))

reporte_mot = reporte_mot.rename(columns={'No_Ejecutado': 'No Ejecutado',
                                          'Porcentaje_Ejecución': '% Ejecución',
                                          'Por_No_ha_caido_la_promocion': '% No ha caido la promoción',
                                          'Por_No_tengo_suficiente_producto': '% No tengo suficiente producto',
                                          'Por_No_puedo_senalizar': '% No puedo señalizar',
                                          'Por_No_he_ejecutado': '% No he ejecutado'
                                          })

writter = pd.ExcelWriter('RepPromos.xlsx')

workbook = writter.book

print('Escribiendo Reporte MOT')

categorias_values = categorias.Categoría.unique()

formato_cabecera = workbook.add_format()
formato_cabecera.set_font_color('white')
formato_cabecera.set_align('center')
formato_cabecera.set_bg_color('#222B35')
formato_cabecera.set_bold()

formato_titulos = formato_cabecera

row = 2
for cat in categorias_values:
    reporte_mot[reporte_mot['Categoría'] == cat].to_excel(writter, 'Resumen MOT', startrow=row + 1, index=False, header=False)
    worksheet = writter.sheets['Resumen MOT']
    valores = reporte_mot.Cadena[reporte_mot['Categoría'] == cat].count()
    worksheet.merge_range('X' + str(row - 1) + ':AA' + str(row - 1), 'Razones para no ejecutar', formato_cabecera)
    worksheet.merge_range('X' + str(row) + ':Z' + str(row), 'Causal CT', formato_cabecera)
    worksheet.write('AA' + str(row), 'Causal SDO', formato_cabecera)
    cabecera_resumen_mot(row + 1, formato_cabecera)
    worksheet.set_row(row, 33.75)
    worksheet.conditional_format('W' + str(row + 1) + ':W' + str(row + valores + 1), {'type': '3_color_scale'})
    row += (valores + 4)

for cat in categorias_values:
    print('Calculando reporte por Categoría: ' + cat)
    categorias[categorias['Categoría'] == cat].to_excel(writter, cat, index=False)
    print('Escribiendo reporte de Categoría: ' + cat)

#categorias.to_excel(writter, 'categorias')

worksheet = writter.sheets['Resumen MOT']

format = workbook.add_format()
format.set_text_wrap()

porcentaje = workbook.add_format({'num_format': '0%'})

worksheet.set_column('C:V', None, None, {'hidden': True})
worksheet.set_column('W:AA', None, porcentaje)

worksheet.set_column(0, 0, 18.29)
worksheet.set_column(1, 1, 21.29)
worksheet.set_column(22, 26, 21)
# worksheet.set_row(2, 33.75, format)


print('Calculando Reportes por Cadena')
order[order['Cadena'] == 'AC Soriana'].pivot_table(index=['Canal', 'Categoría', 'Cadena', 'Pregunta'],
                                                   columns='Respuesta Opc.Multiple/Texto Abierto', values='Zona',
                                                   fill_value=0).reset_index().to_excel(writter, 'AC Soriana',
                                                                                        index=False)

order[order['Cadena'] == 'AD Chedraui'].pivot_table(index=['Canal', 'Categoría', 'Cadena', 'Pregunta'],
                                                    columns='Respuesta Opc.Multiple/Texto Abierto', values='Zona',
                                                    fill_value=0).reset_index().to_excel(writter, 'AD Chedraui',
                                                                                         index=False)

order[order['Cadena'] == 'AE Comercial Mexicana'].pivot_table(index=['Canal', 'Categoría', 'Cadena', 'Pregunta'],
                                                              columns='Respuesta Opc.Multiple/Texto Abierto',
                                                              values='Zona',
                                                              fill_value=0).reset_index().to_excel(writter,
                                                                                                   'AE Comercial Mexicana',
                                                                                                   index=False)

order[(order['Cadena'] == 'AB Bodega Aurrera') | (order['Cadena'] == 'AA Supercenter') | (order[
                                                                                              'Cadena'] == 'Superama')].pivot_table(
    index=['Canal', 'Categoría', 'Cadena', 'Pregunta'],
    columns='Respuesta Opc.Multiple/Texto Abierto', values='Zona',
    fill_value=0).reset_index().to_excel(writter, 'WM',
                                         index=False)
print('Guardando reporte por Cadena')

print('Generando File')

workbook.close()
writter.save()
print('File Generado')

print('Ejecutado en: ' + str(time.time() - start_time))

time.sleep(3)
