import pandas as pd
import time

start_time = time.time()


def obtener_categoria(categoria):
    return order[order['Categoría'] == categoria]


# data = pd.read_excel('Reporte de Promociones.xlsx', 15)
data = pd.read_excel('pruebafiltro.xlsx', 0)

data = data[['Canal', 'Categoría', 'Cadena', 'Región', 'Zona', 'Tienda ID', 'Nombre Tienda', 'Fecha Captura',
             'Iniciativa', 'Apoyo', 'Respuesta Opc.Multiple/Texto Abierto']]

order = data.groupby(['Canal', 'Categoría', 'Cadena', 'Apoyo', 'Respuesta Opc.Multiple/Texto Abierto'],
                     as_index=False).count()

order = order[['Canal', 'Categoría', 'Cadena', 'Apoyo', 'Respuesta Opc.Multiple/Texto Abierto', 'Zona']]

categorias = order.pivot_table(index=['Canal', 'Categoría', 'Cadena', 'Apoyo'],
                               columns='Respuesta Opc.Multiple/Texto Abierto', values='Zona', fill_value=0).reset_index(
    ['Canal', 'Categoría', 'Cadena', 'Apoyo'])

categorias = categorias.assign(Ejecutado=(categorias['Si'] + categorias['Si pero el precio no corresponde']))

categorias = categorias.assign(No_Ejecutado=(
    categorias['No ha caido la promoción'] + categorias['No he ejecutado'] + categorias['No puedo señalizar'] +
    categorias['No tengo suficiente producto']))

categorias = categorias.assign(
    Porcentaje_Ejecución=(categorias['Ejecutado'] / (categorias['Ejecutado'] + categorias['No_Ejecutado'])))

categorias = categorias.rename(columns={'No_Ejecutado': 'No Ejecutado', 'Porcentaje_Ejecución': '% Ejecución'})

# para categoria y preguntas
# falta hacer pivot para tener orden adecuado
# order[order['Categoría'] == 'Detergentes'].to_excel('Detergentes.xlsx','Detergentes')


reporte_mot = data.groupby(['Categoría', 'Cadena', 'Respuesta Opc.Multiple/Texto Abierto'], as_index=False).count()

reporte_mot = reporte_mot.pivot_table(index=['Categoría', 'Cadena'], columns='Respuesta Opc.Multiple/Texto Abierto',
                                      values='Zona', fill_value=0)

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

reporte_mot.to_excel(writter, 'Resumen MOT')

categorias_values = categorias.Categoría.unique()

for cat in categorias_values:
    categorias[categorias['Categoría'] == cat].to_excel(writter, cat)

order[order['Cadena'] == 'AC Soriana'].pivot_table(index=['Canal', 'Categoría', 'Cadena', 'Apoyo'],
                                                   columns='Respuesta Opc.Multiple/Texto Abierto', values='Zona',
                                                   fill_value=0).reset_index().to_excel(writter, 'AC Soriana',
                                                                                        index=False)

order[order['Cadena'] == 'AD Chedraui'].pivot_table(index=['Canal', 'Categoría', 'Cadena', 'Apoyo'],
                                                    columns='Respuesta Opc.Multiple/Texto Abierto', values='Zona',
                                                    fill_value=0).reset_index().to_excel(writter, 'AD Chedraui',
                                                                                         index=False)

order[order['Cadena'] == 'AE Comercial Mexicana'].pivot_table(index=['Canal', 'Categoría', 'Cadena', 'Apoyo'],
                                                              columns='Respuesta Opc.Multiple/Texto Abierto',
                                                              values='Zona',
                                                              fill_value=0).reset_index().to_excel(writter,
                                                                                                   'AE Comercial Mexicana',
                                                                                                   index=False)

order[(order['Cadena'] == 'AB Bodega Aurrera') | (order['Cadena'] == 'AA Supercenter') | (order[
                                                                                              'Cadena'] == 'Superama')].pivot_table(
    index=['Canal', 'Categoría', 'Cadena', 'Apoyo'],
    columns='Respuesta Opc.Multiple/Texto Abierto', values='Zona',
    fill_value=0).reset_index().to_excel(writter, 'WM',
                                         index=False)

# reporte_mot.ix['Detergentes', 'Superama'].to_frame().to_excel(writter, 'Superama')

writter.save()

print('Ejecutado en: ' + str(time.time() - start_time))

time.sleep(3)
