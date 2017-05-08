import pandas as pd
import stores_cleaner as sc

# stores = sc.clean_stores(create_excel=False)

print('Leyendo Iniciativas')
df1 = pd.read_excel('iniciativas.xlsx', index_col=4, sheetname=0)

# df1 = pd.read_csv('lecturas_iniciativa_006139(2).csv', encoding='mbcs', index_col=4)
print('Iniciativas cargadas en memoria')

print('Eliminando Duplicados')
df1 = df1.drop_duplicates()

print('Filtrando Columnas')
df1 = df1[['Cadena', 'Formato', 'Región', 'Zona', 'Nombre Tienda', 'Fecha Captura', 'Apoyo',
          'Respuesta Opc.Multiple/Texto Abierto']]

# stores = stores[['# Sucursal Cliente', 'Cadena', 'Formato', 'Nombre Tienda', 'Estatus de Tienda', 'Canal']]

# stores = stores.drop_duplicates()

print('Ligando Catálogo de Tiendas con Iniciativas')
# df3 = df1.join(stores.set_index('Stores ID'), rsuffix='_stores')

df3 = df1

print('Filtrando ultimos registros')

sort = df3.sort_values(by='Fecha Captura')

df3 = sort.drop_duplicates(['Nombre Tienda', 'Apoyo'], keep='last').values

final = pd.DataFrame(df3, columns=sort.columns)

print('Generando prueba1.xlsx')
writter = pd.ExcelWriter('prueba1.xlsx')

print('Guardando prueba1.xlsx')

final.to_excel(writter, 'final')

sort.to_excel(writter, 'df3')

writter.save()
