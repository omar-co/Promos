import pandas as pd
import stores_cleaner as sc


stores = sc.clean_stores(create_excel=False)

print('Leyendo Iniciativas')
df1 = pd.read_csv('lecturas_iniciativa_006139(2).csv', encoding='mbcs', index_col=4)
print('Iniciativas cargadas en memoria')

print('Eliminando Duplicados')
df1 = df1.drop_duplicates()

print('Filtrando Columnas')
df1 = df1[['Cadena', 'Formato', 'Región', 'Zona', 'Nombre Tienda', 'Fecha Captura', 'Apoyo',
          'Respuesta Opc.Multiple/Texto Abierto']]

# stores = stores[['# Sucursal Cliente', 'Cadena', 'Formato', 'Nombre Tienda', 'Estatus de Tienda', 'Canal']]

# stores = stores.drop_duplicates()

print('Ligando Catálogo de Tiendas con Iniciativas')
df3 = df1.join(stores.set_index('Stores ID'), rsuffix='_stores')

print('Filtrando ultimos registros')

df3['Max Date'] = df3.groupby(['Nombre Tienda', 'Apoyo'])['Fecha Captura'].transform('max')

new = pd.DataFrame(df3.groupby(['Canal', 'Cadena', 'Región', 'Fecha Captura',
                                'Nombre Tienda', 'Apoyo', 'Respuesta Opc.Multiple/Texto Abierto'])['Max Date'].max())

print('Generando prueba1.xlsx')
writter = pd.ExcelWriter('prueba1.xlsx')

print('Guardando prueba1.xlsx')
new.to_excel(writter, 'df3')

writter.save()
