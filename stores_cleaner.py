import pandas as pd

def clean_stores(file='store.xlsx', create_excel=True):
    print('Leyendo ' + file)
    df1 = pd.read_excel(file, 0)
    print(file + ' cargado en memoria')

    print('Filtrando Columnas')
    df1 = df1[['Stores ID', '# Sucursal Cliente',
           'Cadena', 'Formato', 'Nombre Tienda', 'Region', 'Zona', 'Area Nielsen', 'Estatus de Tienda', 'Canal']]

    print('Filtrando Regiones')
    df2 = df1[(df1.Region.str.contains(r'^(region)\s\d{1,2}$') | df1.Region.str.contains('^DPP1$'))]

    print('Eliminando Duplicados')
    df2['Stores ID'].drop_duplicates()

    if create_excel:
        print('Generando stores_clean_xlsx')
        df2.to_excel('stores_clean.xlsx', 'stores', index=None)

    print(str(df2['Stores ID'].count()) + ' Tiendas Filtradas Correctamente')
    return df2

