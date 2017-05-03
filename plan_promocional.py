import pandas as pd


def leer_plan(file):
    data = pd.read_excel(file, 4, 2)
    return data


# file = input('Ingrese el file con el Plan Promocional de Belleza.')

# TODO crear vaiable farmacia
# TODO crear vaiable quimicos

belleza = leer_plan('Plan Promocional Belleza Mayo 2017.xlsx')

# TODO  leer Farmacia
# TODO  leer Quimicos

# TODO unir farmacias quimicos y belleza = preguntas

# TODO limpiar columnas preguntas

iniciativas = pd.read_excel('iniciativas.xlsx', 0)

iniciativas = iniciativas[
    ['Cadena', 'Región', 'Zona', 'Tienda ID', 'Nombre Tienda', 'Fecha Captura', 'Grupo Categorias',
     'Iniciativa', 'Apoyo', 'Respuesta Opc.Multiple/Texto Abierto']]

preguntas_id = iniciativas['Apoyo'].str.extract("(?P<PreguntaID>\w{5}\d{1,3})(?P<Pregunta>¿.+)", expand=False)

iniciativas = iniciativas.join(preguntas_id)

# TODO cambiar belleza por preguntas
belleza.set_index('NOMBRE DE LA INICIATIVA')

# TODO cambiar belleza por preguntas
iniciativas = iniciativas.join(belleza, on='Pregunta', rsuffix='_preg')
