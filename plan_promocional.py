import pandas as pd


def leer_plan(file):
    data = pd.read_excel(file, 0, 2)
    return data


# file_belleza = input('Ingrese el nombre del file con el Plan Promocional de Belleza.')
# file_farmacia = input('Ingrese el nombre del file con el Plan Promocional de Farmacia.')
# file_quimicos = input('Ingrese el nombre del file con el Plan Promocional de Quimicos.')
#

# belleza = pd.read_excel(file_belleza, 0, 2)
# farmacia = pd.read_excel(file_farmacia, 0, 2)
# quimicos = pd.read_excel(file_quimicos, 0, 2)

belleza = pd.read_excel('Plan Promocional Belleza Mayo 2017.xlsx', 4, 2)
farmacia = pd.read_excel('Plan Promocional Farmacia Mayo 2017.xlsx', 0, 2)
quimicos = pd.read_excel('Plan Promocional Quimicos Mayo 2017.xlsx', 4, 2)

frames = [belleza, farmacia, quimicos]

preguntas = pd.concat(frames)

preguntas = preguntas[['Categoría', 'NOMBRE DE LA INICIATIVA']]

# iniciativas = pd.read_excel('iniciativas.xlsx', 0)

iniciativa_belleza_file = input('Ingrese el nombre del file con la lectura de Belleza.')
iniciativa_farmacia_file = input('Ingrese el nombre del file con la lectura de Farmacia.')
iniciativa_quimicos_file = input('Ingrese el nombre del file con la lectura de Quimicos.')

# iniciativa_belleza = pd.read_csv(iniciativa_belleza_file, encoding='mbcs')
# iniciativa_farmacia = pd.read_csv(iniciativa_farmacia_file, encoding='mbcs')
# iniciativa_quimicos = pd.read_csv(iniciativa_quimicos_file, encoding='mbcs')

iniciativa_belleza = pd.read_csv('lecturas_belleza.csv', encoding='mbcs')
iniciativa_farmacia = pd.read_csv('lecturas_farmacia.csv', encoding='mbcs')
iniciativa_quimicos = pd.read_csv('lecturas_quimicos.csv', encoding='mbcs')

iniciativas = pd.concat([iniciativa_belleza, iniciativa_farmacia, iniciativa_quimicos])

iniciativas = iniciativas[
    ['Cadena', 'Región', 'Zona', 'Tienda ID', 'Nombre Tienda', 'Fecha Captura', 'Grupo Categorias',
     'Iniciativa', 'Apoyo', 'Respuesta Opc.Multiple/Texto Abierto']]

preguntas_id = iniciativas['Apoyo'].str.extract("(?P<PreguntaID>\w{5}\d{1,3})(?P<Pregunta>¿.+)", expand=False)

iniciativas = iniciativas.join(preguntas_id)

preguntas = preguntas.set_index('NOMBRE DE LA INICIATIVA')

iniciativas = iniciativas.join(preguntas, on='Pregunta', rsuffix='_preg')

iniciativas.to_excel('prueba1.xlsx', 'demo')
