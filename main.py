import pandas as pd
from gurobipy import Model, GRB
from gurobipy import quicksum as qsum
from datos.parametros_unicos import h, b, v, k, cd

# Archivos excel con los datos de los parámetros del modelo
archivo_demanda_recursos = "demanda_recursos.xlsx"
archivo_donaciones = "donaciones.xlsx"
archivo_distancias = "distancia_ciudades.xlsx"

# Se crean los conjuntos que representan los subíndices de los parámetros y variables del modelo
ciudades = pd.ExcelFile(archivo_demanda_recursos).sheet_names  # Conjunto de ciudades

donaciones_ciudad_1 = pd.read_excel(archivo_demanda_recursos, sheet_name=ciudades[0])

recursos = donaciones_ciudad_1.columns.tolist()  # Conjunto de recursos
dias = donaciones_ciudad_1.index.tolist()  # Conjunto de días
heridos = range(1, h + 1)  # Conjunto de heridos
centros_medicos = range(1, 3 + 1)  # Conjunto de centros médicos
tipo_medicos = range(1, 4 + 1)  # Conjunto de tipos de médicos

# Se crean diccionarios para almacenar los datos de los archivos excel, la llave de cada
# diccionario es una tupla con los subíndices correspondientes a cada parámetro
e = {}  # Demanda de recursos
g = {}  # Donaciones de recursos
i = {}  # Capacidad de acopio de ciudades
d = {}  # Distancia ciudades
c = {}  # Costo servicio de tipos de médicos

# Usando pandas se leen los datos de los distintos archivos excel que corresponden a 
# los parámetros del modelo y se almacenan en sus diccionarios respectivos
df = pd.read_excel(archivo_demanda_recursos, index_col=0)
for p in df.columns:
    for j in df.index:
        e[p,j] = df.loc[p,j] # La hoja j en la fila t y la columna i

ciudades = pd.ExcelFile(archivo_donaciones).sheet_names
for i in ciudades:
    df = pd.read_excel(archivo_donaciones, sheet_name=j)
    df.index += 1 # Para que los índices de las filas empiecen en 1 y no en 0 y calcen con los de la lista dias
    for j in df.columns:
        for t in df.index:
            g[i,j,t] = df.loc[t,i] # La hoja j en la fila t y la columna i

# Repetir para el resto de parámetros

# Se instancia el modelo
modelo = Model()

# Se establece un tiempo límite de 30 minutos
modelo.setParam("TimeLimit", 30 * 60)

# Se instancian variables de decision

# Cantidad de recurso médico j disponible al inicio del día t
X = modelo.addVars(recursos, dias, vtype=GRB.CONTINUOUS, name="X", lb=0)

# Cantidad del recurso médico j almacenado el día t en el almacén del centro de acopio de la ciudad c
I = modelo.addVars(recursos, dias, ciudades, vtype=GRB.CONTINUOUS, name="I", lb=0)

# Cantidad de recurso médico j que sale al final del día t desde la ciudad c
U = modelo.addVars(recursos, dias, ciudades, vtype=GRB.CONTINUOUS, name="U", lb=0)

# Cantidad de personal médico m atendiendo en el centro de atención médica r el día t en Valparaíso
N = modelo.addVars(tipo_medicos, centros_medicos, dias, vtype=GRB.INTEGER, name="N", lb=0)

# Cantidad de camiones utilizados en el transporte de recursos médicos desde la ciudad c, el día t
Y = modelo.addVars(ciudades, dias, vtype=GRB.INTEGER, name="Y", lb=0)

# Variable binaria que indica si el herido p es atendido el día t en el centro de atención médica r
B = modelo.addVars(heridos, centros_medicos, dias, vtype=GRB.BINARY, name="B")

# Variable binaria que indica si el herido p está en el centro de atención médica r en el tiempo t
R = modelo.addVars(heridos, centros_medicos, dias, vtype=GRB.BINARY, name="R")

# Variable binaria que indica si se cumplió la demanda del insumo médico j del herido p en el centro de atención médica r en el tiempo t
A = modelo.addVars(heridos, recursos, centros_medicos, dias, vtype=GRB.BINARY, name="A")

# Cantidad del recurso j desechada en la ciudad c en el tiempo t
D = modelo.addVars(ciudades, recursos, dias, vtype=GRB.CONTINUOUS, name="D", lb=0)


# Se agregan las variables al modelo
modelo.update()

# Se definen las restricciones del modelo (descripciones disponibles en el informe adjunto)

modelo.addConstrs(qsum(c[m]*N[m,r,t] for m in tipo_medicos for r in centros_medicos for t in dias) +
                  qsum(cd*d[c]*Y[c,t] for c in ciudades for t in dias), name = "R1")


# Se define la función objetivo del modelo
# quicksum(produccion[i,t]*y[i,t] + almacenamiento[i,t]*z[i,t] for i in frutas for t in dias), GRB.MINIMIZE
modelo.setObjective(funcion_objetivo, GRB.MINIMIZE)

# Se optimiza el modelo
modelo.optimize()


# Se escriben los resultados del modelo en un archivo de Excel: "resultados.xlsx"

# var_names = []
# var_values = []
# for var in m.getVars():
#     if var.varName.startswith("X") and var.x > 0:
#         var_names.append(str(var.varName))
#         var_values.append(var.x)

# df = pd.DataFrame({"Nombre": var_names, "Valor": var_values})
# df.to_excel("Resultados.xlsx", index = False)        
