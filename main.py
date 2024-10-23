import pandas as pd
from gurobipy import Model, GRB
from gurobipy import quicksum as qsum
from datos.parametros_unicos import h, b, v, k, cd
from os import path
import time

# Se guarda el tiempo de inicio de la ejecución
start_time = time.time()

# Archivos excel con los datos de los parámetros del modelo
archivo_demanda_recursos = path.join("datos", "demanda_recursos.xlsx")
archivo_donaciones = path.join("datos", "donaciones.xlsx")
archivo_distancias = path.join("datos", "distancia_ciudades.xlsx")
archivo_personal_medico = path.join("datos", "personal_medico.xlsx")
archivo_centros = path.join("datos", "centros.xlsx")

# Se crean los conjuntos que representan los subíndices de los parámetros y variables del modelo
T = [sheet for sheet in pd.ExcelFile(archivo_donaciones).sheet_names if sheet != 'Totales'] # Conjunto de días
J = pd.read_excel(archivo_donaciones, sheet_name=T[0]).columns[2:]  # Conjunto de recursos
C = pd.read_excel(archivo_donaciones, sheet_name=T[0]).iloc[1:-1, 0]  # Conjunto de ciudades
P = range(1, h + 1)  # Conjunto de heridos
R = range(1, 4 + 1)  # Conjunto de centros médicos
M = range(1, 3 + 1)  # Conjunto de tipos de médicos

# Se crean diccionarios para almacenar los datos de los archivos excel, la llave de cada
# diccionario es una tupla con los subíndices correspondientes a cada parámetro
e = {}  # Demanda de recursos de los heridos
g = {}  # Donaciones de recursos
i = {}  # Capacidad de acopio de ciudades
d = {}  # Distancia ciudades
q = {}  # Capacidad del centro de atención médica r
qm = {}  # Capacidad de personal médico del centro de atención r
qh = {}  # Cantidad de heridos que el personal médico de tipo m puede atender por día.
cs = {}  # Costo diario de servicio por cada personal médico de tipo m

# Usando pandas se leen los datos de los distintos archivos excel que corresponden a 
# los parámetros del modelo y se almacenan en sus diccionarios respectivos

# definicion de e
df = pd.read_excel(archivo_demanda_recursos)
for j in df.index:
    # Obtener el nombre del recurso desde la columna correspondiente (por ejemplo, columna 0)
    resource_name = df.iloc[j, 0]  # Cambia el índice según la columna de los nombres de los recursos
    e[resource_name] = df.iloc[j, 3]  # Demanda de recurso j

# definicion de g
for t in T:
    df = pd.read_excel(archivo_donaciones, sheet_name=t)
    for c in df.index[:-1]:  # Ciudades (filas)
        city_name = df.iloc[c, 0]  # Cambia el índice según la columna de los nombres de las ciudades
        for j in df.columns[2:]:  # Recursos (columnas)
            # Asignamos el valor al diccionario g[j,t,c]
            g[j, t, city_name] = df.loc[c, j]  # Usar city_name en lugar de c


# definicion de i
df = pd.read_excel(archivo_centros, sheet_name="centros acopio")
for c in df.index[:-1]:
    city_name = df.iloc[c, 0]  # Cambia el índice según la columna de los nombres de las ciudades
    i[city_name] = df.iloc[c, 2]  # Capacidad del centro de acopio de la ciudad c

# definicion de d
df = pd.read_excel(archivo_distancias)
for index in df.index[:-2]:
    ciudad = df.iloc[index, 0]  # Suponiendo que los nombres de las ciudades están en la primera columna
    distancia = df.iloc[index, 1]  # Suponiendo que las distancias están en la segunda columna
    d[ciudad] = distancia  # Usa el nombre de la ciudad como clave

# definicion de q
df = pd.read_excel(archivo_centros, sheet_name="centros medicos")
for r in df.index:
    q[r] = df.iloc[r, 1]  # Capacidad del centro de atención médica r

# definicion de qm
df = pd.read_excel(archivo_centros, sheet_name="centros medicos")
for r in df.index:
    qm[r] = df.iloc[r, 2]  # Capacidad de personal médico del centro de atención r

# definicion de qh
df = pd.read_excel(archivo_personal_medico)
for m in df.index:
    qh[m] = df.iloc[m, 1]  # Cantidad de heridos que el personal médico de tipo m puede atender por día.

# definicion de cs
df = pd.read_excel(archivo_personal_medico)
for m in df.index:
    cs[m] = df.iloc[m, 3]  # Costo de servicio del personal de tipo m por día


# Se instancia el modelo
modelo = Model()

# Se establece un tiempo límite de 30 minutos
modelo.setParam("TimeLimit", 30 * 60)

# Se instancian variables de decision

print("Creando variables de decisión...")
# Cantidad de recurso médico j disponible al inicio del día t
X = modelo.addVars(J, T, vtype=GRB.CONTINUOUS, name="X", lb=0)

# Cantidad del recurso médico j almacenado el día t en el almacén del centro de acopio de la ciudad c
I = modelo.addVars(J, T, C, vtype=GRB.CONTINUOUS, name="I", lb=0)

# Cantidad de recurso médico j que sale al final del día t desde la ciudad c
U = modelo.addVars(J, T, C, vtype=GRB.CONTINUOUS, name="U", lb=0)

# Cantidad de personal médico m atendiendo en el centro de atención médica r el día t en Valparaíso
N = modelo.addVars(M, R, T, vtype=GRB.INTEGER, name="N", lb=0)

# Cantidad de camiones utilizados en el transporte de recursos médicos desde la ciudad c, el día t
Y = modelo.addVars(C, T, vtype=GRB.INTEGER, name="Y", lb=0)

# Variable binaria que indica si el herido p es atendido el día t en el centro de atención médica r
B = modelo.addVars(P, R, T, vtype=GRB.BINARY, name="B")

# Variable binaria que indica si el herido p está en el centro de atención médica r en el tiempo t
Rc = modelo.addVars(P, R, T, vtype=GRB.BINARY, name="R")

# Variable binaria que indica si se cumplió la demanda del insumo médico j del herido p en el centro de atención médica r en el tiempo t
A = modelo.addVars(P, J, R, T, vtype=GRB.BINARY, name="A")

# Cantidad del recurso j desechada en la ciudad c en el tiempo t
D = modelo.addVars(C, J, T, vtype=GRB.CONTINUOUS, name="D", lb=0)

print(f"Tiempo para la creación de variables: {time.time() - start_time:.2f} segundos")

# Se agregan las variables al modelo
modelo.update()

# Se definen las restricciones del modelo (descripciones disponibles en el informe adjunto)


modelo.addConstr(qsum(cs[m] * N[m,r,t] for m in M for r in R for t in T) +
                 qsum(cd * d[c] * Y[c,t] for c in C for t in T) <= b, name = "R1")

modelo.addConstrs((B[p,r,t] <= (Rc[p,r,t] + qsum(A[p,j,r,t] for j in J)) / (1 + len(J)) for p in P for r in R for t in T), name = "R2")

modelo.addConstrs((qsum(B[p,r,t] for t in T) <= 1 for p in P for r in R), name = "R3")

modelo.addConstrs((qsum(Rc[p,r,t] for r in R) <= 1 for p in P for t in T), name = "R4")

for j in J:
    for t in T:
        if t == "dia1":
            modelo.addConstr(X[j,t] == 0, name = "R5")
        else:
            dia = T[T.index(t) - 1]
            modelo.addConstr(X[j,t] == qsum(U[j,dia,c] for c in C) + X[j,dia] - e[j] * qsum(A[p,j,r,dia] for p in P for r in R), name = "R5")

modelo.addConstrs((e[j] * qsum(A[p,j,r,t] for p in P for r in R) <= X[j,t] for j in J for t in T), name = "R6")

for j in J:
    for c in C:
        for t in T:
            if t == "dia1":
                modelo.addConstr(I[j,t,c] == g[j,t,c], name = "R7")
            else:
                dia = T[T.index(t) - 1]
                modelo.addConstr(I[j,t,c] == I[j,dia,c] + g[j,t,c] - U[j,t,c] - D[c,j,t], name = "R7")

modelo.addConstrs((qsum(I[j,t,c] for j in J) <= i[c] for c in C for t in T), name = "R8")

modelo.addConstrs((qsum(X[j,t] for j in J) <= v for t in T), name = "R9")

modelo.addConstrs((qsum(U[j,t,c] for j in J) <= k * Y[c,t] for c in C for t in T), name = "R10")

modelo.addConstr(qsum(B[p,r,t] for p in P for r in R for t in T) <= h, name = "R11")

modelo.addConstrs((qsum(B[p,r,t] for p in P) <= qsum(N[m,r,t] * qh[m] for m in M) for r in R for t in T), name = "R12")

modelo.addConstrs((qsum(N[m,r,t] for m in M) <= qm[r] for r in R for t in T), name = "R13")

modelo.addConstrs((qsum(Rc[p,r,t] for p in P) <= q[r] for r in R for t in T), name = "R14")

# Se define la función objetivo del modelo
modelo.setObjective(qsum(B[p,r,t] for p in P for r in R for t in T), GRB.MAXIMIZE)

# Se optimiza el modelo
modelo.optimize()

# Se escriben los resultados del modelo en un archivo de Excel: "resultados.xlsx"
var_names = []
var_values = []
for var in m.getVars():
    if var.varName.startswith("X") and var.x > 0:
        var_names.append(str(var.varName))
        var_values.append(var.x)

df = pd.DataFrame({"Nombre": var_names, "Valor": var_values})
df.to_excel("Resultados.xlsx", index = False)        
