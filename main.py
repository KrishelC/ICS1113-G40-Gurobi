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
R = pd.read_excel(archivo_centros, sheet_name="centros medicos").set_index("Centro médico").index  # Conjunto de centros médicos
M = pd.read_excel(archivo_personal_medico).set_index("Personal médico").index  # Conjunto de tipos de médicos

# Se crean diccionarios para almacenar los datos de los archivos excel, la llave de cada
# diccionario es una tupla con los subíndices correspondientes a cada parámetro
e = {}  # Demanda de recursos de los heridos
g = {}  # Donaciones de recursos
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
    nombre_recurso = df.iloc[j, 0]  # Cambia el índice según la columna de los nombres de los recursos
    e[nombre_recurso] = df.iloc[j, 3]  # Demanda de recurso j

# definicion de g
for t in T:
    df = pd.read_excel(archivo_donaciones, sheet_name=t)
    for c in df.index[:-1]:  # Ciudades (filas)
        nombre_ciudad = df.iloc[c, 0]  # Cambia el índice según la columna de los nombres de las ciudades
        for j in df.columns[2:]:  # Recursos (columnas)
            g[j, t, nombre_ciudad] = df.loc[c, j]

# definicion de d
df = pd.read_excel(archivo_distancias)
for index in df.index[:-2]:
    ciudad = df.iloc[index, 0]  # Suponiendo que los nombres de las ciudades están en la primera columna
    distancia = df.iloc[index, 1]  # Suponiendo que las distancias están en la segunda columna
    d[ciudad] = distancia  # Usa el nombre de la ciudad como clave

# definicion de q
df = pd.read_excel(archivo_centros, sheet_name="centros medicos").set_index("Centro médico")
for r in df.index:
    q[r] = df.iloc[df.index.get_loc(r), 0]  # Capacidad del centro de atención médica r

# definicion de qm
df = pd.read_excel(archivo_centros, sheet_name="centros medicos").set_index("Centro médico")
for r in df.index:
    qm[r] = df.iloc[df.index.get_loc(r), 1]  # Capacidad de personal médico del centro de atención r

# definicion de qh
df = pd.read_excel(archivo_personal_medico).set_index("Personal médico")
for m in df.index:
    qh[m] = df.iloc[df.index.get_loc(m), 0]  # Cantidad de heridos que el personal médico de tipo m puede atender por día.

# definicion de cs
df = pd.read_excel(archivo_personal_medico).set_index("Personal médico")
for m in df.index:
    cs[m] = df.iloc[df.index.get_loc(m), 2]  # Costo de servicio del personal de tipo m por día

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

print(f"Tiempo para la creación de variables: {time.time() - start_time:.2f} segundos")

# Se agregan las variables al modelo
modelo.update()

# Se definen las restricciones del modelo (descripciones disponibles en el informe adjunto)


modelo.addConstr(qsum(cs[m] * N[m,r,t] for m in M for r in R for t in T) +
                 qsum(cd * d[c] * Y[c,t] for c in C for t in T) <= b, name = "R1")

modelo.addConstrs((B[p,r,t] <= Rc[p,r,t] for p in P for r in R for t in T), name = "R2")

modelo.addConstrs((B[p,r,t] <= (qsum(A[p,j,r,t] for j in J) / len(J)) for p in P for r in R for t in T), name = "R3")

modelo.addConstrs((qsum(B[p,r,t] for t in T) <= 1 for p in P for r in R), name = "R4")

modelo.addConstrs((qsum(Rc[p,r,t] for r in R) <= 1 for p in P for t in T), name = "R5")

for j in J:
    for t in T:
        if t == "dia1":
            modelo.addConstr(X[j,t] == 0, name = "R6")
        else:
            dia = T[T.index(t) - 1]
            modelo.addConstr(X[j,t] == qsum(U[j,dia,c] for c in C) + X[j,dia] - e[j] * qsum(A[p,j,r,dia] for p in P for r in R), name = "R6")

modelo.addConstrs((e[j] * qsum(A[p,j,r,t] for p in P for r in R) <= X[j,t] for j in J for t in T), name = "R7")

for j in J:
    for c in C:
        for t in T:
            if t == "dia1":
                modelo.addConstr(I[j,t,c] == g[j,t,c], name = "R8")
            else:
                dia = T[T.index(t) - 1]
                modelo.addConstr(I[j,t,c] == I[j,dia,c] + g[j,t,c] - U[j,t,c], name = "R8")

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

# Se instancia la función objetivo y se optimiza el problema
print("-"*10 + "Manejo de las soluciones" + "-"*10)

# Se imprime el valor objetivo
print(f"El valor objetivo es de: {modelo.ObjVal}")


# Se crea una lista para almacenar los valores de todas las variables de decisión
valores_variables = []
# Se crean diccionarios para almacenar los valores de las variables de decisión de interés
# para su posterior análisis
variables_x = {}
variables_b = {}
variables_n = {}
variables_r = {}
variables_a = {}
variables_u = {}

# Recorrer las variables y obtener el nombre y valor
for var in modelo.getVars():
    nombre_var = var.varName
    valor_var = var.X
    valores_variables.append([nombre_var, valor_var])

# Se obtienen resumenes de las variables de interés y sus valores resultantes
for t in T:
    for j in J:
        if j not in variables_x:
            variables_x[j] = {}
        variables_x[j][t] = X[j, t].X

for t in T:
    suma_atendidos = 0
    for r in R:
        for p in P:
            if B[p, r, t].X == 1:
                suma_atendidos += 1
    variables_b[t] = suma_atendidos


for t in T:
    variables_a[t] = 0
    for r in R:
        for p in P:
            cant_recursos_atendidos = 0
            for j in J:
                cant_recursos_atendidos += A[p, j, r, t].X
            if cant_recursos_atendidos == len(J):
                variables_a[t] += 1

for r in R:
    variables_n[r] = {}
    for t in T:
        for m in M:
            if m not in variables_n[r]:
                variables_n[r][m] = {}
            variables_n[r][m][t] = N[m, r, t].X

for r in R:
    variables_r[r] = {}
    for t in T:
        # Suma los valores de N[m, r, t] para todos los r
        variables_r[r][t] = sum(Rc[p, r, t].X for p in P)

for t in T:
    for j in J:
        if j not in variables_u:
            variables_u[j] = {}
        variables_u[j][t] = sum(U[j, t, c].X for c in C)

costo_total = 0
for t in T:
    for m in M:
        for r in R:
            costo_total += cs[m] * N[m, r, t].x
    for c in C:
        costo_total += cd * d[c] * Y[c, t].x

# Se crea un DataFrame general con dos columnas: 'Variable' y 'Valor', este incluye los valores
# que toman todas las variables de decisión en el modelo
df = pd.DataFrame(valores_variables, columns=['Variable', 'Valor'])
# Se agrega una fila al final del excel con el valor objetivo
valor_objetivo = modelo.ObjVal
df = df._append({'Variable': 'Valor Objetivo', 'Valor': valor_objetivo}, ignore_index=True)

# Luego se resumen los resultados obtenidos de las variables de interés en DataFrames
df_x = pd.DataFrame(variables_x)
df_b = pd.Series(variables_b)
df_r = pd.DataFrame(variables_r)
df_a = pd.Series(variables_a)
df_u = pd.DataFrame(variables_u)

# Se guardan los resultados en archivos Excel
df_x.to_excel('resultados_X.xlsx', sheet_name='X_jt', engine='openpyxl')
df_b.to_excel('resultados_B.xlsx', sheet_name='B_t', engine='openpyxl')
df_r.to_excel('resultados_R.xlsx', sheet_name='R_rt', engine='openpyxl')
df_a.to_excel('resultados_A.xlsx', sheet_name='A_t', engine='openpyxl')
df.to_excel('resultados.xlsx', index=False, engine='openpyxl')
df_u.to_excel('resultados_U.xlsx', sheet_name='U_jt', engine='openpyxl')
with pd.ExcelWriter("resultados_N.xlsx", engine="openpyxl") as writer:
    for nombre_centro, data_dict in variables_n.items():
        # Se convierte el diccionario en un DataFrame
        df = pd.DataFrame(data_dict)
        df.to_excel(writer, sheet_name=f"{nombre_centro}")

# Se exporta el costo total a un archivo de texto
with open("costo_total.txt", "w") as file:
    file.write(f"Costo total: {costo_total}")

print("Datos exportados exitosamente.")
