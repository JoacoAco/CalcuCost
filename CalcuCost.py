import pandas as pd
from openpyxl import Workbook
import openpyxl
from datetime import datetime
from dateutil.relativedelta import relativedelta 
import os
#Datetime y fecha
fecha_hoy = datetime.now()
Año = fecha_hoy.year
Mes = fecha_hoy.month
Dia = fecha_hoy.day
Fecha = str(f" {Mes}-{Año}")

print()
print("Bienvenido a CALCUCOST. El programa que le ayudará a calcular todos los costos mensuales: ")
print(f"Hoy es {Dia}/{Mes}/{Año}")

#Directorio y Ruta Relativa
directorio = str(os.path.dirname(os.path.abspath(__file__)))
nombre_archivo = "\Planilla CalcuCost"+Fecha+".xlsx"
ruta = str(directorio + nombre_archivo)

dfP = pd.DataFrame()
dfS = pd.DataFrame()

if os.path.exists(ruta):
    os.remove(ruta)
    Excel = Workbook()
else:
    Excel = Workbook()

Separador1 = ["------------","PRODUCTOS", "------------"]
Separador2 = ["------------","SERVICIOS","------------"]

P = int(input("1. Productos // 2. Servicios: "))
B = 0
X = 0
CantP = None

while B == 0:
    #PRODUCTOS
    if P == 1:
        CantP = int(input("Ingrese la cantidad de PRODUCTOS que desea incluir: "))
        Categoria = [None] * CantP
        Nombre = [None] * CantP
        Unidad = [None] * CantP
        Cantidad = [None] * CantP
        Precio = [None] * CantP
        Delta_Dias = [None] * CantP
        Precio_str = [None] * CantP
        I = 0
        while I < CantP:
            Categoria[I] = f"Producto {I + 1}"
            Nombre[I] = str(input(f"Nombre del {I + 1}° producto: "))
            Unidad[I] = str(input("Unidad de medida del material (Lts, KG, etc): "))
            UnidadP = Unidad[I]
            Cant = float(input(f"Cantidad de Stock ({UnidadP}): "))
            Cantidad[I] = f"{Cant} {Unidad[I]}"
            Precio[I] = float(input(f"Ingrese el COSTO TOTAL de {Nombre[I]}: "))
            Vencimiento_str = input("Ingrese la fecha de vencimiento (AAAA-MM-DD): ")
            Vencimiento = datetime.strptime(Vencimiento_str, "%Y-%m-%d")
            AñoV = Vencimiento.year
            MesV = Vencimiento.month
            DiaV = Vencimiento.day
            Precio_str[I] = str(f"${Precio[I]}")
            diferencia = relativedelta(Vencimiento, fecha_hoy)
            Delta_Dias[I] = f"Vence en {diferencia.years * 365 + diferencia.months * 30 + diferencia.days + 1} días"
            I = I + 1
        Lista1 = [Separador1 ,Categoria, Nombre, Unidad, Cantidad, Precio_str, Delta_Dias]
        dfP = pd.DataFrame(Lista1)
        P = str(input("¿Desea agregar Servicios a la cuenta? (S/N): "))
        if P == "S":
            P = 2
        elif P == "N":
            B = 1
    #SERVICIOS
    if P == 2:
        X = 1
        B = 0
        J = 0
        I = 0
        while J == 0:
            Tipo_S = int(input("¿Qué tipo de servicio desea agregar (1 o 2)? 1. Básico // 2. de Hogar: "))
            if Tipo_S == 1:
                if 'S_Nombre' not in locals():
                    S_Nombre = [None]
                    S_Precio = [None]
                    S_Unidad = [None]
                    S_Gasto = [None]
                    Costo = [None]
                    Vencimiento_Serv = [None]
                else:
                    S_Nombre.append(None)
                    S_Precio.append(None)
                    S_Unidad.append(None)
                    S_Gasto.append(None)
                    Costo.append(None)
                    Vencimiento_Serv.append(None)

                S_Nombre[I] = str(input("Nombre del Servicio: "))
                S_Unidad[I] = str(input("Unidad en la que se mide: "))
                Unidad = str(S_Unidad[I])
                S_Precio[I] = float(input(f"Ingrese el costo por unidad ({Unidad}): "))
                S_Gasto[I] = float(input("¿Cuántas horas se utiliza el servicio p/ día?: "))
                Consumo = float(input(f"Cuantos {S_Unidad[I]} se utilizan diariamente: "))
                Vencimiento_S = input("Ingrese la fecha de vencimiento del servicio (AAAA-MM-DD): ")
                Vencimiento_S = datetime.strptime(Vencimiento_S, "%Y-%m-%d")
                diferencia_s = relativedelta(Vencimiento_S, fecha_hoy)
                Vencimiento_Serv[I] = f"Vence en {diferencia_s.years * 365 + diferencia_s.months * 30 + diferencia_s.days + 1} días"
                V1 = S_Precio[I]  # Precio p/ unidad
                V2 = S_Gasto[I]
                Cons_Total = V2 * 30 * Consumo
                def Gasto_SB(V1, Cons_Total):
                    return Cons_Total / V1
                Costo[I] = f"Costo: ${Cons_Total * V1}"
                S_Gasto[I] = f"Consumo Total: {Cons_Total} {S_Unidad[I]}'s"
                I = I + 1
                M = str(input("¿Desea agregar OTRO Servicio a la cuenta? (S/N): "))
                if M == "N":
                    Lista2 = [Separador2, S_Nombre, S_Gasto, Costo, Vencimiento_Serv]
                    J = 1
                    B = 1
            elif Tipo_S == 2:
                if 'S_Nombre' not in locals():
                    S_Nombre = [None]
                    S_Precio = [None]
                    S_Unidad = [None]
                    S_Gasto = [None]
                    Costo = [None]
                    Vencimiento_Serv = [None]
                else:
                    S_Nombre.append(None)
                    S_Precio.append(None)
                    S_Unidad.append(None)
                    S_Gasto.append(None)
                    Costo.append(None)
                    Vencimiento_Serv.append(None)

                S_Nombre[I] = str(input("Nombre del servicio: "))
                S_Precio[I] = float(input("Costo del servicio($): "))
                Vencimiento_S = input("Ingrese la fecha de vencimiento del servicio (AAAA-MM-DD): ")
                Vencimiento_S = datetime.strptime(Vencimiento_S, "%Y-%m-%d")
                diferencia_s = relativedelta(Vencimiento_S, fecha_hoy)
                Vencimiento_Serv[I] = f"Vence en {diferencia_s.years * 365 + diferencia_s.months * 30 + diferencia_s.days} días"
                S_Gasto[I] = None
                Costo[I] = str(f"${S_Precio[I]}")
                I = I + 1
                M = str(input("¿Desea agregar OTRO Servicio a la cuenta? (S/N): "))
                if M == "N":
                    Lista2 = [Separador2, S_Nombre, S_Gasto, Costo, Vencimiento_Serv]
                    J = 1
                    B = 1
        dfS = pd.DataFrame(Lista2)
if CantP == None:
    print(dfS)
    dfS.to_excel(ruta, sheet_name='Servicios', index=False)
elif CantP != 0 and X == 0:
    print(dfP)
    dfP.to_excel(ruta, sheet_name='Productos', index=False)
else:
    with pd.ExcelWriter(ruta, engine='xlsxwriter') as writer:
        dfP.to_excel(writer, sheet_name='Productos', index=False)
        dfS.to_excel(writer, sheet_name='Servicios', index=False)
    df_final = pd.concat([dfP, dfS], ignore_index=True)
    df_final.to_excel(writer, sheet_name='P y Servicios', index=False)
    print(df_final)
                
