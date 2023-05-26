#Variables
nombreExcel = "DATOS CONTRATO V1.2.xlsx"
nombreHoja = "DATOS"
hojaFinal = 'DATOS FINALES'

#Imports
print(f"Iniciando: {nombreExcel}")
import pandas as pd 
import AlgoATexto 
import xlwings as xw
import os
import sys

#Cerrar el programa de cara al usuario
def cierre():
    print("Puedes cerrar esta ventana")
    print("Presione cualquier tecla para cerrar...")
    os.system("pause >nul")
    sys.exit(0)

print()
print(f"Asegurece que el archivo '{nombreExcel}' este cerrado")
print(f"No abra el archivo '{nombreExcel}' hasta que el proseso termine")
print()
print("---------------------------------------------------------")
print()


#Leer los datos desde el excel
try:
    print(f"Extrayendo informacion de la hoja '{nombreHoja}' de '{nombreExcel}'")
    dfBase = pd.read_excel(nombreExcel, sheet_name=nombreHoja)
    dfBase = dfBase.dropna()

except FileNotFoundError as e:
    print()
    print("ERROR de archivo:", e)
    print("El programa se detuvo debido a un error de busqueda.")
    print(f"Asegurate el archivo '{nombreExcel}' se encuentre en la misma carpeta que el programa y/o el nombre este correctamente escrito.")
    print()
    cierre()

except ValueError as e:
    print()
    print("ERROR de valor:", e)
    print("El programa se detuvo debido a un error en el valor de una variable.")
    print(f"Asegurate exista la hoja '{nombreHoja}' en el archivo '{nombreExcel}'.")
    print()
    cierre()


# Crear df para guardar la informacion
try:
    print(f"Prosesando la informacion de {nombreExcel}")
    df = pd.DataFrame()

    # Transformar y guardar datos
    df['ID'] = dfBase['ID']

    #Datos Deudora
    df['Nombre May.'] = dfBase['Nombre'].apply(AlgoATexto.CosaATexto.mayus)
    df['Apellidos May.'] = dfBase['Apellidos'].apply(AlgoATexto.CosaATexto.mayus)
    df['Cedula en letras'] = dfBase['Cédula'].apply(AlgoATexto.CosaATexto.DNI)
    df['Estado civil may.'] = dfBase['Estado civil'].apply(AlgoATexto.CosaATexto.mayus)
    df['Direc. Domicilio May.'] = dfBase['Dirección del domicilio'].apply(AlgoATexto.CosaATexto.mayus)
    df['MUNICIPIO'] = dfBase['Municipio'].apply(AlgoATexto.CosaATexto.mayus)
    df['DEPARTAMENTO'] = dfBase['Departamento'].apply(AlgoATexto.CosaATexto.mayus)
    df['Tipo de neg. May.'] = dfBase['Tipo de negocio'].apply(AlgoATexto.CosaATexto.mayus)
    df['Profesión_Deudora'] = dfBase['Profesión Deudora'].apply(AlgoATexto.CosaATexto.mayus)

    #Datos Fiador
    df['Nombre y apellido fiador may'] = dfBase['Nombre y apellidos del Fiador'].apply(AlgoATexto.CosaATexto.mayus)
    df['Estado civil fiador may'] = dfBase['Estado civil Fiador'].apply(AlgoATexto.CosaATexto.mayus)
    df['Profesion fiador may'] = dfBase['Profesión Fiador'].apply(AlgoATexto.CosaATexto.mayus)
    df['Direccion domicilio may'] = dfBase['Dirección domicilio Fiador'].apply(AlgoATexto.CosaATexto.mayus)
    df['cedula fiad letras'] = dfBase['Cédula Fiador'].apply(AlgoATexto.CosaATexto.DNI)

    #Datos Credito
    df['Fecha cheque letras'] = dfBase['Fecha de cheque'].apply(AlgoATexto.CosaATexto.fecha)
    df['Fecha de vencimiento letras'] = dfBase['Fecha de vencimiento'].apply(AlgoATexto.CosaATexto.fecha)
    df['INVERSION LETRAS CORDOBAS'] = dfBase['Inversión Cordobas'].apply(AlgoATexto.Cordoba.deletrear)
    df['Inversion en letras'] = dfBase['Inversión Dolares'].apply(AlgoATexto.Dolar.deletrear)
    df['Cantidad a pagar en letras'] = dfBase['Cantidad total a pagar'].apply(AlgoATexto.Dolar.deletrear)
    df['Cuotas en letras'] = dfBase['Cuotas'].apply(AlgoATexto.Dolar.deletrear)
    df['Tasa_Nominal'] = dfBase['Tasa Nominal'].apply(AlgoATexto.NumTasa.tasas_float)
    df['Comision_Desembolso'] = dfBase['Comision desembolso'].apply(AlgoATexto.NumTasa.tasas_float)

    dfBase['TCEM'] = dfBase.apply(lambda row: AlgoATexto.Fin.TIR_mensual(row['Inversión Dolares'], row['Meses de gracia'], row['Numero Cuotas'], row['Cuotas'],row['Comision desembolso']), axis=1)
    df['TCEM'] = dfBase['TCEM'].apply(AlgoATexto.NumTasa.tasas)

    dfBase['TCEA'] = dfBase.apply(lambda row: AlgoATexto.Fin.TIR_anual(row['Inversión Dolares'], row['Meses de gracia'], row['Numero Cuotas'], row['Cuotas'],row['Comision desembolso']), axis=1)
    df['TCEA'] = dfBase['TCEA'].apply(AlgoATexto.NumTasa.tasas)

    dfBase['Tasa_Mora'] = dfBase.apply(lambda row: AlgoATexto.Fin.tasa_mora(row['Inversión Dolares'], row['Meses de gracia'], row['Numero Cuotas'], row['Cuotas'],row['Comision desembolso']), axis=1)
    df['Tasa_Mora'] = dfBase['Tasa_Mora'].apply(AlgoATexto.NumTasa.tasas)

    df['Meses_de_Gracia'] = dfBase['Meses de gracia'].apply(AlgoATexto.NumTasa.enteros)
    df['Numero_de_Cuotas'] = dfBase['Numero Cuotas'].apply(AlgoATexto.NumTasa.enteros)
    df['Plazo_Meses'] = dfBase.apply(lambda row: AlgoATexto.NumTasa.enteros(row["Meses de gracia"]+row["Numero Cuotas"]), axis=1)
    df['Dia_de_Pago'] = dfBase['Dia de pago'].apply(AlgoATexto.NumTasa.enteros)

except KeyError as e:
    print()
    print("ERROR de llave:", e)
    print("El programa se detuvo debido a un error en el valor de una variable.")
    print(f"Probablemente se modifico el nombre de una columna en la hoja '{nombreHoja}' en el archivo '{nombreExcel}'.")
    print("Revierta el cambio en el nombre de la columna del excel o consiga una version del excel anterior")
    print()
    cierre()


except ValueError as e:
    print()
    print("ERROR de lectura:", e)
    print("El programa se detuvo debido a un error al conectarce al excel.")
    print(f"Probablemente hiso correr el programa sin guardar el archivo '{nombreExcel}' entre las ejecuciones.")
    print(f"Abra el archivo {nombreExcel} y guarde el contenido, luego ejecute el programa")
    print()
    cierre()
    
# Crear un objeto ExcelWriter dirigido al excel existente
try:
    print(f"Guardando '{hojaFinal}' en '{nombreExcel}'")

    # Abrir el libro de Excel, configurando visible a False
    app = xw.App(visible=False)
    book = app.books.open(nombreExcel)

    # Verificar si la hoja ya existe y eliminarla en ese caso
    if hojaFinal in [sheet.name for sheet in book.sheets]:
        book.sheets[hojaFinal].delete()

    # Crear una nueva hoja con el nombre deseado
    sheet = book.sheets.add(hojaFinal)

    # Escribir los datos en la hoja de Excel
    sheet.range('A1').value = df

    # Ajustar el ancho de las columnas
    sheet.autofit('c')

    # Formatear la primera fila (los nombres de las columnas)
    first_row = sheet.range('1:1')
    first_row.api.Font.Size = 12  # Cambiar el tamaño de la fuente
    first_row.api.Font.Bold = True  # Poner en negrita
    first_row.api.Interior.Color = xw.utils.rgb_to_int((173, 216, 230))  # Cambiar el color de fondo

    # Guardar los cambios y cerrar
    book.save()
    app.quit()


except PermissionError as e:
    print()
    print("ERROR de permiso:", e)
    print("El programa se detuvo debido a un error de permiso.")
    print(f"Asegurate que el archivo '{nombreExcel}' no este abierto al momento de ejecutar el programa.")
    print()
    cierre()


print("Todo salio bien :D")
print()
print("---------------------------------------------------------")
print()
print("Siguientes pasos:")
print(f"1) Abrir '{nombreExcel}'")
print(f"2) Revisar que en la hoja '{hojaFinal}' este todo correcto, corregir casillas si es necesario.")
print("3) Guardar el archivo")
print()
cierre()