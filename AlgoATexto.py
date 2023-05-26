from pandas import to_datetime
from numpy_financial import irr 

class PlataATexto:
    #Muy basado en https://github.com/wzorroman/numbers_to_letters/blob/master/numbers_to_letters.py
    def __init__(self, MONEDA_SINGULAR, MONEDA_PLURAL, SIGLA):
        self.MONEDA_SINGULAR = MONEDA_SINGULAR
        self.MONEDA_PLURAL = MONEDA_PLURAL
        self.SIGLA = SIGLA


    CENTIMOS_SINGULAR = 'CTV'
    CENTIMOS_PLURAL = 'CTVS'

    MAX_NUMERO = 999999999999

    UNIDADES = (
        'CERO',
        'UNO',
        'DOS',
        'TRES',
        'CUATRO',
        'CINCO',
        'SEIS',
        'SIETE',
        'OCHO',
        'NUEVE'
    )

    DECENAS = (
        'DIEZ',
        'ONCE',
        'DOCE',
        'TRECE',
        'CATORCE',
        'QUINCE',
        'DIECISEIS',
        'DIECISIETE',
        'DIECIOCHO',
        'DIECINUEVE'
    )

    DIEZ_DIEZ = (
        'CERO',
        'DIEZ',
        'VEINTE',
        'TREINTA',
        'CUARENTA',
        'CINCUENTA',
        'SESENTA',
        'SETENTA',
        'OCHENTA',
        'NOVENTA'
    )

    CIENTOS = (
        '_',
        'CIENTO',
        'DOSCIENTOS',
        'TRESCIENTOS',
        'CUATROSCIENTOS',
        'QUINIENTOS',
        'SEISCIENTOS',
        'SETECIENTOS',
        'OCHOCIENTOS',
        'NOVECIENTOS'
    )


    def numero_a_letras(self, numero):
        numero_entero = int(numero)
        if numero_entero > self.MAX_NUMERO:
            raise OverflowError('Número demasiado alto')
        if numero_entero < 0:
            negativo_letras = self.numero_a_letras(abs(numero))
            return f"MENOS {negativo_letras}"
        letras_decimal = ''
        parte_decimal = int(round((abs(numero) - abs(numero_entero)) * 100))
        if parte_decimal > 9:
            letras_decimal = f"PUNTO {self.numero_a_letras(parte_decimal)}"
        elif parte_decimal > 0:
            letras_decimal = f"PUNTO CERO {self.numero_a_letras(parte_decimal)}"
        if numero_entero <= 99:
            resultado = self.leer_decenas(numero_entero)
        elif numero_entero <= 999:
            resultado = self.leer_centenas(numero_entero)
        elif numero_entero <= 999999:
            resultado = self.leer_miles(numero_entero)
        elif numero_entero <= 999999999:
            resultado = self.leer_millones(numero_entero)
        else:
            resultado = self.leer_millardos(numero_entero)
        resultado = resultado.replace('UNO MIL', 'UN MIL')
        resultado = resultado.strip()
        resultado = resultado.replace(' _ ', ' ')
        resultado = resultado.replace('  ', ' ')
        if parte_decimal > 0:
            resultado = f"{resultado} {letras_decimal}"
        return resultado


    def numero_a_moneda(self, numero):
        numero_entero = int(numero)
        parte_decimal = int(round((abs(numero) - abs(numero_entero)) * 100))
        centimos = self.CENTIMOS_SINGULAR if parte_decimal == 1 else self.CENTIMOS_PLURAL
        moneda = self.MONEDA_SINGULAR if numero_entero == 1 else self.MONEDA_PLURAL
        letras = self.numero_a_letras(numero_entero)
        letras = letras.replace('UNO', 'UN')
        aux_decimal = self.numero_a_letras(parte_decimal).replace('UNO', 'UN')
        letras_decimal = f"con {aux_decimal} {centimos}"
        letras = f"{letras} {letras_decimal} {moneda}"
        return letras


    def leer_decenas(self, numero):
        if numero < 10:
            return self.UNIDADES[numero]
        decena, unidad = divmod(numero, 10)
        if numero <= 19:
            resultado = self.DECENAS[unidad]
        elif numero <= 29:
            resultado = f"VEINTI{self.UNIDADES[unidad]}"
            if resultado == "VEINTICERO":
                resultado = "VEINTE"
        else:
            resultado = self.DIEZ_DIEZ[decena]
            if unidad > 0:
                resultado = f"{resultado} Y {self.UNIDADES[unidad]}"
        return resultado


    def leer_centenas(self, numero):
        centena, decena = divmod(numero, 100)
        if numero == 0:
            resultado = 'CIEN'
        else:
            resultado = self.CIENTOS[centena]
            if decena > 0:
                decena_letras = self.leer_decenas(decena)
                resultado = f"{resultado} {decena_letras}"
            if resultado == "CIENTO":
                resultado = "CIEN"
        return resultado


    def leer_miles(self,numero):
        millar, centena = divmod(numero, 1000)
        resultado = ''
        if millar == 1:
            resultado = ''
        if (millar >= 2) and (millar <= 9):
            resultado = self.UNIDADES[millar]
        elif (millar >= 10) and (millar <= 99):
            resultado = self.leer_decenas(millar)
        elif (millar >= 100) and (millar <= 999):
            resultado = self.leer_centenas(millar)
        resultado = f"{resultado} MIL"
        if centena > 0:
            centena_letras = self.leer_centenas(centena)
            resultado = f"{resultado} {centena_letras}"
        return resultado.strip()


    def leer_millones(self, numero):
        millon, millar = divmod(numero, 1000000)
        resultado = ''
        if millon == 1:
            resultado = ' UN MILLON '
        if (millon >= 2) and (millon <= 9):
            resultado = self.UNIDADES[millon]
        elif (millon >= 10) and (millon <= 99):
            resultado = self.leer_decenas(millon)
        elif (millon >= 100) and (millon <= 999):
            resultado = self.leer_centenas(millon)
        if millon > 1:
            resultado = f"{resultado} MILLONES"
        if (millar > 0) and (millar <= 999):
            centena_letras = self.leer_centenas(millar)
            resultado = f"{resultado} {centena_letras}"
        elif (millar >= 1000) and (millar <= 999999):
            miles_letras = self.leer_miles(millar)
            resultado = f"{resultado} {miles_letras}"
        return resultado


    def leer_millardos(self, numero):
        millardo, millon = divmod(numero, 1000000)
        miles_letras = self.leer_miles(millardo)
        millones_letras = self.leer_millones(millon)
        return f"{miles_letras} MILLONES {millones_letras}"


    def numero_a_moneda_sunat(self, numero):
        numero_entero = int(numero)
        parte_decimal = int(round((abs(numero) - abs(numero_entero)) * 100))
        moneda = self.MONEDA_SINGULAR if numero_entero == 1 else self.MONEDA_PLURAL

        letras = self.numero_a_letras(numero_entero)
        letras = letras.replace('UNO', 'UN')
        letras = f"{letras} {moneda} CON {parte_decimal}/100"
        return letras
    
    def deletrear(self, monto):
        monto_letras = self.numero_a_moneda_sunat(monto)
        monto_redondeado = round(monto, 2)
        final = f"{monto_letras} ({monto_redondeado} {self.SIGLA})"
        return final
    
    def tasas(self, tasa):
        tasa = round(tasa, 2)
        num_letras = self.numero_a_letras(tasa)
        return f"{num_letras} POR CIENTO ({tasa}%)"
    
    def tasas_float(self, tasa):
        tasa = tasa*100
        texto = self.tasas(tasa)
        return texto

    def enteros(self, entero):
        entero = int(entero)
        num_letras = self.numero_a_letras(entero)
        return f"{num_letras} ({entero})"


class CosasATexto:
    
    def __init__(self):
        self.diccionario_numeros = {
        '0': 'CERO',
        '1': 'UNO',
        '2': 'DOS',
        '3': 'TRES',
        '4': 'CUATRO',
        '5': 'CINCO',
        '6': 'SEIS',
        '7': 'SIETE',
        '8': 'OCHO',
        '9': 'NUEVE'
    }

    def DNI(self, dni):
        resultado = ''
        for caracter in dni:
            if caracter.isdigit():
                resultado += self.diccionario_numeros[caracter] + ', '
            elif caracter == "-":
                resultado += "GUION, "
            else:
                resultado += 'LETRA ' + caracter + ', '
        resultado += '(' + dni + ')'
        return resultado

    def mayus(self, texto):
        return texto.upper()
    
    
    def fecha(self, fecha):
        ts = to_datetime(str(fecha)) 
        fecha_str = ts.strftime('%d/%m/%Y')

        dic_meses = {
            1: 'ENERO',
            2: 'FEBRERO',
            3: 'MARZO',
            4: 'ABRIL',
            5: 'MAYO',
            6: 'JUNIO',
            7: 'JULIO',
            8: 'AGOSTO',
            9: 'SEPTIEMBRE',
            10: 'OCTUBRE',
            11: 'NOVIEMBRE',
            12: 'DICIEMBRE'
        }
        
        dic_numeros = {
            1: 'UNO',
            2: 'DOS',
            3: 'TRES',
            4: 'CUATRO',
            5: 'CINCO',
            6: 'SEIS',
            7: 'SIETE',
            8: 'OCHO',
            9: 'NUEVE',
            10: 'DIEZ',
            11: 'ONCE',
            12: 'DOCE',
            13: 'TRECE',
            14: 'CATORCE',
            15: 'QUINCE',
            16: 'DIECISÉIS',
            17: 'DIECISIETE',
            18: 'DIECIOCHO',
            19: 'DIECINUEVE',
            20: 'VEINTE',
            21: 'VEINTIUNO',
            22: 'VEINTIDOS',
            23: 'VEINTITRÉS',
            24: 'VEINTICUATRO',
            25: 'VEINTICINCO',
            26: 'VEINTISÉIS',
            27: 'VEINTISIETE',
            28: 'VEINTIOCHO',
            29: 'VEINTINUEVE',
            30: 'TREINTA',
            31: 'TREINTA Y UNO'
        }
        
        dic_anio = {
            2015: 'DOS MIL QUINCE',
            2016: 'DOS MIL DIECISÉIS',
            2017: 'DOS MIL DIECISIETE',
            2018: 'DOS MIL DIECIOCHO',
            2019: 'DOS MIL DIECINUEVE',
            2020: 'DOS MIL VEINTE',
            2021: 'DOS MIL VEINTIUNO',
            2022: 'DOS MIL VEINTIDÓS',
            2023: 'DOS MIL VEINTITRÉS',
            2024: 'DOS MIL VEINTICUATRO',
            2025: 'DOS MIL VEINTICINCO',
            2026: 'DOS MIL VEINTISÉIS',
            2027: 'DOS MIL VEINTISIETE',
            2028: 'DOS MIL VEINTIOCHO',
            2029: 'DOS MIL VEINTINUEVE',
            2030: 'DOS MIL TREINTA',
            2031: 'DOS MIL TREINTA Y UNO',
            2032: 'DOS MIL TREINTA Y DOS',
            2033: 'DOS MIL TREINTA Y TRES',
            2034: 'DOS MIL TREINTA Y CUATRO',
            2035: 'DOS MIL TREINTA Y CINCO',
            2036: 'DOS MIL TREINTA Y SEIS',
            2037: 'DOS MIL TREINTA Y SIETE',
            2038: 'DOS MIL TREINTA Y OCHO',
            2039: 'DOS MIL TREINTA Y NUEVE',
            2040: 'DOS MIL CUARENTA',
            2041: 'DOS MIL CUARENTA Y UNO',
            2042: 'DOS MIL CUARENTA Y DOS',
            2043: 'DOS MIL CUARENTA Y TRES',
            2044: 'DOS MIL CUARENTA Y CUATRO',
            2045: 'DOS MIL CUARENTA Y CINCO',
            2046: 'DOS MIL CUARENTA Y SEIS',
            2047: 'DOS MIL CUARENTA Y SIETE',
            2048: 'DOS MIL CUARENTA Y OCHO',
            2049: 'DOS MIL CUARENTA Y NUEVE',
            2050: 'DOS MIL CINCUENTA'
        }





        dia, mes, anio = fecha_str.split("/")
        dia = int(dia)
        mes = int(mes)
        anio = int(anio)
        
        fecha_palabras = dic_numeros[dia] + " DE " + dic_meses[mes] + " DEL " + dic_anio[anio]

        return fecha_palabras

class Financiero:
    
    def __init__(self):
        pass

    def flujos(self, monto_prestamo, meses_gracia, numero_cuotas, monto_cuotas, comision):
        flujos = [-(monto_prestamo * (1-comision))]
        flujos.extend([0]*int(meses_gracia))
        flujos.extend([monto_cuotas]*int(numero_cuotas))
        return flujos

    def TIR_mensual(self, monto_prestamo, meses_gracia, numero_cuotas, monto_cuotas, comision):
        tir = irr(self.flujos(monto_prestamo, meses_gracia, numero_cuotas, monto_cuotas, comision)) * 100
        return round(tir, 2)

    def TIR_anual(self, monto_prestamo, meses_gracia, numero_cuotas, monto_cuotas, comision):
        tir = irr(self.flujos(monto_prestamo, meses_gracia, numero_cuotas, monto_cuotas, comision))
        tir = (1 + tir)**12
        tir = tir -1
        tir = tir * 100
        return round(tir, 2)

    def tasa_mora(self, monto_prestamo, meses_gracia, numero_cuotas, monto_cuotas, comision):
        tir = irr(self.flujos(monto_prestamo, meses_gracia, numero_cuotas, monto_cuotas, comision))
        tir = (1 + tir)**12
        tir = tir -1
        tir = tir * 100
        tir = tir/4
        return round(tir, 2)

Cordoba = PlataATexto("CORDOBA", "CORDOBAS", "NIO")
Dolar = PlataATexto("DOLAR", "DOLARES", "USD")
NumTasa = PlataATexto("", "", "")
CosaATexto = CosasATexto()
Fin = Financiero()

