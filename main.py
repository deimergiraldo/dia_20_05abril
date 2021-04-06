#Ejercicio 1 sobre Fechas y Horas
import openpyxl
import datetime
import datetime as dt

fechas=openpyxl.load_workbook('alumnos.xlsx')

inicio=fechas.active

mi_fecha = datetime.datetime(2016, 10, 1, 23,50, 1, 1) 

#En está parte estoy introduciendo la fecha personalizada a la columna C
inicio['C1']='Fecha personalizada'
for i in range(2,5):
  inicio[f'C{i}']=mi_fecha


fecha_ajustada = dt.datetime.utcnow() +dt.timedelta(hours=-5)

inicio['D1']='Fecha Ajustada'
for i in range(2,5):
  inicio[f'D{i}']=fecha_ajustada


Horalocal = dt.datetime.now().time()

inicio['E1']='Fecha Ajustada'
for i in range(2,5):
  inicio[f'E{i}']=Horalocal


fechas.save('Minuevoarchivo.xlsx')


#Ejercicio 2
##En este ejercicio no voy a subir un nuevo archivo sino que haré uso del archivo anterior 

horas=openpyxl.load_workbook('Minuevoarchivo.xlsx')

columna=horas.active

t1 = dt.datetime.strptime('8:43:12', '%H:%M:%S')
t2 = dt.datetime.strptime('12:00:00', '%H:%M:%S')
suma=t1-t2


columna['F1']='Suma de fechas-horas'
for i in range(2,5):
  columna[f'F{i}']=suma


horas.save('Suma_horas.xlsx')





