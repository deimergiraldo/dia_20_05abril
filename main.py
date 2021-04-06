import openpyxl
import datetime
import datetime as dt

fechas=openpyxl.load_workbook('alumnos.xlsx')

inicio=fechas.active

#dia=time.strftime('%x')

mi_fecha = datetime.datetime(2016, 10, 1, 23,50, 1, 1) 

inicio['C1']='Fecha personalizada'
for i in range(2,5):
  inicio[f'C{i}']=mi_fecha

fecha_ajustada = dt.datetime.utcnow() +dt.timedelta(hours=-5)

inicio['D1']='Fecha Ajustada'
for i in range(2,5):
  inicio[f'D{i}']=fecha_ajustada

fechas.save('Minuevoarchivo.xlsx')













