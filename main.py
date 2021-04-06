import openpyxl
import datetime
import time

fechas=openpyxl.load_workbook('alumnos.xlsx')

inicio=fechas.active

#dia=time.strftime('%x')

mi_fecha = datetime.datetime(2016, 10, 1, 23,50, 1, 1) 

for i in range(1,4):
  inicio[f'C{i}']=mi_fecha

fechas.save('Minuevoarchivo.xlsx')













