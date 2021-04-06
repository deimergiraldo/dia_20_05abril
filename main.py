import openpyxl
import datetime as dt
import time

fechas=openpyxl.load_workbook('alumnos.xlsx')

inicio=fechas.active

dia=time.strftime('%x')


for i in range(1,4):
  inicio[f'C{i}']=dia

fechas.save('Minuevoarchivo.xlsx')













