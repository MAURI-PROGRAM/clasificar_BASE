#importar pandas
import pandas as pd

#importar sqldf para manejar sql
from pandasql import sqldf
print('hola')
df_act=pd.read_excel(open('Sistema de cobro 9-02-2018.xlsx','rb'), sheet_name='Asignacion')


pysqldf = lambda q: sqldf(q, globals())

writer = pd.ExcelWriter('Clasificacion 1-02-2018.xlsx', engine='xlsxwriter')
workbook = writer.book

format1 = workbook.add_format()
format1.set_center_across()

numero_multas=pysqldf("""SELECT IDENTIFICACION_DEUDOR,count(*) as NUM_MULTAS,min(DIAS_MORA) as MIN_MORA	FROM df_act group by IDENTIFICACION_DEUDOR;""")

base_menor=pysqldf("""SELECT df_act.*	FROM df_act 
						INNER JOIN numero_multas ON 
						numero_multas.IDENTIFICACION_DEUDOR = df_act.IDENTIFICACION_DEUDOR 
						where numero_multas.NUM_MULTAS=1 and SALDO<=60;""")
base_menor.to_excel(writer, sheet_name='Menor_Cuantia')
worksheet = writer.sheets['Menor_Cuantia']

base_faseI=pysqldf("""SELECT df_act.*	FROM df_act 
						INNER JOIN numero_multas ON 
						numero_multas.IDENTIFICACION_DEUDOR = df_act.IDENTIFICACION_DEUDOR 
						where not(numero_multas.NUM_MULTAS=1 and SALDO<=60) AND numero_multas.MIN_MORA<=180;""")
base_faseI.to_excel(writer, sheet_name='faseI')
worksheet = writer.sheets['faseI']

base_faseII=pysqldf("""SELECT df_act.*	FROM df_act 
						INNER JOIN numero_multas ON 
						numero_multas.IDENTIFICACION_DEUDOR = df_act.IDENTIFICACION_DEUDOR 
						where not(numero_multas.NUM_MULTAS=1 and SALDO<=60) AND numero_multas.MIN_MORA>180;""")
base_faseII.to_excel(writer, sheet_name='faseII')
worksheet = writer.sheets['faseII']

writer.save()