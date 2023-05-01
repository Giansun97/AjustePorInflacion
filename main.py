import pandas as pd
import numpy as np
from functions import calcular_coeficiente, unir_dataframes, limpiar_df_gastos
import openpyxl
import xlrd
import datetime

excel_indices: str = input('Ingrese la ubicación donde se encuentra el archivo de indices')
excel_gastos: str = input('Ingrese la ubicación donde se encuentra el archivo de gastos/ingresos')
fecha_numerador = datetime.datetime.strptime(input('Ingrese la fecha de cierre de balance (formato: AAAA-MM-DD): '),
                                             '%Y-%m-%d').date().replace(day=1)


def calcular_ajuste(excel_indices: str, excel_gastos: str, fecha_numerador) -> None:
    print('AJUSTE POR INFLACION')
    print('-------------')
    print('Iniciando ajuste ... ')

    # leer excels con la informacion y realizar limpieza de los datos
    df_indices = (
        pd.read_excel(excel_indices)
            .pipe(calcular_coeficiente, fecha_numerador=fecha_numerador)
            .assign(month=lambda x: x['mes'].dt.month)
            .drop(columns=['mes'])
    )

    df_gastos = (
        pd.read_excel(excel_gastos, skiprows=1)
            .pipe(limpiar_df_gastos)
            .assign(
            Importe=lambda x: x['Debe'] - x['Haber'],
            Fecha=lambda x: pd.to_datetime(x['Fecha'], dayfirst=True),
            month=lambda x: x['Fecha'].dt.month
        )
    )

    # Unir data frames y creamos el work paper ordenado por cuenta
    df_merge = unir_dataframes(df_gastos, df_indices).sort_values(by='Cuenta')

    df_merge = df_merge.drop(columns=['Identificador', 'Número', 'Descripción  (concepto)', 'Detalle del pase',
                                      'Saldo', 'Unidades', 'Saldo Unidades'])

    # Creamos la pivot table resumen
    pivot = pd.pivot_table(df_merge, index=['Cuenta'], values='recpam', aggfunc=np.sum)

    print('Guardando en excel los resultados ... ')
    # Exportamos la pivot y el WP a excel
    pivot.to_excel('resumen_recpam.xlsx', )
    df_merge.to_excel('wp_ajuste.xlsx', index=False)


calcular_ajuste(excel_indices, excel_gastos, fecha_numerador)
print('Proceso Finalizado')
