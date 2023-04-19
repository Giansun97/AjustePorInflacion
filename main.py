import pandas as pd
import numpy as np
from functions import calcular_coeficiente, unir_dataframes, limpiar_df_gastos
import openpyxl
import xlrd


def calcular_ajuste(excel_indices: str, excel_gastos: str) -> None:
    # leer excels con la informacion y realizar limpieza de los datos
    df_indices = (
        pd.read_excel(excel_indices)
            .pipe(calcular_coeficiente, fecha_numerador='2021-12-01')
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

    # Exportamos la pivot y el WP a excel
    pivot.to_excel('resumen_recpam.xlsx', )
    df_merge.to_excel('wp_ajuste.xlsx', index=False)


# Definir variables con nombre de los archivos
if __name__ == "__main__":
    excel_indices: str = input('Ingrese el archivo de indices')
    excel_gastos: str = input('Ingrese el archivo de gastos a ajustar')

    calcular_ajuste(excel_indices, excel_gastos)
