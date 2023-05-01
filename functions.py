import pandas as pd
import openpyxl
import xlrd


def calcular_coeficiente(df, fecha_numerador,
                         columna_fecha='mes', columna_ipc='ipc'):
    """
    Calcula los coeficiente que luego nos van a servir para multiplicarlos por la columna importe

    :param

        df: dataframe de coeficientes.
        fecha_numerador: fecha del último mes del ejercicio. (después la vamos a usar para dividir)
        columna_fecha: el dataframe de de coeficientes debe tener una columna 'mes'
        columna_ipc: el dataframe de de coeficientes debe tener una columna 'ipc'

    :return

        df: data frame de indices con una nueva columna que contiene los coeficientes ya calculados.
    """
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce', format='%y-%m%-d')
    fecha_numerador = pd.to_datetime(fecha_numerador, format='%Y-%m-%d')
    diciembre = df.loc[df[columna_fecha] == fecha_numerador, columna_ipc].iloc[0]
    df['coeficiente'] = diciembre / df[columna_ipc]

    return df


def limpiar_df_gastos(df_gastos):
    """
    Limpieza de datos df_gastos: Esta funcion elimina la columna de Fecha, cambia el nombre de Unnamed 1 a Fecha,
    convierte los valores de la columna Debe y Haber en numeros.

    :param

        df_gastos: toma como input el dataframe de gastos

    :return

        Devuelve el mismo data frame con los datos listos para analizar.
    """

    df_gastos = df_gastos.drop('Fecha', axis=1)
    df_gastos.fillna(0, inplace=True)
    df_gastos['Debe'] = df_gastos['Debe'].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
    df_gastos['Haber'] = df_gastos['Haber'].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
    df_gastos = df_gastos.drop(0)
    df_gastos = df_gastos.rename(columns={'Unnamed: 1': 'Fecha'})

    return df_gastos


def unir_dataframes(df_gastos, df_indices):
    """
    Esta funcion realiza el merge de los Dataframes de gastos e indices, calcula el importe ajustado y el recpam

    :param

        df_gastos: DataFrames de gastos
        df_indices: DataFrames de índices

    :return
        devuelve un nuevo data frame llamado df_merge
    """
    df_merge = pd.merge(df_gastos,
                        df_indices[['coeficiente', 'month']],
                        how='outer')

    df_merge['importe ajustado'] = df_merge['Importe'] * df_merge['coeficiente']

    df_merge_with_nans = df_merge[df_merge.isna().any(axis=1)]

    df_merge = df_merge.dropna()

    df_merge['recpam'] = df_merge['importe ajustado'] - df_merge['Importe']

    return df_merge
