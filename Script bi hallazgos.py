
import pandas as pd
import openpyxl
import os


# %% [markdown]
# Selección carpeta "DESCARGAS SIAPO"  
# Selección carpeta "CONSOLIDADO DE FLOTA"

# %%
folder_path = r"C:\Users\jose.alonso\OneDrive - Grupo Express\Python\BI HALLAZGOS\DESCARGAS SIAPO"
folder_path2 = r"C:\Users\jose.alonso\OneDrive - Grupo Express\Python\BI HALLAZGOS\CONSOLIDADO DE FLOTA"

# %% [markdown]
# Selección archivo ""LISTADO DE INFRACCIONES""
# Selección archivo ¨FLOTA CEXP SIEF"

# %%
df_listado_infracciones = pd.read_excel('C:/Users/jose.alonso/OneDrive - Grupo Express/Python/BI HALLAZGOS/LISTADO DE INFRACCIONES/Listado infracciones ICO (Zonal-Troncal).xlsx',sheet_name='Infracciones')
df_flota = pd.read_excel('C:/Users/jose.alonso/OneDrive - Grupo Express/Python/BI HALLAZGOS/FLOTA CEXP SIEF (TOTAL)/FLOTA CEXP SIEF.xlsx',sheet_name='FLOTA CEXP')

# %%
#Lista de todos los archivos en la carpeta "DESCARGAS SIAPO"

# %%
file_list= [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

# %%
#Lista de todos los archivos en la carpeta "CONSOLIDADO DE FLOTA"

# %%
file_list2= [f for f in os.listdir(folder_path2) if f.endswith('.xlsx')]

# %%
#Cargar archivos en un Dataframe "DESCARGAS SIAPO"

# %%
dataframes=[]

# %%
for file in file_list:
    file_path = os.path.join(folder_path,file)
    df=pd.read_excel(file_path)
    dataframes.append(df)

# %%
combined_df = pd.concat(dataframes,ignore_index=True)

# %%
#Cargar archivos en un Dataframe "CONSOLIDADO DE FLOTA"

# %%
dataframes2=[]

# %%
for file in file_list2:
    file_path2 = os.path.join(folder_path2,file)
    df=pd.read_excel(file_path2)
    dataframes2.append(df)

# %%
combined_df2 = pd.concat(dataframes2,ignore_index=True)

# %%
#Validación de información "DESCARGA HALLAZGOS"

# %%
combined_df.head()

# %%
#Validación de información "CONSOLIDADO DE FLOTA"

# %%
combined_df2.sample(5)

# %%
#Eliminar columnas no necesarias

# %%
combined_df = combined_df.drop(columns=["Columna1","Columna2"])

# %%
combined_df.head()

# %%
#Crear columna calculada CONCAT "DESCARGA HALLAZGOS"

# %%
combined_df['Placa_fecha'] = combined_df['Placa']+combined_df['Fecha Novedad'].dt.strftime('%Y%m%d')

# %%
combined_df.head()

# %%
#Crear columna calculada CONCAT "CONSOLIDADO DE FLOTA"


# %%
combined_df2['Placa_fecha'] = combined_df2['Placa']+combined_df2['Fecha'].dt.strftime('%Y%m%d')

# %%
combined_df2.head()

# %%
#Función de reemplazo en valores de "Linea" de la "Flota"

# %%
def reemplazar_linea(Linea):
    if Linea.startswith('B330R'):
        return 'B330R'
    elif Linea.startswith('NPR'):
        return 'NPR'
    else:
        return Linea

# %%
#Reemplazar columna "Linea" en el Dataframe

# %%
df_flota['Linea'] = df_flota['Linea'].apply(reemplazar_linea)

# %%
# Calcular Marca - Linea en "FLOTA"

# %%
df_flota['Marca - linea'] = df_flota['Marca']+" "+df_flota['Linea']

# %%
df_flota.sample(5)

# %%
#Dejar unicamente información necesaria en el Dataframe "CONSOLIDADO DE FLOTA"

# %%
combined_df2_1 = combined_df2[['Placa_fecha','Centro Operación']]

# %%
#Dejar unicamente información necesaria en el Dataframe "LISTADO DE INFRACCIONES"

# %%
df_listado_infracciones2 = df_listado_infracciones[['CÓDIGO INFRACCIÓN','DESCRIPCIÓN','PUNTAJE','DÍAS DE CORRECCIÓN']]

# %%
#Ajustar nombre de columna para hacer MERGE

# %%
df_listado_infracciones2 = df_listado_infracciones2.rename(columns={'CÓDIGO INFRACCIÓN': 'Tipo Novedad'})

# %%
df_listado_infracciones2.head()

# %%
#Dejar unicamente información necesaria en el Dataframe "FLOTA"

# %%
df_flota2=df_flota[['Placa','Marca - linea']]

# %%
df_flota2.sample(5)

# %%
#Agregar columna calculada "Centro de Operación"

# %%
combined_df = pd.merge(combined_df,combined_df2_1, on='Placa_fecha',how='left')

# %%
combined_df.sample(7)

# %%
#Agregar columnas calculadas "Marca - linea"

# %%
combined_df = pd.merge(combined_df,df_flota2, on='Placa',how='left')

# %%
combined_df.sample(5)

# %%
#Agregar columnas calculadas ""LISTADO DE INFRACCIONES""

# %%
combined_df = pd.merge(combined_df,df_listado_infracciones2, on='Tipo Novedad',how='left')

# %%
combined_df.head()

# %%
combined_df = combined_df.drop(columns=["Placa_fecha"])

# %%
combined_df.head()

# %%
#Calcular estado del hallazgo sengún información de "Ultima_Etapa","Estado_Ultima_Etapa" y "Tiempo_Restante"

# %%
def calcular_estado(Ultima_Etapa, Estado_Ultima_Etapa,Tiempo_Restante):
    if Ultima_Etapa == 'ETAPA 0.0':
        return 'PENDIENTE POR CONTESTAR'
    elif Ultima_Etapa == 'ETAPA 0.1':
        return 'CONTESTADO'
    elif (Ultima_Etapa == 'ETAPA 0.2' and Estado_Ultima_Etapa == 'HALLAZGO CONTESTADO CONTUNDENTE POR CONCESIONARIO EN TIEMPO DE CORRECCION') or (Ultima_Etapa == 'ETAPA 0.3' and Estado_Ultima_Etapa == 'HALLAZGO CONTESTADO CONTUNDENTE POR CONCESIONARIO EN TIEMPO DE CORRECCION'):
        return 'CONTESTADO CONTUNDENTE'
    elif (Ultima_Etapa == 'ETAPA 0.2' and Estado_Ultima_Etapa == 'HALLAZGO CONTESTADO NO CONTUNDENTE POR CONCESIONARIO EN TIEMPO DE CORRECCION') or (Ultima_Etapa == 'ETAPA 0.3' and Estado_Ultima_Etapa == 'HALLAZGO CONTESTADO NO CONTUNDENTE POR CONCESIONARIO EN TIEMPO DE CORRECCION')or (Ultima_Etapa == 'ETAPA 1.0' and Tiempo_Restante == 'Detenido'):
        return 'NO CONTUNDENTE'
    elif (Ultima_Etapa == 'ETAPA 0.4') or (Ultima_Etapa == 'ETAPA 1.0' and Tiempo_Restante == 'Vencido'):
        return 'VENCIDO'
    else:
        return 'ERROR IDENTIFICACIÓN ESTADO HALLAZGO-VALIDAR BASE SUMINISTRADA'

# %%
#Incluir columna estado en el Dataframe

# %%
combined_df['Estado'] = combined_df.apply(lambda row: calcular_valor(row['Ultima Etapa'], row['Estado Ultima Etapa'], row['Tiempo Restante']), axis=1)

# %%
combined_df.head()

# %%
#Validación de información en Blanco

# %%
nan_por_columna2 = combined_df.isna().sum()
print(nan_por_columna2)

# %%
print('NOTA: No es posible tener valores NaN en las columnas Código del Bus, Placa, Centro de operación, Descripción, puntaje, días de corrección y estado')

# %%
#Carpeta para guardar base hallazgos

# %%
save_path= r"C:\Users\jose.alonso\OneDrive - Grupo Express\Python\BI HALLAZGOS\BASE FINAL\archivo_combinado.xlsx"

# %%
combined_df.to_excel(save_path,index=False)

# %%
print("Archivo exportado en: C:/Users/jose.alonso/OneDrive - Grupo Express/Python/BI HALLAZGOS/BASE FINAL/archivo_combinado.xlsx")


