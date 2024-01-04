import pandas as pd

name_archivo = 'Calculos a 150 corridas t cola y nivel de servicio c Proy 2034'
input_folder = 'input_tiempo_espera_servicio'
output_folder = 'output_tiempo_espera_servicio'

def naves_de_carga(value):
    if value == 1:
        return 1
    elif value == 3:
        return 1
    else:
        return None
    
def terminal(value):
    if value == 2:
        return 1
    elif value == 3:
        return 1
    else:
        return None
    
def main(input_folder, output_folder, name_archivo):
    df_cola_v = pd.read_csv(input_folder+'/Tpo_Cola_Nave_V.csv')
    df_cola_v = df_cola_v[['Replica','Numero de Nave', 'Terminal', 'Tipo Nave', 'Tpo Espera en Cola [dias]', 'Tipo_Carga']]
    df_cola_v['Naves de Carga'] = df_cola_v['Tipo Nave'].apply(naves_de_carga)
    df_cola_v['T2'] = df_cola_v['Terminal'].apply(terminal)
    df_cola_v['Puerto'] = 'V'
    df_cola_v['Clave'] = df_cola_v['Puerto'] + '-' + df_cola_v['Replica'].astype(str) + '-' + df_cola_v['Numero de Nave'].astype(str)
    
    df_cola_sa = pd.read_csv(input_folder+'/Tpo_Cola_Nave_SA.csv')
    df_cola_sa = df_cola_sa[['Replica','Numero de Nave', 'Terminal', 'Tipo Nave', 'Tpo Espera en Cola [dias]', 'Tipo_Carga']]
    df_cola_sa['Naves de Carga'] = df_cola_sa['Tipo Nave'].apply(naves_de_carga)
    df_cola_sa['T2'] = df_cola_sa['Terminal'].apply(terminal)
    df_cola_sa['Puerto'] = 'SA'
    df_cola_sa['Clave'] = df_cola_sa['Puerto'] + '-' + df_cola_sa['Replica'].astype(str) + '-' + df_cola_sa['Numero de Nave'].astype(str)
    
    result_df_cola = pd.concat([df_cola_v,df_cola_sa],ignore_index=True)
    
    df_servicio_v = pd.read_csv(input_folder+'/Tpo_Servicio_Nave_V.csv')
    df_servicio_v = df_servicio_v[['Replica','Numero de Nave', 'Terminal', 'Tipo Nave', 'Tpo Atencion Nave [dias]', 'Tipo_Carga']]
    df_servicio_v['Naves de Carga'] = df_servicio_v['Tipo Nave'].apply(naves_de_carga)
    df_servicio_v['Puerto'] = 'V'
    df_servicio_v['T2'] = df_servicio_v['Terminal'].apply(terminal)
    df_servicio_v['Clave'] = df_servicio_v['Puerto'] + '-' + df_servicio_v['Replica'].astype(str) + '-' + df_servicio_v['Numero de Nave'].astype(str)
    column_order = ['Clave'] + [col for col in df_servicio_v.columns if col != 'Clave']
    df_servicio_v = df_servicio_v[column_order]
    
    df_servicio_sa = pd.read_csv(input_folder+'/Tpo_Servicio_Nave_SA.csv')
    df_servicio_sa = df_servicio_sa[['Replica','Numero de Nave', 'Terminal', 'Tipo Nave', 'Tpo Atencion Nave [dias]', 'Tipo_Carga']]
    df_servicio_sa['Naves de Carga'] = df_servicio_sa['Tipo Nave'].apply(naves_de_carga)
    df_servicio_sa['Puerto'] = 'SA'
    df_servicio_sa['T2'] = df_servicio_sa['Terminal'].apply(terminal)
    df_servicio_sa['Clave'] = df_servicio_sa['Puerto'] + '-' + df_servicio_sa['Replica'].astype(str) + '-' + df_servicio_sa['Numero de Nave'].astype(str)
    column_order = ['Clave'] + [col for col in df_servicio_sa.columns if col != 'Clave']
    df_servicio_sa = df_servicio_sa[column_order]
    
    result_df_servicio = pd.concat([df_servicio_v,df_servicio_sa],ignore_index=True)
    
    result_df = result_df_cola.merge(result_df_servicio[['Tpo Atencion Nave [dias]','Clave']], how='left', on='Clave')
    result_df = result_df.rename(columns={'Tpo Atencion Nave [dias]':'Tpo Servicio'})
    result_df['Nivel'] = result_df['Tpo Espera en Cola [dias]']/result_df['Tpo Servicio'] if result_df['Tpo Servicio'].empty == False else None
    
    with pd.ExcelWriter(output_folder + '/' + name_archivo + '_final' + '.xlsx',engine='xlsxwriter') as writter:
            result_df.to_excel(writter, index=False, sheet_name='T Espera')
            result_df_servicio.to_excel(writter, index=False, sheet_name='serv')
            
if __name__ == '__main__':
    main(input_folder,output_folder,name_archivo)