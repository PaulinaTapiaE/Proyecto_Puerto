import pandas as pd
import os
import glob

input_folder = 'input_nave_y_carga'
xlsx_files = glob.glob(os.path.join(input_folder,'*.xlsx'))
sheet_name = 'Datos'

output_folder = 'output_nave_y_carga'

number = 1.74

def main(input_folder, files, sheet_name, output_folder, number):
    
    for archivo in files:
        name_archivo = archivo.split('.xlsx')[0].split(input_folder)[1]
        df = pd.read_excel(archivo,sheet_name=sheet_name)
        
        df_unpivot = df.copy()
        df_unpivot = df_unpivot[df_unpivot['Name'].isin(['CONT_TEU_T1','CONT_TEU_T3','CONT_TEU_T1_SA','CONT_TEU_T2_SA','CONT_TON_T1','CONT_TON_T1_SA','CONT_TON_T2_SA','CONT_TON_T2','CONT_TON_T3','Cont_NCC_T1_SA','Cont_NCC_T1_V','Cont_NCC_T2_SA','Cont_NCC_T2_V','Cont_NCC_T3_V','Cont_NSC_T1_SA','Cont_NSC_T2_SA','Cont_NSC_T1_V','Cont_NSC_T2_V','Cont_NSC_T3_V','Cont_NAt_NCC_T1_SA','Cont_NAt_NCC_T2_SA','Cont_NAt_NCC_T1_V','Cont_NAt_NCC_T2_V', 'Cont_NAt_NCC_T3_V', 'Cont_NAt_NSC_T1_SA', 'Cont_NAt_NSC_T2_SA', 'Cont_NAt_NSC_T1_V', 'Cont_NAt_NSC_T2_V', 'Cont_NAt_NSC_T3_V'])]        
        unpivot = df_unpivot.pivot(index='Replication', columns='Name', values='RecordedValue').reset_index()
        unpivot['sum_box'] = unpivot['CONT_TEU_T1']+unpivot['CONT_TEU_T3']+unpivot['CONT_TEU_T1_SA']+unpivot['CONT_TEU_T2_SA']
        unpivot['sum_teu'] = unpivot['sum_box']*number
        unpivot['sum_ton'] = unpivot['CONT_TON_T1']+unpivot['CONT_TON_T2']+unpivot['CONT_TON_T3']+unpivot['CONT_TON_T1_SA']+unpivot['CONT_TON_T2_SA']
        unpivot['sum_ton_t2'] = unpivot['CONT_TON_T2']+unpivot['CONT_TON_T3']
        unpivot['ncc_inc'] = unpivot['Cont_NCC_T1_V']+unpivot['Cont_NCC_T2_SA']+unpivot['Cont_NCC_T2_V']+unpivot['Cont_NCC_T1_SA']+unpivot['Cont_NCC_T3_V']
        unpivot['nsc_inc'] = unpivot['Cont_NSC_T1_SA']+unpivot['Cont_NSC_T1_V']+unpivot['Cont_NSC_T2_SA']+unpivot['Cont_NSC_T2_V']+unpivot['Cont_NSC_T3_V']
        unpivot['ncc_at'] = unpivot['Cont_NAt_NCC_T1_SA']+unpivot['Cont_NAt_NCC_T1_V']+unpivot['Cont_NAt_NCC_T2_SA']+unpivot['Cont_NAt_NCC_T2_V']+unpivot['Cont_NAt_NCC_T3_V']
        unpivot['nsc_at'] = unpivot['Cont_NAt_NSC_T1_SA']+unpivot['Cont_NAt_NSC_T1_V']+unpivot['Cont_NAt_NSC_T2_SA']+unpivot['Cont_NAt_NSC_T2_V']+unpivot['Cont_NAt_NSC_T3_V']
        
        unpivot = unpivot[['sum_box','sum_teu','sum_ton','sum_ton_t2','ncc_inc','nsc_inc','ncc_at','nsc_at']]
        
        avg = df.groupby('Name')['RecordedValue'].mean().reset_index()
        avg = avg.rename(columns={'RecordedValue':'avg'})
        avg['avg_mult'] = avg['avg']*number
        
        std = df.groupby('Name')['RecordedValue'].std().reset_index()
        std = std.rename(columns={'RecordedValue':'std'})
        std['std_mult'] = std['std']*number
        
        result = pd.merge(avg,std, on='Name')
        
        with pd.ExcelWriter(output_folder+name_archivo+'_final'+'.xlsx',engine='xlsxwriter') as writter:
            
            result.to_excel(writter, index=False, sheet_name='Hoja 1')
            unpivot.to_excel(writter, index=False, sheet_name='Hoja 2')
        
if __name__ == '__main__':
    main(input_folder,xlsx_files,sheet_name,output_folder,number)