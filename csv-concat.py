import pandas as pd
import glob, os
import sys
from tqdm import tqdm
import time

def csvConcat(folder,output_path):
    path = folder
    output_path = output_path

    extensions = ("*.xls", "*.xlsx")
    file_list = []
    os.chdir(path)

    print('STEP 1 GETTING LIST OF FILES IN FOLDER')
    for files in extensions:
        file_list.extend(glob.glob(files))

    MASTER_DF = pd.DataFrame()
    issue_list = []

    for file in tqdm(file_list):
        time.sleep(1)
        try:
            #set default values
            top_pn = 'empty'
            plant_tab = 'empty'

            #read pdf
            temp_df = pd.read_excel('{}\\{}'.format(path,file), index_col=None)

            if temp_df.columns[2] == 'Unnamed: 2':
                temp_df = pd.read_excel('{}\\{}'.format(path,file), index_col=None, header=1)

            #get tab names
            xl = pd.ExcelFile('{}\\{}'.format(path,file))
            tab = xl.sheet_names[0].split()
            try:
                top_pn = tab[0]
            except:
                continue
            try:
                plant_tab = tab[1]
            except:
                continue

            if top_pn == 'KMAT':
                top_pn = temp_df['Name'][0]
                plant_tab = tab[-1]

            temp_df['TOP LEVEL PN'] = [top_pn]*temp_df.shape[0]
            temp_df['TOP LEVEL PLANT'] = [plant_tab]*temp_df.shape[0]

            temp_df[['Current Plant Status']] = temp_df[['Current Plant Status']].fillna(value=plant_tab)

            MASTER_DF = pd.concat([MASTER_DF,temp_df], sort=True)
        except:
            issue_list.append(file)
            continue

    df_issue = pd.DataFrame()
    df_issue['names'] = issue_list
    df_issue.to_csv('{}\\issue_list.csv'.format(output_path), index=False)

    print('STEP 2 CREATING OUTPUT FILES')

    out_path1 = output_path + '\\output_all.xlsx'
    writer1 = pd.ExcelWriter(out_path1 , engine='xlsxwriter')
    MASTER_DF.to_excel(writer1, sheet_name='Sheet1', index=False)
    writer1.save()

    out_path2 = output_path + '\\output_removed_dups.xlsx'
    writer2 = pd.ExcelWriter(out_path2 , engine='xlsxwriter')
    MASTER_DF.to_excel(writer2, sheet_name='Sheet1', index=False)
    writer2.save()

    print('SUMMARY-------------------------------------------------------------')
    print('Processed {} files.'.format(len(file_list)))
    print('Issues with {} files. See issue list in current directory'.format(len(issue_list)))

    return print('FINISHED PLEASE CHECK OUTPUT FOLDER: {}'.format(output_path))

if __name__ == "__main__":
    folder = sys.argv[1]
    output_path = sys.argv[2]
    csvConcat(folder,output_path)
