import pandas as pd
import glob, os
import sys

def csvConcat(folder,output_path):
    path = folder
    output_path = output_path

    file_list = []
    os.chdir(path)

    print('STEP 1 GETTING LIST OF FILES IN FOLDER')
    for file in glob.glob("*.xlsx"):
        file_list.append(file)

    MASTER_DF = pd.DataFrame()

    print('STEP 2 CONCAT ALL FILES')
    for file in file_list:
        #read pdf
        temp_df = pd.read_excel('{}\\{}'.format(path,file), index_col=None)

        #get tab names
        xl = pd.ExcelFile('{}\\{}'.format(path,file))
        tab = xl.sheet_names[0].split()
        top_pn = tab[0]
        plant_tab = tab[1]

        if top_pn == 'KMAT':
            top_pn = temp_df['Name'][0]
            plant_tab = tab[-1]

        temp_df['TOP LEVEL PN'] = [top_pn]*temp_df.shape[0]
        temp_df['TOP LEVEL PLANT'] = [plant_tab]*temp_df.shape[0]

        temp_df[['Current Plant Status']] = temp_df[['Current Plant Status']].fillna(value=plant_tab)

        MASTER_DF = pd.concat([MASTER_DF,temp_df])

    print('STEP 3 CREATING OUTPUT FILES')
    MASTER_DF.to_excel("{}\\output_all.xlsx".format(output_path), index=None)
    # dropping duplicate values
    MASTER_DF_2 = MASTER_DF.drop_duplicates(subset=['Name'], keep='first')
    MASTER_DF_2.to_excel("{}\\output_removed_dups.xlsx".format(output_path), index=None)

    return print('FINISHED PLEASE CHECK OUTPUT FOLDER: {}'.format(output_path))

if __name__ == "__main__":
    folder = sys.argv[1]
    output_path = sys.argv[2]
    csvConcat(folder,output_path)
