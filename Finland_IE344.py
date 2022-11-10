import pandas as pd
import xlwings as xw
import numpy as np

import os
import re

from pathlib import Path
from setup import  get_desktop_path, get_tempate_file, open_filedialog, show_messagebox


def main():
    create_manifest()

def create_manifest():
    """Choose two files, creates df and new name for template file"""

    file_ell = open_filedialog("Select ELL file")

    file_mrn = open_filedialog("Select MRN file", Path(file_ell).parent)

    df_ell = pd.read_excel(file_ell, sheet_name="Manifest")
    df_mrn = pd.read_excel(file_mrn, sheet_name="Finland Customs")
    path_template = get_tempate_file()
    ell_name_end = re.sub(r'\.\w+', '', Path(file_ell).name)
    new_name = os.path.join(get_desktop_path(), ell_name_end)
    
    # merge dataframes, adding 'MRN' and 'Sequence No'
    df_mrn = df_mrn.rename(columns={'Container No': 'Marks & Nos'})
    df_mrn = df_mrn.loc[: ,['Marks & Nos', 'MRN', 'Sequence No']]
    df_merged = df_ell.merge(df_mrn, how='left', on='Marks & Nos')

    mrn, seq = df_merged.pop('MRN'), df_merged.pop('Sequence No')
    df_merged.insert(5, 'MRN', mrn)
    df_merged.insert(6, 'Sequence No', seq)
    df_merged['MRN'].replace("", np.nan, regex=True)

    if df_merged['MRN'].isna().all():
        show_messagebox("No match")
        exit()

    # groupby mlo
    mlo_group = df_merged.groupby('Rcvr')

    # create file to terminal
    with xw.App(visible=False) as app:
        wb = app.books.open(path_template)
        sht = wb.sheets('Manifest')

        #pasting MRN values
        sht.range('A2').options(pd.DataFrame, index=False, header=False).value = df_merged

        name_terminal_file = new_name + ".xlsx"
        wb.save(name_terminal_file)
        wb.close()

    #create files for MLOs
    for mlo, mlo_data in mlo_group:

        with xw.App(visible=False) as app:
            wb = app.books.open(path_template)
            sht = wb.sheets('Manifest')

            #pasting values corresponding to mlo
            sht.range('A2').options(pd.DataFrame, index=False, header=False).value = mlo_data

            name_mrn_file = new_name + "_" + mlo + ".xlsx"
            wb.save(name_mrn_file)
            wb.close()
    
    show_messagebox("OK")
        
main()