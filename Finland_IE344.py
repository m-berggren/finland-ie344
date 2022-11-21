"""
Program used to merge data from two chosen excel files, first file (ELL)
has all the data needed for manifestation and second file has
Movement Reference Number (MRN) from Finnish customs tied to containers.
The data is merged and then new files are created;

- One excel file with all information merged together.
- Each Main Line Operator (MLO) will have their own file.

If script is run directly then the two files in 'example_files' will be used
and the result will be saved on the desktop.
"""

import pandas as pd
import xlwings as xw
import numpy as np

import os
import re
from pathlib import Path

from functions import (
    get_desktop_path,
    get_template_file,
    open_filedialog,
    show_messagebox,
    get_example_files,
)


def create_manifest(file_ell=None, file_mrn=None) -> str:

    # If arguments are none, choose files to load data from. First file is ELL, second MRN-file.
    if file_ell is None:
        file_ell = open_filedialog("Select ELL-file")

    if file_mrn is None:
        file_mrn = open_filedialog("Select MRN-file", Path(file_ell).parent)

    # Pandas reads the two excel files into dataframes
    df_ell = pd.read_excel(file_ell, sheet_name="Manifest")
    df_mrn = pd.read_excel(file_mrn, sheet_name="Finland Customs")

    if "MRN" in df_ell.columns:
        show_messagebox("MRN-column")
        exit()

    # Checks if path to excel template file exists, else it will use template file from repository.
    path_template = get_template_file()

    # Extracts the name of the chosen ELL and uses this name for the new file.
    # New file will be saved to desktop.
    ell_name_end = re.sub(r"\.\w+", "", Path(file_ell).name)
    new_name = os.path.join(get_desktop_path(), ell_name_end)

    # Merge dataframes, adding 'MRN' and 'Sequence No'
    df_mrn = df_mrn.rename(
        columns={"Container No": "Marks & Nos"}
    )  # Renames column to another.
    df_mrn = df_mrn.loc[
        :, ["Marks & Nos", "MRN", "Sequence No"]
    ]  # df_mrn now only consists of these 3 columns.
    df_merged = df_ell.merge(
        df_mrn, how="left", on="Marks & Nos"
    )  # merges the 3 columns in df_mrn into df_ell.

    # Changes places of two columns
    mrn, seq = df_merged.pop("MRN"), df_merged.pop("Sequence No")
    df_merged.insert(5, "MRN", mrn)
    df_merged.insert(6, "Sequence No", seq)

    # Replaces empty cells with np.nan to avoid error in below boolean check.
    df_merged["MRN"].replace("", np.nan, regex=True)

    # If no value in MRN column then the chosen excel files do not match or there is no data to be worked on.
    if df_merged["MRN"].isna().all():
        show_messagebox("No match")
        exit()

    # Groupby mlo (each party).
    mlo_group = df_merged.groupby("Rcvr")

    # Creates one file with all information gathered (usually sent to terminal).
    with xw.App(visible=False) as app:
        wb = app.books.open(path_template)
        sht = wb.sheets("Manifest")

        # Pasting MRN values
        sht.range("A2").options(
            pd.DataFrame, index=False, header=False
        ).value = df_merged

        name_terminal_file = new_name + ".xlsx"
        wb.save(name_terminal_file)
        wb.close()

    # Then create files for each mlo (party) with a loop.
    for mlo, mlo_data in mlo_group:

        with xw.App(visible=False) as app:
            wb = app.books.open(path_template)
            sht = wb.sheets("Manifest")

            # Paste values corresponding to mlo.
            sht.range("A2").options(
                pd.DataFrame, index=False, header=False
            ).value = mlo_data

            name_mrn_file = new_name + "_" + mlo + ".xlsx"
            wb.save(name_mrn_file)
            wb.close()

    show_messagebox("OK")


if __name__ == "__main__":
    # Load the files from folder 'example_files' and will save the results on the desktop.
    create_manifest(file_ell=get_example_files()[0], file_mrn=get_example_files()[1])
