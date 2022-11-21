import getpass
import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox


def get_desktop_path():
    """
    Function to get the correct desktop path.
    - Checks if English or Swedish spelling.
    - Checks if desktop is synchronized with OneDrive.

    """

    username = getpass.getuser()

    if os.path.exists(f"C:\\Users\\{username}\\Skrivbordet"):
        desktop_path = f"C:\\Users\\{username}\\Skrivbordet"

    elif os.path.exists(f"C:\\Users\\{username}\\Desktop"):
        desktop_path = f"C:\\Users\\{username}\\Desktop"

    elif os.path.exists(f"C:\\Users\\{username}\\OneDrive - BOLLORE\\Skrivbordet"):
        desktop_path = f"C:\\Users\\{username}\\OneDrive - BOLLORE\\Skrivbordet"

    elif os.path.exists(f"C:\\Users\\{username}\\OneDrive - BOLLORE\\Desktop"):
        desktop_path = f"C:\\Users\\{username}\\OneDrive - BOLLORE\\Desktop"

    else:
        show_messagebox("Path to desktop not found.")
        exit()

    return desktop_path


def get_file_name():
    home_path = str(Path.home())
    service_dir = r"\BOLLORE\XPF - Documents\SERVICES"
    paths_joined = os.path.join(home_path + service_dir)

    return paths_joined


def open_filedialog(file_title, path=None):

    service_path = get_file_name() if path is None else path

    root = tk.Tk()
    root.lift()
    root.withdraw()

    filename = filedialog.askopenfilename(
        initialdir=service_path,
        title=file_title,
        filetypes=[("Excel files", ".xls .xlsx")],
    )

    root.quit()

    if filename == "":
        exit()

    return filename


def show_messagebox(type):

    if type == "OK":
        messagebox.showinfo(title="Info", message="The files are now on the desktop")

    if type == "No match":
        messagebox.showwarning(
            title="Info",
            message="The files do not match, please run program again and choose new files.",
        )

    if type == "No path found":
        messagebox.showwarning(title="Info", message="Path to desktop not found.")

    if type == "MRN-column":
        messagebox.showwarning(
            title="MRN-column found already exist",
            message="ELL-file already includes a MRN-column. Please choose a new file.",
        )


def get_template_file():

    username = getpass.getuser()
    path = f"C:\\Users\\{username}\\Documents\\python_templates\\template-mrn.xlsx"

    if not os.path.exists(path):
        path = os.path.join(os.getcwd(), "templates\\template-mrn.xlsx")

    return path


def get_example_files():

    ell_path = os.path.join(os.getcwd(), "example_files\\ell\\ARATL_22032_NLRTM.xls")
    mrn_path = os.path.join(os.getcwd(), "example_files\\mrn\\MRN_NLRTM.xls")

    return ell_path, mrn_path
