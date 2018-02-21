import sys
from cx_Freeze import setup, Executable
import os

# this was going to be used for tk UI
PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')

# GUI applications require a different base on Windows (the default is for a
# console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [
    Executable('main.py',icon="038 Ninetales.ico", targetName="compare_data.exe"),
    Executable('ui_tk.py',icon="038 Ninetales.ico", targetName="compare_GUI.exe",base=base)
]


include_files = [
    "config.ini",
    #TODO: fix these to use a generic location
    r"C:\Program Files\Python 3.5\DLLs\tcl86t.dll",
    r"C:\Program Files\Python 3.5\DLLs\tk86t.dll",
    r"C:\Program Files\Python 3.5\vcruntime140.dll",

]
includes = [
    "compare_sql_xl",
    "compare_xl",
    "sql_to_xl"

]
packages = [
            "os",
            "argparse",
            "datetime",
            "itertools",
            "openpyxl",
            "pyodbc"
        ]
options = {
    'build_exe': {
        "include_msvcr":True,
        "include_files": include_files,
        "includes" :includes ,

        "packages" : packages
    }
}

setup(  name = "main",
        version = "1.0",
        description = "Compare data from databases and/or excel files and store differences in Excel.",
        options = options,
        executables = executables)
