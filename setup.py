import sys
from cx_Freeze import setup, Executable


# GUI applications require a different base on Windows (the default is for a
# console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [
    Executable('main.py',icon="038 Ninetales.ico", targetName="compare_data.exe")
]

include_files = [
    "config.ini"
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
