import os
import sys
from cx_Freeze import setup, Executable

sys.path.append(os.path.abspath('../Utilities'))

target_name = 'Glosas'

with open("../compile_path.txt") as f:
    path = f'{f.readline()}\\{target_name}'

base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

executables = [Executable('main.py', base=base, target_name=target_name,
                          icon='XML-Icon.ico')]

setup(
    name=target_name,
    version='1.0',
    description='Glosas',
    executables=executables,
    options={"build_exe": {
        "packages": ["multiprocessing", "clipboard"],
        "include_files": [('codigosGlosa.csv', 'codigosGlosa.csv')],
        "build_exe": path,
        "includes": "atexit",
        "zip_include_packages": ["*"],
        "zip_exclude_packages": [],
    }},
)
