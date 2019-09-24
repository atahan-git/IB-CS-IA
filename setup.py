import sys, os
from cx_Freeze import setup, Executable
PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')

__version__ = "1.0.0"

include_files = ["Dictionary.xls", "Output Template.docx",
                os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
                os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'),
                "icon.ico",]
excludes = ["tkinter"]
packages = ["tkinter","os", "datetime", "threading", "xlrd", "docx", "xlutils", "xlwt", "sys"]

setup(
    name = "E-okul to IB Converter",
    description='Converts XLS transcripts from e-okul to the IB format',
    version=__version__,
    options = {"build_exe": {
        'packages': packages,
        'include_files': include_files,
        'include_msvcr': True,

    }},
    executables = [Executable("E-okul to IB Converter.py",base="Win32GUI")],
)