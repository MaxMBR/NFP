import sys
from cx_Freeze import setup, Executable

build_exe_options = {"packages": ["os"], "includes": ["tkinter", "pandas", "webdriver_manager.chrome", "selenium.webdriver.chrome.service",
"datetime", "time", "pathlib"], "include_files": ["dinheiro.ico"]}

#Ativar abaixo para não mostrar o console CMD.
#base = None
#if sys.platform == "win32":
#    base = "Win32GUI"

setup(
    name="NFP",
    version="1.0",
    description="Lançamento Automatizado de NFP",
    options={"build_exe": build_exe_options},
    #executables=[Executable(script="RoboNfs_fim.py", base=base, icon="dinheiro.ico")]
    executables=[Executable(script="RoboNfs_fim.py", icon="dinheiro.ico")]
)



#CX_FREEZE
#python setup_cx.py build

#PYINSTALLER
#pyinstaller --onefile RoboNfs_fim.py 
#pyinstaller --onefile RoboNfs_fim.py -w