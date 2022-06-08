import sys
from cx_Freeze import setup, Executable

build_exe_options = {"packages": ["os"], "includes": ["tkinter", "pandas", "webdriver_manager.chrome", "selenium.webdriver.chrome.service",
"datetime", "time", "pathlib"], "include_files": ["dinheiro.ico"]}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="NFP",
    version="1.0",
    description="Lançamento Automatizado de NFP",
    options={"build_exe": build_exe_options},
    executables=[Executable(script="RoboNfs_fim.py", base=base, icon="dinheiro.ico")]
)

#python setup_cx.py build