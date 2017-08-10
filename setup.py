# setup file for cx_freeze

import sys
from cx_Freeze import setup, Executable
from os.path import expanduser

#home_directory = expanduser("~")+"\Desktop"

base = None
if sys.platform == "win32":
    base = "Win32GUI"

includes = ["atexit", "re"]
files = ["AddImagesToWordDocument.psm1", "icon.ico", "tkicon.ico", "powershellscript.ps1"]

build_exe_options = {"packages": ["os"], "excludes": ["tkinter"], "includes": includes, "include_files" : files}

setup(
    name = "Screenshot Utility",
    version = "0.1",
    description = "This utility performs screen capturing and file creation with taken screenshots.",
    options = {"build_exe": build_exe_options},
    executables = [Executable("M3ScreenshotUtility.py", base = base, icon="icon.ico",
                              shortcutName="Captor Utility", shortcutDir="DesktopFolder")]
    )


