import sys
from setuptools import setup, find_packages
from cx_Freeze import setup, Executable
import os.path
PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')


if sys.platform == 'win32':
    base = 'Win32GUI'
base=None

if sys.platform=="win32":
    base="Win32GUI"

executables=[Executable(script="__main__.py",icon='images/vcf.ico',base=base,
              shortcutName="Excel to VCF ",shortcutDir="DesktopFolder",
                        targetName="Excel to VCF.exe",copyright="Copyright (C) 2019 MHA")]

options={"build_exe": {
   "include_files":[os.path.join(PYTHON_INSTALL_DIR, 'DLLs',
    'tk86t.dll'),os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'),"images","images/vcf.ico"],
    "excludes":["matplotlib","scipy","chromedriver","appdirs","packaging","PyQt5","cv2","django","IPython","notebook","pythonwin","wx","PIL","test","numpy","pandas"]}}

            

setup(name="Excel to VCF",options=options,version="1.0.0",
                description="Excel to VCF ",
     long_description="This program used to convert from Excel to csv utf 8 , from CSV utf 8 to VCF  and from Excel to VCF",
      author='E.Mahmoud Hussein Ahmed',
      author_email='m.hussein.esp@gmail.com',
       data_files = [("", ["images/vcf.ico"]
                       )],
      
     
                executables=executables,
      classifiers = [
        "Programming Language :: Python",
        "Programming Language :: Python :: 3",
        "Development Status :: 4 - Beta",
        "Environment :: Other Environment",
        "License ::  MIT LicenseLicense (MIT)",
        "Operating System :: Windows",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Text Processing :: Linguistic"]

                  )


