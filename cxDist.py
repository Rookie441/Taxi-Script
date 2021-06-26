import sys
from cx_Freeze import setup, Executable

build_exe_options = {"packages": ["xlrd","xlsxwriter","time","re","pyxlsb","openpyxl"]}

base = None
if sys.platform == "win32":
    base = "Console"

setup( name = "StreamMobilityScript_v1.3_dist",version = "1.3",description="Rookie441's distributable script",options = {"build_exe": build_exe_options},executables = [Executable("TaxiBooking_Compiled_1.3.py", base=base)])

#change file name

#Commandline#
#python cxDist.py bdist_msi
