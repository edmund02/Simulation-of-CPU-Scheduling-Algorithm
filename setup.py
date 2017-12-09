import sys
from cx_Freeze import setup, Executable

setup(
    name = "Scheduling",
    version = "1.0",
    description = "Simple Python program to test different processor scheduling algorithm. Works with Excel File.",
    executables = [Executable("scheduling.py", base='Console')])