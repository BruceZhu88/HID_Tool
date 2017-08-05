# -*- coding: utf-8 -*-

import sys

from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os"], "include_msvcr": True}#, "compressed": True}

base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

executables = [
    Executable('HID.py', base=base)
]

setup(name='spectre',
      version='1.3',
      description='Tymphany BT_Platform HID Tools',
      options = {"build_exe": build_exe_options},
      executables=executables
      )
