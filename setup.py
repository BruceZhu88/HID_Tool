# -*- coding: utf-8 -*-

import sys
import os
import zipfile
import configparser

from cx_Freeze import setup, Executable

conf = configparser.ConfigParser()
try:
    conf.read('./ui/version.ini')
    appVerson = conf.get("version", "app")
except Exception as e:
    print(e)
    sys.exit()
    
# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os"],
                     "include_msvcr": True,
                     "include_files": ['HID_config.ini']
                     }#, "compressed": True}

base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

executables = [
    Executable('HID.py', base=base, icon='./ui/icon/it_medieval_blueftp_amain_48px.ico')
]

setup(name='HID_Tool',
      version=appVerson,
      description='Tymphany SQA BT_Platform HID Tools',
      options = {"build_exe": build_exe_options},
      executables=executables
      )

def zip_dir(dirname, zipfilename):
    filelist = []
    #Check input ...
    fulldirname = os.path.abspath(dirname)
    fullzipfilename = os.path.abspath(zipfilename)
    print("Start to zip {0} to {1}...".format(fulldirname, fullzipfilename))
    if not os.path.exists(fulldirname):
        print("Dir/File %s is not exist, Press any key to quit..."% fulldirname)
        inputStr = input()
        return
    if os.path.isdir(fullzipfilename):
        tmpbasename = os.path.basename(dirname)
        fullzipfilename = os.path.normpath(os.path.join(fullzipfilename, tmpbasename))
    if os.path.exists(fullzipfilename):
        print("%s has already exist, are you sure to modify it ? [Y/N]"% fullzipfilename)
        while 1:
            inputStr = input()
            if inputStr == "N" or inputStr == "n" :
                return
            else:
                if inputStr == "Y" or inputStr == "y" :
                    print("Continue to zip files...")
                    break

    #Get file(s) to zip ...
    if os.path.isfile(dirname):
        filelist.append(dirname)
        dirname = os.path.dirname(dirname)
    else:
        #get all file in directory
        for root, dirlist, files in os.walk(dirname):
            for filename in files:
                filelist.append(os.path.join(root,filename))

    #Start to zip file ...
    destZip = zipfile.ZipFile(fullzipfilename, "w")
    for eachfile in filelist:
        destfile = eachfile[len(dirname):]
        print("Zip file {0}...".format(destfile))
        destZip.write(eachfile, destfile)
    destZip.close()
    print("Zip folder succeed!")

dirname = "./build/exe.win-amd64-3.4/"
zipfilename = "HID_Tool_v{0}.zip".format(appVerson)
zip_dir(dirname, zipfilename)
