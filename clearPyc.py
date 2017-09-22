#-------------------------------------------------------------------------------
# Name:        ClearPyc
# Purpose:
#
# Author:      Bruce.Zhu
#
# Created:     14/09/2017
# Copyright:   (c) SQA 2017
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import os
import shutil

def walkFolders(folder):
    folderscount=0
    filescount=0
    size=0
    #walk(top,topdown=True,onerror=None)
    for root,dirs,files in os.walk(folder):
        folderscount+=len(dirs)
        filescount+=len(files)
        size+=sum([os.path.getsize(os.path.join(root,name)) for name in files])
    return folderscount,filescount,size

def RemoveFile(path, fileType):
    if os.path.exists(path):
        folderscount,filescount,size = walkFolders(path)
        for parent,dirnames,filenames in os.walk(path):
            for filename in filenames:
                if fileType in filename:
                    try:
                        delFilePath = os.path.join(parent,filename)
                        os.remove(delFilePath)
                    except Exception as e:
                        print(e)

def RemoveFolder(path, folderType):
    if os.path.exists(path):
        folderscount,filescount,size = walkFolders(path)
        for parent,dirnames,filenames in os.walk(path):
            if folderType in parent:
                #print(parent)
                try:
                    shutil.rmtree(parent)
                except Exception as e:
                    print(e)

#RemoveFile('./src', '.pyc')
RemoveFolder('./src', '__pycache__')
shutil.rmtree('./log')
