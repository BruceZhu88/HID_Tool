#-------------------------------------------------------------------------------
# Name:        clearLogs
# Purpose:
#
# Author:      Bruce.Zhu
#
# Created:     04/09/2017
# Copyright:   (c) SQA 2017
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import os
import _thread
import logging

logPath = '.\\log'
maxLogFiles = 150

class clearLogs:

    def __init__(self, messagebox, logfilename):
        self.messagebox = messagebox
        self.logfilename = logfilename

    def walkFolders(self, folder):
        folderscount=0
        filescount=0
        size=0
        #walk(top,topdown=True,onerror=None)
        for root,dirs,files in os.walk(folder):
            folderscount+=len(dirs)
            filescount+=len(files)
            size+=sum([os.path.getsize(os.path.join(root,name)) for name in files])
        return folderscount,filescount,size

    def clearLog(self):
        if os.path.exists(logPath):
            folderscount,filescount,size = self.walkFolders(logPath)
            if filescount > maxLogFiles:
                ask = self.messagebox.askokcancel('Tips', 'Your log files have been exceeded {0}!\nDo you want to clear?'.format(maxLogFiles))
                if ask:
                    for parent,dirnames,filenames in os.walk(logPath):
                        for filename in filenames:
                            if filename not in self.logfilename:
                                try:
                                    delFilePath = os.path.join(parent,filename)
                                    os.remove(delFilePath)
                                except Exception as e:
                                    logging.log(logging.DEBUG, "Delete file {0} failed: {1}".format(delFilePath, e))
        else:
            logging.log(logging.DEBUG, "Object directory {0} does not exist!!".format(logPath))