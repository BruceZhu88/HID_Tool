# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        appUpdate
# Purpose:     Update application online
#
# Author:      Bruce.zhu
#
# Created:     03/09/2017
# Copyright:   (c) SQA 2017
# Licence:     <your licence>
#-------------------------------------------------------------------------------
from time import *
import tkinter.ttk as ttk
import tkinter as tk
from tkinter import messagebox
#-----------------------------------
import os
import re
import threading
import subprocess
import sys
import logging
import configparser
import _thread
from .checkUpdates import checkUpdates

class appUpdate:

    def __init__(self, root, messagebox):
        self.messagebox = messagebox
        self.root = root
        conf = configparser.ConfigParser()
        try:
            conf.read('.\\src\\config\\version.ini')
            self.currentVer = conf.get("version", "app")
            self.server = conf.get("Update", "server")
            self.version_config_path =  self.server + conf.get("Update", "version_config_path")
            self.install_path = conf.get("Update", "install_path")
            self.app_path = self.server + conf.get("Update", "app_path")
            self.unzip_path = conf.get("Update", "unzip_path")
            self.app_zip = conf.get("Update", "app_zip")
        except Exception as e:
            logging.log(logging.DEBUG, 'Error: {0}'.format(e))

    def checkForUpdates(self):
        if not os.path.exists('.\\download'):
            os.makedirs('.\\download')
        dest_dir = '.\\download\\downVer.ini'
        checkupdates = checkUpdates()
        if not checkupdates.downLoadFromURL(self.version_config_path, dest_dir):
            self.messagebox.showinfo('Tips', 'Cannot communicate with new version server!\nPlease check your network!')
            return
        downVer = checkupdates.getVer(dest_dir)
        logging.log(logging.DEBUG, 'Starting compare version')
        if checkupdates.compareVer(downVer, self.currentVer):
            ask = self.messagebox.askokcancel('Tips', 'New version %s is detected !\n Do you want to update now?'%downVer)
            if ask:
                self.downloadThread(downVer)
                logging.log(logging.DEBUG, 'Starting download')
        else:
            self.messagebox.showinfo('Tips', 'No new version!')

    def downloadThread(self, downVer):
        try:
            _thread.start_new_thread(self.downloadZip, (downVer,) )
        except:
            logging.log(logging.DEBUG, 'Cannot start power cycle thread!!!')

    def downloadZip(self, downVer):
        newVerPath = '.\\download\\{}'.format(self.app_zip)
        installFile = '.\\download\\{}'.format(self.install_path)
        checkupdates = checkUpdates()
        url = self.app_path+"_v"+downVer+".zip"
        logging.log(logging.DEBUG, 'url = '+url)
        if not checkupdates.downLoadFromURL(url, newVerPath):
            self.messagebox.showinfo('Tips', 'Cannot communicate with new version server!\nPlease check your network!')
            return
        if not checkupdates.downLoadFromURL(self.server + self.install_path, installFile):
            self.messagebox.showinfo('Tips', 'Cannot communicate with new version server!\nPlease check your network!')
            return
        #download process
        checkupdates.unzip_dir(newVerPath, '.\\download\\{}'.format(self.unzip_path))
        ask = self.messagebox.askokcancel('Tips', 'Do you want to install this new version?')
        if ask:
            logging.log(logging.DEBUG, "Starting install")
            self.installThread()
            logging.log(logging.DEBUG, "Close UI")
            self.root.destroy()
            logging.log(logging.DEBUG, "System exit")
            sys.exit()

    def installThread(self):
        batPath = r'"%s\\download\\%s"'%(os.getcwd(), self.install_path) #Note: path must be '"D:\Program Files"' to avoid include space in path
        logging.log(logging.DEBUG, "Run %s"%batPath)
        try:
            _thread.start_new_thread(self.execBat, (batPath,) )
        except Exception as e:
           logging.log(logging.DEBUG, 'Error when install: {0}'.format(e))

    def execBat(self, path):
        os.system(path)
        #subprocess.Popen(Path, shell=True, stdout=subprocess.PIPE)
