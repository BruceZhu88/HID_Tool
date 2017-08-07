from time import *

from tkinter.filedialog import *
from tkinter.font import *
from tkinter.scrolledtext import *

from .Settings import Settings
import threading
import subprocess
import configparser

class MainFrame:
    def __init__(self,usbHelper):

        self.__tk = Tk()
        self.title = 'HID v1.5.2 SQA'
        self.__settings = Settings('.\settings.joson')
        
        self.mainColor = '#1979ca'
        self.buttonColor = 'deep sky blue'
        self.buttonState = ACTIVE
        self.controlFont = Font(family = 'Consolas', size = 10)
        self.consoleFont = Font(family = 'Consolas', size = 10)
        
        self.getlog = 'init'
        self.COMMAND = ''
        self.checkall = False
        self.changesn = False
        self.dfumode = False
        self.psTool = False
        self.noPrint = False
        self.runTemp = False
        self.vendor_id = ''
        self.ini_filename = ''
        self.update = False
        self.times = 0
        self.totalTimes = 0
        self.High_version_FW = ''
        self.High_version_DSP = ''
        self.High_version_path = ''
        
        self.Low_version_FW = ''
        self.Low_version_DSP = ''
        self.Low_version_path = ''
        self.usbScan = 1  ##Design for dfu mode(On this mode, still can scan device) or pstool mode, represent disconnect, just find one device.
        self.textSN = StringVar()
        
        self.uiStart = 0
        self.button = [0]*13
        self.Btn_SaveLog = None
        self.Btn_Clear = None
        self.Btn_Send = None
        self.__usbHelper = usbHelper
        self.__curActiveDeviceIndex = -1

        self.__firstLogIdx = 0
        self.__lastLogIdx = 0
        self.__lastSavedLogIdx = 0
        self.__autoSave = False
        self.__autoSavePath = None

        self.__listBoxDevice = None
        self.__textLog = None
        self.__autoSaveTimer = None
        
#        self.__textDeviceInfo1 = StringVar()
#        self.__textDeviceInfo2 = StringVar()
#        self.__textDeviceInfo3 = StringVar()
#        self.__textDeviceInfo4 = StringVar()
        self.__textBtnAutoSave = StringVar()
        self.__textCommand = StringVar()
        self.__cacheCommand = []
        self.__cachePos = -1

        self.__updateAutoSave()

        self.__usbHelper.registerDeviceListChangeHandler(lambda : self.__onDeviceListChange())
        self.__usbHelper.registerActiveDeviceChangeHandler(lambda : self.__onActiveDeviceChange())
        self.__usbHelper.registerReportRecievedHandler(lambda log: self.__onLogRecieved(log))
        self.__usbHelper.registerPnpHandler(self.__tk.winfo_id(), lambda event: self.__onDevicePnp(event))
        
        self.send_run_thread()
        self.__settings.loadSettings()

    def mainLoop(self):
        self.__setupContent()
        self.__setupRoot()
        self.uiStart = 1
        self.__onRefreshBtnClick()
        self.__tk.mainloop()
        self.__settings.saveSettings()

    def printText(self,text):
        self.__addLog(text, '->')

    def autoClick(self,command,str,noprint):
        if noprint=='noprint':
            self.noPrint=True
        else:
            self.noPrint=False
        ret = self.__usbHelper.sendReport(command)
        if ret is None:
            pass
            #self.__addLog(command, '>')
        else:
            self.__addLog('Error: ' + ret, '#')
            self.__onRefreshBtnClick()    
        self.COMMAND = str

    def returnRecieved(self):
        return self.getlog
    
    def __setupRoot(self):
        # setup root
        self.__tk.title(self.title)
        self.__tk.resizable(True, True)
        self.__tk.minsize(800, 600)
        self.__tk.configure(bg ='white')

        # center root
        screenWidth = self.__tk.winfo_screenwidth()
        screenHeight = self.__tk.winfo_screenheight()
        rootWidth = 800
        rootHeight = 600
        size = '%dx%d+%d+%d' % (rootWidth, rootHeight, (screenWidth - rootWidth)/2, (screenHeight - rootHeight)/2)
        self.__tk.geometry(size)
        self.__tk.attributes('-topmost', 1)
        #self.__tk.iconbitmap("%s\\it_medieval_blueftp_amain_48px.ico"%os.getcwd())
        self.__tk.iconbitmap(".\it_medieval_blueftp_amain_48px.ico")
            
    def __setupContent(self):
        leftPaneBorder = Frame(self.__tk, bg = self.mainColor)
        leftPane = Frame(leftPaneBorder, bg = 'white')
        leftPane.pack(fill = BOTH, expand = YES, padx = 1, pady = 1)
        self.__setupLeftPane(leftPane)

        rightPane = Frame(self.__tk, bg = 'white')
        self.__setupRightPane(rightPane)

        Frame(self.__tk, height = 10, bg ='white').pack(side = TOP, fill = X, expand = NO)
        Frame(self.__tk, height = 10, bg ='white').pack(side = BOTTOM, fill = X, expand = NO)
        Frame(self.__tk, width = 10, bg ='white').pack(side = LEFT, fill = X, expand = NO)
        leftPaneBorder.pack(side = LEFT, fill = Y)
        Frame(self.__tk, width = 5, bg ='white').pack(side = LEFT, fill = X, expand = NO)
        Frame(self.__tk, width = 10, bg ='white').pack(side = RIGHT, fill = X, expand = NO)
        rightPane.pack(side = RIGHT, fill = BOTH, expand = YES)

    def __setupLeftPane(self, master):
#        topPane = Frame(master, bg = ui.mainColor)
        Button(master, text = 'Rescan', font = self.controlFont, bg = self.mainColor, fg = 'white', anchor = W, relief = FLAT, command = self.__onRefreshBtnClick).pack(side = TOP, fill = X, expand = NO)
#        Label(master, textvariable = self.__textDeviceInfo4, bg = 'white', anchor = W).pack(side = BOTTOM, fill = X, expand = NO)
#        Label(master, textvariable = self.__textDeviceInfo3, bg = 'white', anchor = W).pack(side = BOTTOM, fill = X, expand = NO)
#        Label(master, textvariable = self.__textDeviceInfo2, bg = 'white', anchor = W).pack(side = BOTTOM, fill = X, expand = NO)
#        Label(master, textvariable = self.__textDeviceInfo1, bg = 'white', anchor = W).pack(side = BOTTOM, fill = X, expand = NO)
#        Frame(master, height = 1, bg = ui.mainColor).pack(side = BOTTOM, fill = X, expand = NO)

        self.__listBoxDevice = Listbox(master, font = self.consoleFont, bd = 0, relief = FLAT, highlightthickness = 0, bg = 'white', selectbackground = 'white', selectforeground = 'black', selectmode = SINGLE, width = 30)
        self.__listBoxDevice.bind('<<ListboxSelect>>', self.__onDeviceSelect)
#        topPane.pack(side = TOP, fill = X, expand = NO, ipadx = 2)
        self.__listBoxDevice.pack(side = TOP, fill = BOTH, expand = YES, pady = 5, padx = 2)
        
        Label(master,text= 'HID Command:',font=("Helvetica", 25),fg="blue", anchor = W).pack(side = TOP, fill = X, expand = NO)
        self.button[0] = Button(master, text = 'Check All', font = self.controlFont, bg = self.buttonColor, fg = 'white', anchor = W, relief = FLAT, command = self.__onCheckallBtnClick)
        #self.button[0].configure(state = self.buttonState)
        self.button[0].pack(side = TOP, fill = X, expand = NO)
        
        self.button[1]=Button(master, text= 'Device name', font = self.controlFont, bg = self.buttonColor, fg = 'white', anchor = W, relief = FLAT, command = self.__onDevicenameBtnClick)
        #self.button[1].configure(state = self.buttonState)
        self.button[1].pack(side = TOP, fill = X, expand = NO)
        
        self.button[2]=Button(master, text= 'version', font = self.controlFont, bg = self.buttonColor, fg = 'white', anchor = W, relief = FLAT, command = self.__onVersionBtnClick)
        #self.button[2].configure(state = self.buttonState)
        self.button[2].pack(side = TOP, fill = X, expand = NO)
        
        self.button[3]=Button(master, text= 'MAC address', font = self.controlFont, bg = self.buttonColor, fg = 'white', anchor = W, relief = FLAT, command = self.__onMacaddressBtnClick)
        #self.button[3].configure(state = self.buttonState)
        self.button[3].pack(side = TOP, fill = X, expand = NO)
        
        self.button[4]=Button(master, text= 'Serial Number', font = self.controlFont, bg = self.buttonColor, fg = 'white', anchor = W, relief = FLAT, command = self.__onSerialnumberBtnClick)
        #self.button[4].configure(state = self.buttonState)
        self.button[4].pack(side = TOP, fill = X, expand = NO)
        
        self.button[5]=Button(master, text= 'Battery', font = self.controlFont, bg = self.buttonColor, fg = 'white', anchor = W, relief = FLAT, command = self.__onBatteryBtnClick)
        #self.button[5].configure(state = self.buttonState)
        self.button[5].pack(side = TOP, fill = X, expand = NO)
        
        self.button[6]=Button(master, text= 'Temperature', font = self.controlFont, bg = self.buttonColor, fg = 'white', anchor = W, relief = FLAT, command = self.__onTemperatureBtnClick)
        #self.button[6].configure(state = self.buttonState)
        self.button[6].pack(side = TOP, fill = X, expand = NO)
        
        self.button[7]=Button(master, text= 'power button', font = self.controlFont, bg = self.buttonColor, fg = 'white', anchor = W, relief = FLAT, command = self.__onPowerBtnClick)
        #self.button[7].configure(state = self.buttonState)
        self.button[7].pack(side = TOP, fill = X, expand = NO)
        
        self.button[8]=Button(master, text= 'DFU', font = self.controlFont, bg = self.buttonColor, fg = 'white', anchor = W, relief = FLAT, command = self.__onDFUBtnClick)
        #self.button[8].configure(state = self.buttonState)
        self.button[8].pack(side = TOP, fill = X, expand = NO)
        
        self.button[9]=Button(master, text= 'PsTool Enable', font = self.controlFont, bg = self.buttonColor, fg = 'white', anchor = W, relief = FLAT, command = self.__onPsToolBtnClick)
        #self.button[9].configure(state = self.buttonState)
        self.button[9].pack(side = TOP, fill = X, expand = NO)

        self.button[10]=Button(master, text= 'Auto Update', font = self.controlFont, bg = self.buttonColor, fg = 'white', anchor = W, relief = FLAT, command = self.__onAutoUpdateBtnClick)
        self.button[10].pack(side = TOP, fill = X, expand = NO)
  
        self.button[11]=Button(master, text= 'Change Serial Number', font = self.controlFont, bg = self.buttonColor, fg = 'white', anchor = W, relief = FLAT, command = self.__onChangeSNBtnClick)
        #self.button[11].configure(state = self.buttonState)
        self.button[11].pack(side = TOP, fill = X, expand = NO)
        #self.__textSN = Text(master, font = self.consoleFont, relief = FLAT)
        #self.__textSN.pack(side = BOTTOM, fill = X)
        Label(master, text= 'Input:', fg="blue", anchor = W ).pack(side = LEFT, expand = NO)
        tkinter.Entry(master, textvariable=self.textSN).pack(side=LEFT, fill = X, expand = NO)
        
        self.button[12]=Button(master, text= 'Change', font = self.controlFont, bg = self.mainColor, fg = 'white', anchor = W, relief = FLAT, command = self.__onChangeSNBtnClick)
        #self.button[12].configure(state = self.buttonState)
        self.button[12].pack(side = RIGHT, fill = X, expand = NO)        
        #Label(master,text= 'Auto Run:',font=("Helvetica", 25),fg="blue", anchor = W).pack(side = TOP, fill = X, expand = NO)
        
        #Button(master, text = 'Battery', font = self.controlFont, bg = self.mainColor, fg = 'white', anchor = W, relief = FLAT, command = self.__onRunbattBtnClick).pack(side = TOP, fill = X, expand = NO)
        #Button(master, text = 'Temperature', font = self.controlFont, bg = self.mainColor, fg = 'white', anchor = W, relief = FLAT, command = self.__onRuntempBtnClick).pack(side = TOP, fill = X, expand = NO)

    def __setupRightPane(self, master):
        logPaneBorder = Frame(master, bg = self.mainColor)
        logPane = Frame(logPaneBorder)
        logPane.pack(fill = BOTH, expand = YES, padx = 1, pady = 1)
        self.__setupLogPane(logPane)

        cmdPaneBorder = Frame(master, bg = self.mainColor)
        cmdPane = Frame(cmdPaneBorder)
        cmdPane.pack(fill = X, expand = NO, padx = 1, pady = 1)
        self.__setupCmdPane(cmdPane)

        cmdPaneBorder.pack(side = BOTTOM, fill = X, expand = NO)
        Frame(master, height = 5, bg = 'white').pack(side = BOTTOM, fill = X, expand = NO)
        logPaneBorder.pack(side = RIGHT, fill = BOTH, expand = YES)

    def __setupLogPane(self, master):
        topPane = Frame(master, bg = self.mainColor)
        self.Btn_Clear = Button(topPane, text='Clear', font=self.controlFont, bg=self.mainColor, fg='white', anchor=W, relief=FLAT, command=self.__onClearBtnClick)
        self.Btn_Clear.pack(side=RIGHT, fill=NONE, expand=NO, padx = 5)
        self.Btn_SaveLog = Button(topPane, textvariable = self.__textBtnAutoSave, font = self.controlFont, bg = self.mainColor, fg = 'white', anchor = W, relief = FLAT, command = self.__onAutoSaveBtnClick)
        self.Btn_SaveLog.pack(side = LEFT, fill = X, expand = YES)
#        Label(topPane, text = '', bg = ui.mainColor, fg = 'white', anchor = W).pack(side = LEFT, fill = X, expand = NO)
        topPane.pack(side=TOP, fill=X, expand=NO)
        self.__textLog = ScrolledText(master, font = self.consoleFont, relief = FLAT, spacing1 = 5, spacing2 = 5)
        self.__textLog.pack(side = BOTTOM, fill = BOTH, expand = YES, ipadx = 5, ipady = 5)

#        topPane.pack(side = TOP, fill = X, expand = NO, ipadx = 2)

    def __setupCmdPane(self, master):
        self.Btn_Send = Button(master, text = 'Send', font = self.controlFont, bg = self.mainColor, fg = 'white', relief = FLAT, command = self.__onSendCommand)
        self.Btn_Send.pack(side = RIGHT, fill = Y, expand = NO, ipadx = 5)
        entryCommand = Entry(master, textvariable = self.__textCommand, relief = FLAT)
        entryCommand.bind('<Return>', self.__onSendCommand)
        entryCommand.bind('<Up>', self.__onLastCacheCommand)
        entryCommand.bind('<Down>', self.__onNextCacheCommand)
        entryCommand.pack(fill = BOTH, expand = YES)

    def __onRefreshBtnClick(self):
        self.__usbHelper.scan()
     
    #---------------------------------------------------    
    def __sendCommand(self,command):
        if self.__usbHelper.getActiveDevice() is None:
            return
        if self.usbScan!=0:
            return            
        ret = self.__usbHelper.sendReport(command)
        if ret is None:
            self.__addLog(command, '>')

    def __onVersionBtnClick(self):
        self.__sendCommand('get_version')
        
    def __onPowerBtnClick(self):
        self.__sendCommand('power_button_press')
        
    def __onBatteryBtnClick(self):
        self.__sendCommand('get_batt_cap')
        
    def __onTemperatureBtnClick(self):
        self.__sendCommand('get_batt_temp')
        
    def __onMacaddressBtnClick(self):
        self.__sendCommand('get_bt_mac')  
              
    def __onDevicenameBtnClick(self):
        self.__sendCommand('get_dev_name')
        
    def __onSerialnumberBtnClick(self):
        self.__sendCommand('get_ser_no')        

    def __onDFUBtnClick(self):
        if self.__usbHelper.getActiveDevice() is None:
            return
        if self.usbScan!=0:
            return          
        self.dfumode=True
        
    def __onChangeSNBtnClick(self):
        if self.__usbHelper.getActiveDevice() is None:
            return     
        if self.usbScan!=0:
            return                 
        self.changesn=True
        
    def __onCheckallBtnClick(self):
        if self.__usbHelper.getActiveDevice() is None:
            return
        if self.usbScan!=0:
            return            
        self.checkall=True

    def __onPsToolBtnClick(self):
        if self.__usbHelper.getActiveDevice() is None:
            return
        if self.usbScan!=0:
            return            
        self.psTool=True

    def __onRuntempBtnClick(self):
        if self.__usbHelper.getActiveDevice() is None:
            return         
        if self.usbScan!=0:
            return         
        self.runTemp=True

    def __onAutoUpdateBtnClick(self):
        if self.__usbHelper.getActiveDevice() is None:
            return     
        if self.usbScan!=0:
            return 
        self.ini_filename = askopenfilename(initialdir = 'C:/', title = "Choose config.ini",filetypes = [("ini files","*.ini")])
        if "config.ini" not in self.ini_filename:
            self.update = False
            self.printText("Your selected file maybe is wrong! Please choose again!!")
        else:
            self.times = 0
            self.update=True
   
    def __onAutoSaveBtnClick(self):
        if self.__autoSave:
            self.__saveLog()
            self.__autoSave = False
        else:
            self.__autoSavePath = asksaveasfilename(title = 'Save log to',
                                                    defaultextension = 'log',
                                                    filetypes = [('Log files', '*.log'), ('Text files', '*.txt'), ('All files', '*.*')],
                                                    initialdir = self.__settings.get('log_path'),
                                                    initialfile = 'log' + strftime('%Y%m%d-%H%M%S ', localtime(time())))
            if self.__autoSavePath is not None and self.__autoSavePath != '':
                self.__autoSave = True
                if self.__saveLog():
                    self.printText("Log file is OK!")
                else:
                    self.printText("Maybe you need to restart HID Tool!!!")
                self.__settings.set('log_path', os.path.dirname(self.__autoSavePath))

        self.__updateAutoSave()

    def __onDeviceSelect(self, event):
        if len(self.__listBoxDevice.curselection()) == 0:
            selectedIndex = -1
        else:
            selectedIndex = self.__listBoxDevice.curselection()[0]

        self.__usbHelper.activateDevice(selectedIndex)

    def __onSendCommand(self, event = None):
        if self.__usbHelper.getActiveDevice() is None:
            return

        command = self.__textCommand.get()
        if command is None:
            return
        command = command.lstrip().rstrip()
        if command == '':
            return

        ret = self.__usbHelper.sendReport(command)
        if ret is None:
            self.__addLog(command, '>')
        else:
            self.__addLog('Error: ' + ret, '#')
            self.__onRefreshBtnClick()
        self.__textCommand.set('')

        if self.__cachePos < 0 or command != self.__cacheCommand[self.__cachePos]:
            self.__cacheCommand.insert(0, command)
            self.__cachePos = -1

    def __onLastCacheCommand(self, event = None):
        command = self.__textCommand.get()
        if  self.__cachePos < 0 or command is None or command != self.__cacheCommand[self.__cachePos]:
            self.__cachePos = -1

        if self.__cachePos < len(self.__cacheCommand) - 1:
            self.__cachePos = self.__cachePos + 1

        if self.__cachePos >= 0 and self.__cachePos < len(self.__cacheCommand):
            self.__textCommand.set(self.__cacheCommand[self.__cachePos])

    def __onNextCacheCommand(self, event = None):
        command = self.__textCommand.get()
        if  self.__cachePos < 0 or command is None or command != self.__cacheCommand[self.__cachePos]:
            self.__cachePos = -1
            return

        self.__cachePos = self.__cachePos - 1
        if self.__cachePos >= 0 and self.__cachePos < len(self.__cacheCommand):
            self.__textCommand.set(self.__cacheCommand[self.__cachePos])

    def __onDeviceListChange(self):
        self.usbScan+=1 #Design for duf mode 
        self.__updateDeviceList()
        self.__addLog('Scan: {0} devices found'.format(len(self.__usbHelper.getDevices())), '#')
        #if self.__listBoxDevice.get(0)=='':
            #for i in range(0,len(self.button)): #self.button[i].configure(state = DISABLED)        

    def __onDevicePnp(self, event):
        self.__addLog('Device ' + event, '#')
        self.__usbHelper.scan()

    def __onActiveDeviceChange(self):
        self.__updateActiveDevice()
        device = self.__usbHelper.getActiveDevice()
        if device is not None:
            self.__addLog('Connect to device: {0} [{1:04x}], {2} [{3:04x}]\n'
                          .format(device.product_name, device.product_id, device.vendor_name, device.vendor_id), '#')
            self.vendor_id = "0x"+'{0:04x}'.format(device.vendor_id)
            self.usbScan=0 #Design for duf mode 
            #if self.__listBoxDevice.get(0)!='':
                #for i in range(0,len(self.button)):
                    #self.button[i].configure(state = NORMAL)

    def __onLogRecieved(self, log):
        if self.noPrint==True:
            self.noPrint=False
            self.getlog=log
        else:
            self.__addLog(log, '-> %s'%self.COMMAND)
            self.getlog=log
            self.COMMAND=''

    def __onClearBtnClick(self):
        self.__saveLog()

        self.__textLog.configure(state = NORMAL)

        firstTage = 'log{0}'.format(self.__firstLogIdx)
        firstTageIndex = self.__textLog.tag_ranges(firstTage)
        lastTage = 'log{0}'.format(self.__lastLogIdx - 1)
        lastTageIndex = self.__textLog.tag_ranges(lastTage)

        #if not use try, when text is empty, it will report error
        #Bruce change it
        try:
            self.__textLog.delete(firstTageIndex[0], lastTageIndex[1])
    
            while self.__lastLogIdx > self.__firstLogIdx:
                self.__textLog.tag_delete(firstTage)
                self.__firstLogIdx = self.__firstLogIdx + 1
                firstTage = 'log{0}'.format(self.__firstLogIdx)
        except:
            pass
        self.__textLog.see(END)
        self.__textLog.configure(state = DISABLED)

    def __updateDeviceList(self):
        self.__listBoxDevice.delete(0, self.__listBoxDevice.size())

        for device in self.__usbHelper.getDevices():
            self.__listBoxDevice.insert(END, '  {0}'.format(device.product_name))

        self.__listBoxDevice.insert(END, '')

    def __updateActiveDevice(self):
        if self.__curActiveDeviceIndex == self.__usbHelper.getActiveDeviceIndex():
            return

        if self.__curActiveDeviceIndex >= 0 and self.__curActiveDeviceIndex < self.__listBoxDevice.size():
            self.__listBoxDevice.delete(self.__curActiveDeviceIndex)
            device = self.__usbHelper.getDevices()[self.__curActiveDeviceIndex]
            self.__listBoxDevice.insert(self.__curActiveDeviceIndex, '  {0}'.format(device.product_name))

        self.__curActiveDeviceIndex = self.__usbHelper.getActiveDeviceIndex()

        if self.__curActiveDeviceIndex >= 0 and self.__curActiveDeviceIndex < self.__listBoxDevice.size():
            device = self.__usbHelper.getActiveDevice()
            self.__listBoxDevice.delete(self.__curActiveDeviceIndex)
            self.__listBoxDevice.insert(self.__curActiveDeviceIndex, '> {0}'.format(device.product_name))
            self.__listBoxDevice.selection_set(self.__curActiveDeviceIndex)
            self.__tk.title(self.title+'- ' + device.product_name)

            self.__curActiveDeviceInstanceId = device.instance_id
 #           self.__textDeviceInfo1.set('Product: {0} [0x{1:04x}]'.format(device.product_name, device.product_id))
 #           self.__textDeviceInfo2.set('Vendor: {0} [0x{1:04x}]'.format(device.vendor_name, device.vendor_id))
 #           self.__textDeviceInfo3.set('Version No: {0}'.format(device.version_number))
 #           self.__textDeviceInfo4.set('Serial No: {}'.format(device.serial_number))
        else:
            self.__tk.title(self.title)
            self.__curActiveDeviceInstanceId = ''
 #           self.__textDeviceInfo1.set('no device selected')
 #           self.__textDeviceInfo2.set('')
 #           self.__textDeviceInfo3.set('')
 #           self.__textDeviceInfo4.set('')

    def __updateAutoSave(self):
        if self.__autoSave:
            self.__textBtnAutoSave.set('Auto save on: ' + self.__autoSavePath)
        else:
            self.__textBtnAutoSave.set('Auto save off')

    def __addLog(self, log, tag):
        if len(log) != 0:#avoid value is empty 
            if log[-1] != '\n':
                log = log + '\n'
        else:
            log = 'None\n'
        
        timeStr = strftime('%H:%M:%S ', localtime(time()))

        self.__textLog.configure(state = NORMAL)
        try:
            self.__textLog.insert(END, timeStr + tag + ' ' + log, 'log{0}'.format(self.__lastLogIdx))
        except:
            self.__textLog.insert(END, timeStr + tag + ' ' + 'Sorry!Illegal characters!(e.g Sticker)\n', 'log{0}'.format(self.__lastLogIdx))

        self.__lastLogIdx = self.__lastLogIdx + 1

        if self.__lastLogIdx - self.__firstLogIdx > 1000:
            firstTage = 'log{0}'.format(self.__firstLogIdx)
            firstTageIndex = self.__textLog.tag_ranges(firstTage)
            self.__textLog.delete(firstTageIndex[0], firstTageIndex[1])
            self.__textLog.tag_delete(firstTage)
            self.__firstLogIdx = self.__firstLogIdx + 1

        self.__textLog.see(END)

        if self.__lastLogIdx - self.__lastSavedLogIdx > 100:
            self.__saveLog()

        self.__textLog.configure(state = DISABLED)

    def __saveLog(self):
#        print('auto save {0}, {1}'.format(self.__lastSavedLogIdx, self.__lastLogIdx))
        if self.__autoSaveTimer is not None:
            self.__tk.after_cancel(self.__autoSaveTimer)

        if not self.__autoSave:
            return False

        self.__autoSaveTimer = self.__tk.after(5000, lambda: self.__saveLog())

        if self.__lastSavedLogIdx >= self.__lastLogIdx \
                or self.__autoSavePath is None \
                or self.__autoSavePath == '':
            return False

        file = open(self.__autoSavePath, 'a')
        firstTag = 'log{0}'.format(self.__lastSavedLogIdx)
        firstTagIndex = self.__textLog.tag_ranges(firstTag)
        lastTag = 'log{0}'.format(self.__lastLogIdx - 1)
        lastTagIndex = self.__textLog.tag_ranges(lastTag)
        try:
            content = self.__textLog.get(firstTagIndex[0], lastTagIndex[1])
            file.write(content)
        except Exception as e:
            print(e)
            self.printText("Cannot save your log file! Please try again!!!")
            file.close()
            self.__lastSavedLogIdx = self.__lastLogIdx
            return False
        file.close()
        self.__lastSavedLogIdx = self.__lastLogIdx
        return True

    def send_run_thread(self):
        if self.uiStart == 1:
            while True:
                self.Task_Auto_send()
                """
                if self.button[10]["state"] == DISABLED:
                    self.Btn_SaveLog.configure(state = NORMAL)
                    self.Btn_Clear.configure(state = NORMAL)
                    self.Btn_Send.configure(state = NORMAL)
                    for i in range(0,len(self.button)):
                        self.button[i].configure(state = NORMAL)
                """
        
        self.thread_send_run = threading.Timer(1, self.send_run_thread) #wait 4s to launch, avoid affect main thread(UI)
        self.thread_send_run.setDaemon(True)
        self.thread_send_run.start()

    def Task_Auto_send(self):
        sleep(0.2)
        if self.checkall==True:
            self.printText('--------------------------------------------')
            self.autoClick("get_dev_name","Device: ",'print')
            sleep(0.02)
            self.autoClick("get_mode","State: ",'print')
            sleep(0.02)                
            self.autoClick("get_version","Firmware: ",'print')
            sleep(0.02)
            self.autoClick("get_batt_cap","Battery: ",'print')
            sleep(0.02)
            self.autoClick("get_batt_temp","Temperature: ",'print')
            sleep(0.02)
            self.autoClick("get_bt_mac","MAC address: ",'print')
            sleep(0.02)
            self.autoClick("get_ser_no","Serial number: ",'print')   
            sleep(0.02)
            self.printText('--------------------------------------------')         
            self.checkall=False
        
        if self.dfumode==True:
            if self.vendor_id == "0xabcd":
                self.autoClick('dfu_mode','','print')
                self.printText("dfu_mode command has been sent!")
            elif self.vendor_id == "0x0cd4":  
                self.autoClick('get_mode','','noprint')
                sleep(0.02)
                get_mode=self.returnRecieved()
                if get_mode == 'ON':
                    self.printText('Your sample is on state, so shutting down...')
                    self.__onPowerBtnClick()
                    #self.autoClick('power_button_press','','noprint')
                    sleep(7)
                self.printText('dfu_mode command has been sent!') 
                self.printText('Please check your sample state!') 
                sleep(1)
                self.autoClick('dfu_mode','','noprint')    
                self.printText('Your sample should be in DFU mode!')             
            else:
                self.printText('Cannot identify your device!!')
            self.dfumode = False
        
        if self.psTool == True:
            self.autoClick('get_mode','','noprint')
            sleep(0.02)
            get_mode = self.returnRecieved()
            if get_mode == 'ON':
                self.printText('Your sample is on state, so shutting down...') 
                self.autoClick('power_button_press','','noprint')
                sleep(7)
            self.printText('pstool_enable command has been sent!') 
            self.printText('Please check your sample state!') 
            sleep(1)
            self.autoClick('pstool_enable','','noprint')  
            self.printText('Your sample should be in pstool mode!')             
            self.psTool=False                
            
        if self.changesn==True:
            if self.textSN.get()=='':
                self.printText('Please input serial number!')
            else:
                self.autoClick('get_ser_no','Original Serial number:','noprint')
                sleep(0.02)
                origin_ser_no=self.returnRecieved()                    
                if origin_ser_no!='init':
                    self.printText('--------------------------------------------')  
                    self.printText('Your original serial number is: %s'%origin_ser_no)
                    
                    change_ser_no = self.textSN.get()
                    self.autoClick(("set_ser_no %s"%change_ser_no),'','noprint') 
                    sleep(0.02)
                    self.autoClick('get_ser_no','','noprint')
                    sleep(0.02)
                    
                    changed_ser_no=self.returnRecieved()
                    if changed_ser_no==change_ser_no:
                        self.printText('Your serial number has been changed to: %s'%changed_ser_no)
                    else:
                        self.printText('Failed to change serial number!')
                    
                    self.autoClick(("set_ser_no %s"%origin_ser_no),'','noprint') 
                    sleep(0.02)
                    self.autoClick('get_ser_no','','noprint')
                    sleep(0.02)
                    
                    get_ser_no=self.returnRecieved()
                    if get_ser_no==origin_ser_no:
                        self.printText('Your serial number has been changed back to: %s'%get_ser_no)
                    else:
                        self.printText('Failed to change back to original serial number!')
                    #self.printText('So, your serial number can be changed successfully!')
                    self.printText('--------------------------------------------')  
                else:
                    self.printText('Please click again!') 
            self.changesn=False 
        '''    
        if self.runTemp==True:  
            self.autoClick("get_batt_temp","Temperature: ",'print')
            self.runTemp=False   
        '''
        if self.update == True:
            self.times = 0
            """
            self.Btn_SaveLog.configure(state = DISABLED)
            self.Btn_Clear.configure(state = DISABLED)
            self.Btn_Send.configure(state = DISABLED)
            for i in range(0,len(self.button)):
                self.button[i].configure(state = DISABLED)
            """
            if os.path.isfile(self.ini_filename) == False:
                self.printText("Cannot find file config.ini")
                self.update = False
                return 
            conf = configparser.ConfigParser()
            try:
                conf.read(self.ini_filename)
                self.usbPID = conf.get("USBPID", "usbPID")
                self.High_version_FW = conf.get("Firmware", "High_version_FW")
                self.High_version_DSP = conf.get("Firmware", "High_version_DSP")
                self.High_version_path = conf.get("Firmware", "High_version_path")
                
                self.Low_version_FW = conf.get("Firmware", "Low_version_FW")
                self.Low_version_DSP = conf.get("Firmware", "Low_version_DSP")
                self.Low_version_path = conf.get("Firmware", "Low_version_path")
                self.totalTimes = conf.getint("Upate_Times", "times")
            except Exception as e:
                self.printText("Error when read config.ini : "+e)
                self.update = False                    
                return

            while(self.times < self.totalTimes):
                self.times += 1
                try:
                    if self.High_version_FW == self.Low_version_FW:
                        self.printText('-----------------------Update to %s --- %s times'%(self.High_version_FW, self.times))
                        if not self.__autoUpdate(self.High_version_path, self.High_version_DSP, self.High_version_FW, self.usbPID):
                            self.update = False
                            return
                    else:
                        self.printText('-----------------------Update to %s --- %s times'%(self.High_version_FW, self.times))
                        if not self.__autoUpdate(self.High_version_path, self.High_version_DSP, self.High_version_FW, self.usbPID):
                            self.update = False
                            return
                        self.printText('-----------------------Downgrade to %s --- %s times'%(self.Low_version_FW, self.times))
                        if not self.__autoUpdate(self.Low_version_path, self.Low_version_DSP, self.Low_version_FW, self.usbPID):
                            self.update = False
                            return
                except Exception as e:
                    self.printText('Error when update: '+e)
                    self.update = False
                    return
            
            #need add disable other button
            self.totalTimes = 0
            self.times = 0
            self.update = False
            self.printText("***********************************")
            self.printText("*****Finished all update!**********")
            self.printText("***********************************")

    def __versionCheck(self, version_DSP, version_FW):
        sleep(6)
        #self.__onVersionBtnClick('get_version')
        self.autoClick('get_version','','print')
        sleep(0.8)
        version = self.returnRecieved()
        sleep(0.3)
        if version != '':
            try:
                FW = re.findall(r'FW=(.*),', version)[0]
                DSP = re.findall(r'DSP=(.*)', version)[0]
                if (FW == version_FW) and (DSP == version_DSP):
                    self.printText("Upgrade really succeeded!!")
                    sleep(2)
                else:
                    self.printText("FW="+FW+" version_FW="+version_FW)
                    self.printText("DSP="+DSP+" version_DSP="+version_DSP)
                    self.printText("Upgrade failed !!!!! ---- version numbers are not equal!")
            except Exception as e:
                self.printText("Cannot connect with device")
                self.printText("Error : "+e)
                return False
            return True
        else:
            self.printText("Cannot read version!!")
            self.printText("Update Failed!!!!!")
            return False

    def __autoUpdate(self, firmware_path, version_DSP, version_FW, usbPID):
        dfu_file = firmware_path+'\\firmware.dfu'
        dfu_cmd = firmware_path+'\HidDfuCmd.exe'
        if os.path.isfile(dfu_file) == False or os.path.isfile(dfu_cmd) == False:
            self.printText("Cannot find file %s or %s"%(dfu_file, dfu_cmd))
            self.printText("Please check your related files or path in config.ini!!!")
            return False
        self.printText("Going to DFU mode ...")        
        self.autoClick('dfu_mode','','noprint')
        sleep(7)
        self.printText("Updating ...")
        cmd = r"{0} upgrade {1} {2} 0 0 {3} < .\input.txt".format(dfu_cmd, self.vendor_id, usbPID, dfu_file)
        try:
            s=subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
            pipe=s.stdout.readlines()
        except Exception as e:
            self.printText("Error : "+e)
            return False
        sleep(1.5)  #add this to get enough time
        if "Device reset succeeded\r\n" in  pipe[4].decode('utf-8'):
            self.printText("DFU_Update is finished, going to check version...")
        else:
            self.printText("Failed to upgrade!Going to exit...")
            return False
            #sys.exit(1)
        if s.wait() != 0:
            self.printText("Error on subprocess!!!")
        if not self.__versionCheck(version_DSP, version_FW):
            return False
        return True
