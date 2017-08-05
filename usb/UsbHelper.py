from .pywinusb import hid

class PnpHelper(hid.HidPnPWindowMixin):
    def __init__(self, wndId, handler):
        self.__pnpHandler = handler
        hid.HidPnPWindowMixin.__init__(self, wndId)

    def on_hid_pnp(self, hid_event = None):
        if self.__pnpHandler is not None:
            self.__pnpHandler(hid_event)

class UsbHelper:
    def __init__(self):
        self.__devices = []
        self.__activeIndex = -1
        self.__activeReport = None
        self.__activeDeviceInstanceId = ''

        self.__deviceListChangeHandler = None
        self.__activeDeviceChangeHandler = None
        self.__reportRecievedHandler = None
        self.__pnpHelper = None

    def registerDeviceListChangeHandler(self, handler):
        self.__deviceListChangeHandler = handler

    def registerActiveDeviceChangeHandler(self, handler):
        self.__activeDeviceChangeHandler = handler

    def registerReportRecievedHandler(self, handler):
        self.__reportRecievedHandler = handler

    def registerPnpHandler(self, wndId, handler):
        self.__pnpHandler = PnpHelper(wndId, handler)

    def scan(self):
        self.activateDevice(-1)
        self.__devices = hid.HidDeviceFilter(vendor_id = 0x0cd4).get_devices()
        if len(self.__devices) == 0:
            self.__devices = hid.HidDeviceFilter(vendor_id = 0xabcd).get_devices()
#        print('scan HID device')
#        print(self.__devices)

        if self.__deviceListChangeHandler is not None:
            self.__deviceListChangeHandler()

#        print('last: {0}'.format(self.__activeDeviceInstanceId))

        if self.__activeDeviceInstanceId != '':
            for device in self.__devices:
#                print(' current: {0}'.format(device.instance_id))
                if device.instance_id == self.__activeDeviceInstanceId:
                    self.activateDevice(self.__devices.index(device))

        if self.__activeIndex < 0 and len(self.__devices) > 0:
            self.activateDevice(0)

        device = self.getActiveDevice()
        if device is None:
            self.__activeDeviceInstanceId = ''
        else:
            self.__activeDeviceInstanceId = device.instance_id

    def getDevices(self):
        return self.__devices

    def activateDevice(self, index):
        if index >= len(self.__devices):
            return False
        if index < 0:
            index = -1

        if self.__activeIndex == index:
            return False

        if self.__activeIndex >= 0:
            device = self.__devices[self.__activeIndex]
            device.close()
            self.__activeReport = None

        self.__activeIndex = index

        if self.__activeIndex >= 0:
            device = self.__devices[self.__activeIndex]
            device.open()
            device.set_raw_data_handler(lambda data: self.__onReportRecieved(data))
            #changed by Bruce
            try:
                self.__activeReport = device.find_output_reports()[0]
                self.__activeDeviceInstanceId = device.instance_id
            except:
                return False
                #print('You are in duf mode')
                
        if self.__activeDeviceChangeHandler is not None:
            self.__activeDeviceChangeHandler()
        return True

    def getActiveDevice(self):
        if self.__activeIndex < 0 or self.__activeIndex > len(self.__devices):
            return None

        return self.__devices[self.__activeIndex]

    def getActiveDeviceIndex(self):
        if self.__activeIndex < 0 or self.__activeIndex > len(self.__devices):
            return -1
        return self.__activeIndex

    def sendReport(self, report):
        if self.__activeReport is None:
            return None

        buf = list(self.__activeReport.get_raw_data())
        if len(buf) - 1 < len(report):
            return 'command too long.'

        b = bytes(report, encoding = "utf8")
        i = 0
        while i < len(b):
            buf[i + 1] = b[i]
            i = i + 1
#        print("send: {0}".format(buf))
        try:
            self.__activeReport.send(buf)
        except hid.HIDError as e:
            print(e)
            return 'write failed.'

        return None

    def __onReportRecieved(self, data):
#        print("recieve: {0}".format(data))
        i = 0
        while data[i] != 0:
            i = i + 1
        report = bytes(data[1:i]).decode()
        
        if self.__reportRecievedHandler is not None:
            self.__reportRecievedHandler(report)
    
