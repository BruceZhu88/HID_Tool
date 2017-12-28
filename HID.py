import sys
import subprocess
import threading
from src.ui import MainFrame
from src.usb import UsbHelper

if __name__ == '__main__':
        # print('HID app start ...')
    cmd = "tasklist|find /i \"HID.exe\""
    p = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
    tasklist = p.stdout.readlines()
    if len(tasklist) > 1:
        sys.exit()
    usbHelper = UsbHelper()
    mainFrame = MainFrame(usbHelper)

    threads = []
    main = threading.Thread(target=mainFrame.mainLoop())
    threads.append(main)
    threads[0].start()

#    print('HID app exit ...')
