
from ui import MainFrame
from usb import UsbHelper
import threading

if __name__ == '__main__':
#    print('spectre app start ...')
    usbHelper = UsbHelper()
    mainFrame = MainFrame(usbHelper)     
    
    threads = []
    main = threading.Thread(target=mainFrame.mainLoop())
    threads.append(main)
    threads[0].start()
    
#    print('spectre app exit ...')
