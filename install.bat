rd /s /Q .\ui\
rd /s /Q .\usb\
rd /s /Q .\src\
xcopy .\download\HID\* .\ /s /h /y
::set current_dir=..\
::pushd %current_dir% 
rd /s /Q .\download\HID\ 
del .\download\HID.zip
del .\download\downVer.ini
COLOR 0A
CLS
@ECHO Off
@echo *******************************************************
@echo ***********New version install successfully!***********
@echo *******************************************************
ECHO.
ECHO Press any key to continue . . .
pause>nul
start .\HID.exe