""" 
#import AE_PyJsx.py
import winreg

#help(winreg)
i=0
while True:
    try:
        # self.aeKey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\Adobe\\After Effects\\16.0")
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Adobe",0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
        assert key != None, "Key = None"
        printTest = winreg.EnumKey(key, i)
        print(printTest)
        i+=1
    except WindowsError:
        break
 """

import os

testPath = os.path.join(os.path.expanduser('~'), "Desktop", "Projects", "Python", "AE_PyJsx", "Scripts", "testWrite.jsx")

with open(testPath, "wb") as f:
    f.write(b"Woops! I have deleted the content!")
    f.close()