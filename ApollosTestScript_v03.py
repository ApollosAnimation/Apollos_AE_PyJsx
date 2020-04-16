import Apollos_AE_PyJsx as AE
import os, time

aeApp = AE.AE_JSInterface(aeVersion = "16.0", returnFolder = os.path.join(os.path.expanduser('~'), "Desktop", "Projects", "Python", "AE_PyJsx", "AE_Projects"))

aeApp.openAE()
#start = time.time()
if aeApp.waitingAELoading():
    #end = time.time()
    #print(end-start)
    # Launch function if AE is ready
    aeApp.jsOpenScene(r"C:/Users/agibson/Desktop/Projects/Python/AE_PyJsx/AE_Projects/PythonTest_v01.aep")
    #time.sleep(3)
    print('AEP opening')
    aeApp.clearJSX()
    aeApp.addDirectCommand('var1 = "Foop"')
    aeApp.addAlert('Added Alert')
    #aeApp.addDirectCommand('var1 = "Foop"')
    #aeApp.jsGetActiveDocument()
    #time.sleep(1)
    #aeApp.jsAlert('Message Recieved')