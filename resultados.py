
import os
import sys
import comtypes.client

#set the following flag to True to attach to an existing instance of the program
#otherwise a new instance of the program will be started
AttachToInstance = True

#set the following flag to True to manually specify the path to ETABS.exe
#this allows for a connection to a version of ETABS other than the latest installation
#otherwise the latest installed version of ETABS will be launched
SpecifyPath = False

#if the above flag is set to True, specify the path to ETABS below
ProgramPath = r"C:\Program Files\Computers and Structures\ETABS 19\ETABS.exe"

#full path to the model
#set it to the desired path of your model
APIPath = r'C:\Users\joftv\OneDrive\Documentos\MAMPRO\CSI API'
if not os.path.exists(APIPath):
    try:
        os.makedirs(APIPath)
    except OSError:
        pass
ModelPath = APIPath + os.sep + 'API_1-001.edb'

#create API helper object
helper = comtypes.client.CreateObject('ETABSv1.Helper')
helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)

if AttachToInstance:
    #attach to a running instance of ETABS
    try:
        #get the active ETABS object
        myETABSObject = helper.GetObject("CSI.ETABS.API.ETABSObject") 
    except (OSError, comtypes.COMError):
        print("No running instance of the program found or failed to attach.")
        sys.exit(-1)
else:
    if SpecifyPath:
        try:
            #'create an instance of the ETABS object from the specified path
            myETABSObject = helper.CreateObject(ProgramPath)
        except (OSError, comtypes.COMError):
            print("Cannot start a new instance of the program from " + ProgramPath)
            sys.exit(-1)
    else:
        try: 
            #create an instance of the ETABS object from the latest installed ETABS
            myETABSObject = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject") 
        except (OSError, comtypes.COMError):
            print("Cannot start a new instance of the program.")
            sys.exit(-1)

    #start ETABS application
    myETABSObject.ApplicationStart()

#create SapModel object
SapModel = myETABSObject.SapModel



results1 = SapModel.Results.JointDispl("231",1,200,"231", "231", "TH 01", "Time", 20 )

results2 = SapModel.Results.JointDispl("231", 0)


print(results1)

