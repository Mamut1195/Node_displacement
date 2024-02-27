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

#establece las unidades a kgf_mm_C
SapModel.SetPresentUnits(7)

class JointDisplacement:
    def __init__(self, caso:str, nodo:int):
        self.caso = caso
        self.nodo = nodo

    def JointUniqueName(self):
        TableKey = 'Joint Displacements'
        FieldKeyList = ['UniqueName']
        # set the group you want the results for, you can pick either 'All', 'Left Nodes', 'Right Nodes'
        GroupName = 'Top Node'

        TableVersion = 1
        FieldsKeysIncluded = []
        NumberRecords = 1
        TableData = []

        return list(SapModel.DatabaseTables.GetTableforDisplayArray(TableKey, 
                                                                FieldKeyList,
                                                                GroupName, 
                                                                TableVersion, 
                                                                FieldsKeysIncluded, 
                                                                NumberRecords, 
                                                                TableData)[4])
    def OutPutCase(self):
        TableKey = 'Joint Displacements'
        FieldKeyList = ['OutputCase']
        # set the group you want the results for, you can pick either 'All', 'Left Nodes', 'Right Nodes'
        GroupName = 'Top Node'

        TableVersion = 1
        FieldsKeysIncluded = []
        NumberRecords = 1
        TableData = []

        return list(SapModel.DatabaseTables.GetTableforDisplayArray(TableKey, 
                                                                FieldKeyList,
                                                                GroupName, 
                                                                TableVersion, 
                                                                FieldsKeysIncluded, 
                                                                NumberRecords, 
                                                                TableData)[4])
    
    def Tiempo(self):
        TableKey = 'Joint Displacements'
        FieldKeyList = ['StepNumber']
        # set the group you want the results for, you can pick either 'All', 'Left Nodes', 'Right Nodes'
        GroupName = 'Top Node'

        TableVersion = 1
        FieldsKeysIncluded = []
        NumberRecords = 1
        TableData = []

        return list(SapModel.DatabaseTables.GetTableforDisplayArray(TableKey, 
                                                                FieldKeyList,
                                                                GroupName, 
                                                                TableVersion, 
                                                                FieldsKeysIncluded, 
                                                                NumberRecords, 
                                                                TableData)[4])


    def JointDisplacementUx(self):
        TableKey = 'Joint Displacements'
        FieldKeyList = ['Ux']
        # set the group you want the results for, you can pick either 'All', 'Left Nodes', 'Right Nodes'
        GroupName = 'Top Node'

        TableVersion = 1
        FieldsKeysIncluded = []
        NumberRecords = 1
        TableData = []

        return list(SapModel.DatabaseTables.GetTableforDisplayArray(TableKey, 
                                                                FieldKeyList,
                                                                GroupName, 
                                                                TableVersion, 
                                                                FieldsKeysIncluded, 
                                                                NumberRecords, 
                                                                TableData)[4])
    
    def JointDisplacementUy(self):
        TableKey = 'Joint Displacements'
        FieldKeyList = ['Uy']
        # set the group you want the results for, you can pick either 'All', 'Left Nodes', 'Right Nodes'
        GroupName = 'Top Node'

        TableVersion = 1
        FieldsKeysIncluded = []
        NumberRecords = 1
        TableData = []

        return list(SapModel.DatabaseTables.GetTableforDisplayArray(TableKey, 
                                                                FieldKeyList,
                                                                GroupName, 
                                                                TableVersion, 
                                                                FieldsKeysIncluded, 
                                                                NumberRecords, 
                                                                TableData)[4])
    def JointDisplacementUz(self):
        TableKey = 'Joint Displacements'
        FieldKeyList = ['Uz']
        # set the group you want the results for, you can pick either 'All', 'Left Nodes', 'Right Nodes'
        GroupName = 'Top Node'

        TableVersion = 1
        FieldsKeysIncluded = []
        NumberRecords = 1
        TableData = []

        return list(SapModel.DatabaseTables.GetTableforDisplayArray(TableKey, 
                                                                FieldKeyList,
                                                                GroupName, 
                                                                TableVersion, 
                                                                FieldsKeysIncluded, 
                                                                NumberRecords, 
                                                                TableData)[4])
    
    def ListaDesplazamientos(self):
        import pandas as pd

        Ux = self.JointDisplacementUx()
        Uy = self.JointDisplacementUy()
        Nodo = self.JointUniqueName()
        Caso = self.OutPutCase()
        Tiempo = self.Tiempo()

        return pd.DataFrame({'Caso' : Caso,
                             'Nodo' : Nodo,
                             'Tiempo' : Tiempo,
                             'Ux (mm)' : Ux,
                             'Uy (mm)' : Uy })
    
    def SelectNodo(self):
        import pandas as pd
        df = self.ListaDesplazamientos()

        df['Caso'] = df['Caso'].str.strip()
        df['Nodo'] = df['Nodo'].str.strip()
        df['Tiempo'] = df['Tiempo'].str.strip()
        df['Ux (mm)'] = df['Ux (mm)'].str.strip()
        df['Uy (mm)'] = df['Uy (mm)'].str.strip()

        # Convertir la columna a tipo numérico
        df['Ux (mm)'] = pd.to_numeric(df['Ux (mm)'])
        df['Uy (mm)'] = pd.to_numeric(df['Uy (mm)'])
        df['Tiempo'] = pd.to_numeric(df['Tiempo'])
        df['Nodo'] = pd.to_numeric(df['Nodo'])



        df_filtrado_nodo = df[df['Nodo'] == self.nodo]

        return df_filtrado_nodo
    
    def DesplazamientoNodos(self):
        df = self.SelectNodo()

        df_filtrado_caso = df[df['Caso'] == self.caso]

        return df_filtrado_caso
    
    def GraficarDesplazamiento(self):
        import matplotlib.pyplot as plt
        
        df = self.DesplazamientoNodos()
        
        plt.plot(df['Tiempo'], df['Ux (mm)'])
        plt.xlabel('Tiempo (s)')
        plt.ylabel('Ux (mm)')

        return plt.show

    


