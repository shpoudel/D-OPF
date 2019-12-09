import win32com.client
import matplotlib.pyplot as plt
import json
import math
import numpy as np
import networkx as nx
import os

class Data:
    """
    This code is for extracting the data from OpenDSS (Line parameters, Load, Graph, Xfrms). 
    Also, the load in secondary is transferred into primary side if service transformer is not modeled
    shiva.poudel@wsu.edu 2019-10-25

    """
    def __init__(self, filename = ""):
        """
        Inputs:
            filename - string - DSS input file
        Side effects:
            start DSS via COM
        Contains the following DSS COM objects:
            engine
            text
            circuit
        """
        self.engine = win32com.client.Dispatch("OpenDSSEngine.DSS")
        self.engine.Start("0")

        self.text = self.engine.Text
        self.circuit = self.engine.ActiveCircuit
        self.DSSLines = self.circuit.Lines
        self.DSSXfrm = self.circuit.Transformers
        self.DSSLoad = self.circuit.Loads      
        self.DSSGen =  self.circuit.Generators 
        # self.DSSReg =  self.circuit.Regulators 
        if filename != "":
            self.text.Command = "compile " + filename 
        
        # self.text.Command = "solve" 
        # self.text.Command = "Show Voltages LN Nodes"
        self.text.Command = "Show Elements"
        
    # Checking the interface
    def Extract_Data(self, f):

        Linepar = []
        LoadData = []
        Capacitor = []
        Regulator = []
        load_ind = 0
        line_ind = 0    
        reg_ind = 0    
        cap_ind = 0  
        # File starts from 7th line always in Elements.txt file  
        for i in range(7):
            next (f)  
        
        for line in f:
            if line.strip():                               
                row = line.split()
                element = row[0].strip('"')
                attr = element.split('.') 

                # Store distribution lines from the network model             
                if attr[0] == 'Line': 
                    el_name = attr[1]               
                    self.DSSLines.Name = el_name
                    message = dict(Line = self.DSSLines.Name,
                                   index = line_ind,
                                   r = self.DSSLines.Rmatrix,
                                   x = self.DSSLines.Xmatrix,
                                   from_br = row[1],
                                   to_br = row[2])
                    line_ind += 1
                    Linepar.append(message)

                # Store loads from the network model      
                if attr[0] == 'Load': 
                    el_name = attr[1]               
                    self.DSSLoad.Name = el_name
                    message = dict(load = self.DSSLoad.Name,
                                   bus = row[1],
                                   index = load_ind,
                                   kW = self.DSSLoad.kW,
                                   pf = self.DSSLoad.pf,
                                   kVaR = self.DSSLoad.kW* np.tan(math.acos(self.DSSLoad.pf)))
                    load_ind += 1
                    LoadData.append(message)
                
                # Store reactive power generator (capacitor) from the network model               
                if attr[0] == 'Generator': 
                    el_name = attr[1]               
                    self.DSSGen.Name = el_name
                    message = dict(capacitor = self.DSSGen.Name,
                                   bus = row[1],
                                   index = cap_ind,
                                   kVAR = self.DSSGen.Kvar,
                                   num_phase = self.DSSGen.Phases)
                    cap_ind += 1
                    Capacitor.append(message)
                
                # Store regulator from the network model             
                if attr[0] == 'Transformer': 
                    el_name = attr[1]               
                    self.DSSXfrm.Name = el_name
                    message = dict(regulator = self.DSSXfrm.Name,
                                   bus = [row[1], row[2]],
                                   index = reg_ind,
                                   conn = self.DSSXfrm.IsDelta,
                                   kV = self.DSSXfrm.kV,
                                   kVA = self.DSSXfrm.kva)
                    reg_ind += 1
                    Regulator.append(message)
        # print(LoadData)
        # print(Capacitor)
        # print(Regulator)
      
        # Change the directory where you want to save the network model
        path = os.path.normpath(os.getcwd() + os.sep + os.pardir)
        os.chdir(path + '/D-Net')
        with open('LoadData.json', 'w') as fp:
            json.dump(LoadData, fp)

        with open('LineData.json', 'w') as fp:
            json.dump(Linepar, fp)

        with open('CapData.json', 'w') as fp:
            json.dump(Capacitor, fp)

        with open('RegData.json', 'w') as fp:
            json.dump(Regulator, fp)
                    
if __name__ == '__main__':

    # Change the following path to your OpenDSS Master File
    d1 = Data(fr"C:\Users\Auser\Desktop\Shiva\D-OPF\IEEE-123-Bus\Master-IEEE-123.dss")

    # print(os.listdir())
    # for f in os.listdir():
    #     print(f)

    f1 = open("ieee123_Elements.txt","r")
    d1.Extract_Data(f1)
    


