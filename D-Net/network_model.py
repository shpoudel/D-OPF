import win32com.client
import matplotlib.pyplot as plt
import json
import math
import numpy as np
import networkx as nx
import os
import sys
import argparse

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
    def Extract_Data(self, f, f_type):

        if f_type == 'JSON':
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
                        message = dict(name = self.DSSLines.Name,
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
                        message = dict(name = self.DSSLoad.Name,
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

        if f_type == 'TXT':

            line_ind = 0    
            reg_ind = 0    
            cap_ind = 0
            load_ind = 0
            # File starts from 7th line always in Elements.txt file  
            for i in range(7):
                next (f)  
            file_ln = open("LineData.txt","w") 
            file_ln.write("# Line parameters for given distribution network \n \n")
            file_ln.write("Name \t Index \t r \t \t x \t \t from \t to \t \n")

            file_ld = open("LoadData.txt","w") 
            file_ld.write("# Load parameters for given distribution network \n \n")
            file_ld.write("Name \t Bus \t Index \t kW \t \t PF \t \t kVAR \n")

            file_cp = open("CapData.txt","w") 
            file_cp.write("# Reactive power generators for given distribution network \n \n")
            file_cp.write("Name \t Bus \t Index \t kVAR \t \t Phases\n")

            file_rg = open("RegData.txt","w") 
            file_rg.write("# Regulator parameters for given distribution network \n \n")
            file_rg.write("Name \t Bus \t Index \t Conn \t kV \t \t kVA \n")


            for line in f:
                if line.strip():                               
                    row = line.split()
                    element = row[0].strip('"')
                    attr = element.split('.')                                    
                    # Store distribution lines from the network model             
                    if attr[0] == 'Line': 
                        el_name = attr[1]   
                        self.DSSLines.Name = el_name
                        el_name = attr[1] 
                        row[1] = str(row[1])   
                        row[2] = str(row[2])                  
                        file_ln.write("%s \t %d \t %f \t %f \t %s \t %s \n" %(self.DSSLines.Name,line_ind, self.DSSLines.Rmatrix[0], self.DSSLines.Xmatrix[0], row[1], row[2]))
                        line_ind += 1
          
                    if attr[0] == 'Load': 
                        el_name = attr[1]   
                        self.DSSLoad.Name = el_name
                        el_name = attr[1] 
                        row[1] = str(row[1])                    
                        file_ld.write("%s \t %s \t %d \t %f \t %f \t %f \n" %(self.DSSLoad.Name, row[1], load_ind, self.DSSLoad.kW, self.DSSLoad.pf, self.DSSLoad.kW* np.tan(math.acos(self.DSSLoad.pf))))
                        load_ind += 1
                    
                    if attr[0] == 'Generator': 
                        el_name = attr[1]   
                        self.DSSGen.Name = el_name
                        el_name = attr[1] 
                        row[1] = str(row[1])                    
                        file_cp.write("%s \t %s \t %d \t %f \t %d \n" %(self.DSSGen.Name, row[1], cap_ind, self.DSSGen.Kvar,  self.DSSGen.Phases))
                        cap_ind += 1
                    
                    if attr[0] == 'Transformer': 
                        el_name = attr[1]   
                        self.DSSXfrm.Name = el_name
                        el_name = attr[1] 
                        row[1] = str(row[1]) 
                        if self.DSSXfrm.IsDelta:
                            conn = 'Delta'
                        else:
                            conn = 'wye'                   
                        file_rg.write("%s \t %s \t %d \t %s \t %f \t %f \n" %(self.DSSXfrm.Name, row[1], reg_ind, conn, self.DSSXfrm.kV, self.DSSXfrm.kva))
                        reg_ind += 1                      

if __name__ == '__main__':

    #  Handle the argument from input
    if len(sys.argv) > 2:
        print('Two many arguments')
        sys.exit()

    if len(sys.argv) < 2:
        print('Specify the File type: python filename.py JSON or python filename.py TXT')
        sys.exit()
    f = 0
    if sys.argv[1] == 'JSON' or sys.argv[1] == 'json':
        file_type = 'JSON'
        f = 1

    if sys.argv[1] == 'TXT' or sys.argv[1] == 'txt':
        file_type = 'TXT'
        f = 1

    if f == 0:
        print('File type should be TXT or JSON')
        sys.exit()

    # Change the following path to your OpenDSS Master File
    d1 = Data(fr"C:\Users\Auser\Desktop\Shiva\D-OPF\IEEE-123-Bus\Master-IEEE-123.dss")
    f1 = open("ieee123_Elements.txt","r")
    d1.Extract_Data(f1, file_type)

    
    


