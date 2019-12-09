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
        if filename != "":
            self.text.Command = "compile " + filename 
        
        # self.text.Command = "solve" 
        # self.text.Command = "Show Voltages LN Nodes"
        self.text.Command = "Show Elements"
        
    # Checking the interface
    def Extract_Data(self, f):

        Linepar = []
        # File starts from 7th line always in Elements.txt file  
        for i in range(7):
            next (f)  
        
        for line in f:
            if line.strip():                               
                row = line.split()
                element = row[0].strip('"')
                attr = element.split('.') 

                # Store distribution lines from the network model
                line_ind = 0                 
                if attr[0] == 'Line': 
                    el_name = attr[1]               
                    self.DSSLines.Name = el_name
                    message = dict(Line = self.DSSLines.Name,
                                   index = line_ind,
                                   from_br = row[1],
                                   r = self.DSSLines.Rmatrix,
                                   x = self.DSSLines.Xmatrix,
                                   to_br = row[2])
                    line_ind += 1
                    Linepar.append(message)
                print(Linepar)
                    
        
    

if __name__ == '__main__':
    d1 = Data(fr"C:\Users\Auser\Desktop\Shiva\D-OPF\IEEE-123-Bus\Master-IEEE-123.dss")

    # print(os.listdir())
    # for f in os.listdir():
    #     print(f)

    f1 = open("ieee123_Elements.txt","r")
    d1.Extract_Data(f1)
    


