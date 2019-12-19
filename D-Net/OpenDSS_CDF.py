import win32com.client
import matplotlib.pyplot as plt
import json
import math
import numpy as np
import networkx as nx
import os
import sys
import argparse
from datetime import date

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

        # Elements and buses will help to extract all required attribute from OpenDSS
        self.text.Command = "Show Elements"
        self.text.Command = "Show Buses"

    # Checking the interface
    def Extract_Data(self, f1, f2):

        # File starts from 7th line always in Elements.txt file  
        line_ind = 0    
        xfmr_ind = 0    
        cap_ind = 0
        bus_ind = 0
        today = date.today()
        for i in range(6):
            next (f2)  

        for line in f1:
            if line.strip(): 
                row = line.split()
                name = row[4].upper()
                break

        # File starts from 7th line always in Elements.txt file and two already used in finding the feeder name 
        for i in range(5):
            next (f1)  
        filename = "%s.txt" % name
        Elements = []
        # Store all elements from the feeder
        for line in f1:
            if line.strip():   
                Elements.append(line)
        
        # Store buses and find the longest name for proper indent in txt files
        Buses = []
        for line in f2:
            if line.strip(): 
                row = line.split()
                bus = row[0].strip('"')
                Buses.append(bus)
        sortedwords = sorted(Buses, key=len)
        longest =  sortedwords[-1]

        # Start polluting the txt file
        school = 'WSU'
        MVA_base = 100.0
        test_case = name[:4] + " " + name[-3:] + " Bus Test Case"
        file_ln = open(filename,"w")
        file_ln.write(" %s %s \t  %.1f \t %s  %s \n" % (today, school, MVA_base, today.year, test_case))
        file_ln.write("BUS DATA FOLLOWS \n")

        # Get the load buses. Rest buses will have None as load name and 0 power
        load_bus = []
        for line in Elements:                                                    
                row = line.split()
                element = row[0].strip('"')
                attr = element.split('.')                                    
                # Store distribution loads from the network model                                  
                if attr[0] == 'Load': 
                    el_name = attr[1]   
                    self.DSSLoad.Name = el_name
                    el_name = attr[1] 
                    row[1] = str(row[1])
                    load_bus.append(row[1])

        # Fill BUS data in txt file
        for bus in  Buses:  
            # print(bus)
            kW = 0
            kVAR = 0 
            flag_l = 0
            for line in Elements:                                                    
                row = line.split()
                element = row[0].strip('"')
                attr = element.split('.')                                    
                # Store distribution loads from the network model                                  
                if attr[0] == 'Load': 
                    el_name = attr[1]   
                    self.DSSLoad.Name = el_name
                    el_name = attr[1] 
                    ld_bus = str(row[1])   
                    # print(el_name)
                    if ld_bus == bus:                        
                        kW = self.DSSLoad.kW    
                        kVAR = self.DSSLoad.kW* np.tan(math.acos(self.DSSLoad.pf))   
                        file_ln.write("%d\t" %(bus_ind) + "BUS " + ld_bus.ljust(len(longest)+5) + "%s\t%.3f\t%.3f\n" %( el_name,kW,kVAR))
                        bus_ind += 1
                        flag_l = 1
                        break
            if flag_l == 0:
                kW = 0  
                kVAR = 0
                file_ln.write("%d\t" %(bus_ind) + "BUS " + bus.ljust(len(longest)+5) + "%s\t%.3f\t%.3f\n" %('None',kW,kVAR))
                bus_ind += 1

        # Fill BUS data in txt file   
        file_ln.write("\nLINE DATA FOLLOWS \n")
        for line in Elements:
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
                    line_type = 0                
                    file_ln.write("%d\t" %(line_ind) + "LINE " + self.DSSLines.Name.ljust(8) + "%d\t" %(line_type))
                    file_ln.write(row[1].ljust(len(longest)+5) + row[2].ljust(len(longest)+5) + "%f\t%f\n" %(self.DSSLines.Rmatrix[0], self.DSSLines.Xmatrix[0]))
                    line_ind += 1
        
        # Fill GENERATOR data in txt file
        file_ln.write("\nGENERATOR DATA FOLLOWS \n")
        for line in Elements:
            if line.strip():                               
                row = line.split()
                element = row[0].strip('"')
                attr = element.split('.')                                    
                # Store distribution lines from the network model             
                if attr[0] == 'Generator': 
                    el_name = attr[1]   
                    self.DSSGen.Name = el_name
                    el_name = attr[1] 
                    row[1] = str(row[1])                
                    file_ln.write("%d\t" %(cap_ind) + "GEN " + "%s\t" %(self.DSSGen.Name) + "\t" + row[1].ljust(len(longest)+5))
                    file_ln.write("%.4f\t%d\n" %(self.DSSGen.Kvar, self.DSSGen.Phases))
                    cap_ind += 1
        
        # Fill XFMR data in txt file
        file_ln.write("\nTRANSFORMER DATA FOLLOWS \n")
        for line in Elements:
            if line.strip():                               
                row = line.split()
                element = row[0].strip('"')
                attr = element.split('.')                                    
                # Store distribution lines from the network model             
                if attr[0] == 'Transformer': 
                    line_type = 1    
                    el_name = attr[1]   
                    self.DSSXfrm.Name = el_name
                    el_name = attr[1] 
                    row[1] = str(row[1])   
                    row[2] = str(row[2])                  
                    file_ln.write("%d\t" %(xfmr_ind) + "XFMR " + "%s\t%d\t" %(self.DSSXfrm.Name,line_type))
                    file_ln.write(row[1].ljust(len(longest)+5) + row[2].ljust(len(longest)+5) + "%f\t%f\n" %(self.DSSXfrm.R, self.DSSXfrm.Xhl))
                    xfmr_ind += 1      

if __name__ == '__main__':

    # Change the following path to your OpenDSS Master File
    d1 = Data(fr"C:\Users\Auser\Desktop\Shiva\D-OPF\IEEE-123-Bus\Master-IEEE-123.dss")
    f1 = open("ieee123_Elements.txt","r")
    f2 = open("ieee123_Buses.txt","r")
    d1.Extract_Data(f1, f2)

    
    


