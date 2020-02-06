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
import csv

class Data:
    """
    This code is for extracting the data from OpenDSS (Line parameters, Load, Graph, Xfrms). 
    Follows IEEE Common Data Format
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
        print(self.circuit)
        self.DSSLines = self.circuit.Lines
        self.DSSXfrm = self.circuit.Transformers
        self.DSSLoad = self.circuit.Loads      
        self.DSSGen =  self.circuit.Generators 
        self.DSScap =  self.circuit.Capacitors 
        if filename != "":
            self.text.Command = "compile " + filename 

        # Elements and buses will help to extract all required attribute from OpenDSS
        self.text.Command = "Show Elements"
        self.text.Command = "Show Buses"
        self.text.Command = "Export Loads"

    # Extracting the data
    def Extract_Data(self, f1, f2, f3):

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
        Buses_INFO = []
        for line in f2:
            if line.strip(): 
                row = line.split()
                bus = row[0].strip('"')
                nNodes = int(row[7])
                if nNodes == 1:
                    bus_info = dict(bus = bus.upper(),
                                    kV = float(row[1]),
                                    nNodes = nNodes,
                                    phases = [row[8], " ", " "])
                    Buses_INFO.append(bus_info)
                if nNodes == 2:
                    bus_info = dict(bus = bus.upper(),
                                kV = float(row[1]),
                                nNodes = nNodes,
                                phases = [row[8], row[9], " "])
                    Buses_INFO.append(bus_info)
                if nNodes == 3:
                    bus_info = dict(bus = bus.upper(),
                                kV = float(row[1]),
                                nNodes = nNodes,
                                phases = [row[8], row[9], row[10]])
                    Buses_INFO.append(bus_info)
        for b in Buses_INFO:
            Buses.append(b['bus'])
        sortedwords = sorted(Buses, key=len)
        longest =  sortedwords[-1]

        # Count number of elements in the circuit.       
        nBuses = len(Buses)
        nLines = 0
        nGen = 0
        nXfrm = 0
        nCap = 0        
        for el in Elements:
            row = el.split()
            element = row[0].strip('"')
            attr = element.split('.')                                    
            # Start counting                                  
            if attr[0] == 'Line': 
                nLines += 1
            if attr[0] == 'Transformer': 
                nXfrm += 1
            if attr[0] == 'Generator': 
                nGen += 1
            if attr[0] == 'Capacitor': 
                nCap += 1

        # Start polluting the txt file
        school = 'WSU'
        MVA_base = 100.0
        test_case = name[:4] + " " + name[-3:] + " Bus Test Case"
        file_ln = open(filename,"w")
        file_ln.write(" %s %s \t  %.1f \t %s  %s \n" % (today, school, MVA_base, today.year, test_case))
        file_ln.write("BUS DATA FOLLOWS \t \t %d ITEMS \n"%(nBuses))

        Load_Elements = []
        with open('ieee123_EXP_LOADS.csv') as csvfile:
            readCSV = csv.reader(csvfile, delimiter=',')
            next(readCSV)
            for row in readCSV:
                message = dict(name = row[0],
                               phases = int(row[3]))
                Load_Elements.append(message) 

        # Get the load buses. Rest buses will have None as load name and 0 power
        for el in Elements:
            row = el.split()
            element = row[0].strip('"')
            attr = element.split('.')                                    
            # Start counting                                  
            if attr[0] == 'Load': 
                for ld in Load_Elements:
                    if ld['name'] == attr[1]:
                        ld['bus'] = row[1]

        na = 0.
        for b in  Buses_INFO:  
            kW = 0
            kVAR = 0 
            flag_l = 0
            ph = b['phases']
            for ld in Load_Elements:                                                                                      
                el_name = ld['name'] 
                self.DSSLoad.Name = el_name              
                ld_bus = ld['bus']   
                if ld_bus == b['bus']:  
                    #For three phase unbalanced. The unbalanced loads are among ABC nodes. But, bus name is same for all of them
                    try:
                        ld_1 = Load_Elements[Load_Elements.index(ld)+1]             
                        if ld['phases'] == 1 and ld['bus'] == ld_1 ['bus']:  
                            ld_2 = Load_Elements[Load_Elements.index(ld)+2]  
                            kW_a = self.DSSLoad.kW    
                            kVAR_a = self.DSSLoad.kW* np.tan(math.acos(self.DSSLoad.pf))   
                            self.DSSLoad.Name = ld_1['name']
                            kW_b = self.DSSLoad.kW    
                            kVAR_b = self.DSSLoad.kW* np.tan(math.acos(self.DSSLoad.pf))
                            self.DSSLoad.Name = ld_2['name']
                            kW_c = self.DSSLoad.kW    
                            kVAR_c = self.DSSLoad.kW* np.tan(math.acos(self.DSSLoad.pf))
                            file_ln.write("%d\t" %(bus_ind) + "BUS " + ld_bus.ljust(len(longest)+5) + "%d\t" %(b['nNodes']) + \
                                        "%s\t" %(ph[0]) + "%s\t" %(ph[1]) +"%s\t" %(ph[2]) + "%.3f \t" %(b['kV']) +  \
                                        "%s\t%d\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\n" %(el_name, self.DSSLoad.Model, kW_a, kVAR_a, kW_b, kVAR_b, kW_c, kVAR_c))
                            bus_ind += 1
                            flag_l = 1
                            break         
                        # For single phase loads in single node buses
                        if ld['phases'] == 1 and ld['bus'] != ld_1 ['bus']:   
                            whichphase = b['phases']  
                            whichphase = int(whichphase[0])               
                            kW = self.DSSLoad.kW    
                            kVAR = self.DSSLoad.kW* np.tan(math.acos(self.DSSLoad.pf))   
                            if whichphase == 1:
                                file_ln.write("%d\t" %(bus_ind) + "BUS " + ld_bus.ljust(len(longest)+5) + "%d\t" %(b['nNodes']) + \
                                            "%s\t" %(ph[0]) + "%s\t" %(ph[1]) +"%s\t" %(ph[2]) + "%.3f \t" %(b['kV']) +  \
                                            "%s\t%d\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\n" %(el_name, self.DSSLoad.Model, kW, kVAR, na, na, na, na))
                                bus_ind += 1
                                flag_l = 1
                                break
                            if whichphase == 2:
                                file_ln.write("%d\t" %(bus_ind) + "BUS " + ld_bus.ljust(len(longest)+5) + "%d\t" %(b['nNodes']) + \
                                            "%s\t" %(ph[0]) + "%s\t" %(ph[1]) +"%s\t" %(ph[2]) + "%.3f \t" %(b['kV']) +  \
                                            "%s\t%d\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\n" %(el_name, self.DSSLoad.Model, na, na, kW, kVAR, na, na))
                                bus_ind += 1
                                flag_l = 1
                                break
                            if whichphase == 3:
                                file_ln.write("%d\t" %(bus_ind) + "BUS " + ld_bus.ljust(len(longest)+5) + "%d\t" %(b['nNodes']) + \
                                            "%s\t" %(ph[0]) + "%s\t" %(ph[1]) +"%s\t" %(ph[2]) + "%.3f \t" %(b['kV']) + \
                                            "%s\t%d\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\n" %(el_name, self.DSSLoad.Model, na, na, na, na, kW, kVAR))
                                bus_ind += 1
                                flag_l = 1
                                break
                    except:
                        # For single phase loads in single node buses
                        if ld['phases'] == 1:   
                            whichphase = b['phases']  
                            whichphase = int(whichphase[0])               
                            kW = self.DSSLoad.kW    
                            kVAR = self.DSSLoad.kW* np.tan(math.acos(self.DSSLoad.pf))   
                            if whichphase == 1:
                                file_ln.write("%d\t" %(bus_ind) + "BUS " + ld_bus.ljust(len(longest)+5) + "%d\t" %(b['nNodes']) + \
                                            "%s\t" %(ph[0]) + "%s\t" %(ph[1]) +"%s\t" %(ph[2]) + "%.3f \t" %(b['kV']) +  \
                                            "%s\t%d\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\n" %(el_name, self.DSSLoad.Model, kW, kVAR, na, na, na, na))
                                bus_ind += 1
                                flag_l = 1
                                break
                            if whichphase == 2:
                                file_ln.write("%d\t" %(bus_ind) + "BUS " + ld_bus.ljust(len(longest)+5) + "%d\t" %(b['nNodes']) + \
                                            "%s\t" %(ph[0]) + "%s\t" %(ph[1]) +"%s\t" %(ph[2]) + "%.3f \t" %(b['kV']) +  \
                                            "%s\t%d\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\n" %(el_name, self.DSSLoad.Model, na, na, kW, kVAR, na, na))
                                bus_ind += 1
                                flag_l = 1
                                break
                            if whichphase == 3:
                                file_ln.write("%d\t" %(bus_ind) + "BUS " + ld_bus.ljust(len(longest)+5) + "%d\t" %(b['nNodes']) + \
                                            "%s\t" %(ph[0]) + "%s\t" %(ph[1]) +"%s\t" %(ph[2]) + "%.3f \t" %(b['kV']) + \
                                            "%s\t%d\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\n" %(el_name, self.DSSLoad.Model, na, na, na, na, kW, kVAR))
                                bus_ind += 1
                                flag_l = 1
                                break                        
                        # For three phase balanced: If last load is three phase balanced, it should be inside exception handling
                        if ld['phases'] == 3:               
                            kW = self.DSSLoad.kW/3.   
                            kVAR = kW * np.tan(math.acos(self.DSSLoad.pf))   
                            file_ln.write("%d\t" %(bus_ind) + "BUS " + ld_bus.ljust(len(longest)+5) + "%d\t" %(b['nNodes']) + \
                                        "%s\t" %(ph[0]) + "%s\t" %(ph[1]) +"%s\t" %(ph[2]) + "%.3f \t" %(b['kV']) + \
                                        "%s\t%d\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\n" %(el_name, self.DSSLoad.Model, kW, kVAR, kW, kVAR, kW, kVAR))
                            bus_ind += 1
                            flag_l = 1
                            break
                    # For three phase balanced
                    if ld['phases'] == 3:               
                        kW = self.DSSLoad.kW/3.   
                        kVAR = kW * np.tan(math.acos(self.DSSLoad.pf))   
                        file_ln.write("%d\t" %(bus_ind) + "BUS " + ld_bus.ljust(len(longest)+5) + "%d\t" %(b['nNodes']) + \
                                    "%s\t" %(ph[0]) + "%s\t" %(ph[1]) +"%s\t" %(ph[2]) + "%.3f \t" %(b['kV']) + \
                                    "%s\t%d\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\n" %(el_name, self.DSSLoad.Model, kW, kVAR, kW, kVAR, kW, kVAR))
                        bus_ind += 1
                        flag_l = 1
                        break
                    flag_l = 1
                    break                
            if flag_l == 0:
                file_ln.write("%d\t" %(bus_ind) + "BUS " + b['bus'].ljust(len(longest)+5) + "%d\t" %(b['nNodes']) + \
                            "%s\t" %(ph[0]) + "%s\t" %(ph[1]) +"%s\t" %(ph[2]) + "%.3f \t" %(b['kV']) + \
                            "%s\t%s\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\t%.3f\n" %('None', 'N/A', kW, kVAR, kW, kVAR, kW, kVAR))
                bus_ind += 1
        file_ln.write("-999 \n")

        # Fill LINE data in txt file   
        file_ln.write("LINE DATA FOLLOWS \t \t %d ITEMS \n"%(nLines))
        for line in Elements:
            if line.strip():                               
                row = line.split()
                element = row[0].strip('"')
                attr = element.split('.')                                    
                # Store distribution lines from the network model             
                if attr[0] == 'Line': 
                    el_name = attr[1]   
                    self.DSSLines.Name = el_name
                    row[1] = str(row[1])   
                    row[2] = str(row[2])                  
                    file_ln.write("%d\t" %(line_ind) + "LINE " + self.DSSLines.Name.ljust(8) + "%d\t" %(self.DSSLines.Phases)+ "%.4f\t" %(self.DSSLines.Length))
                    file_ln.write(row[1].ljust(len(longest)+5) + row[2].ljust(len(longest)+5) + "%f\t%f\n" %(self.DSSLines.Rmatrix[0], self.DSSLines.Xmatrix[0]))
                    line_ind += 1
        file_ln.write("-999 \n")

        # Fill GENERATOR data in txt file
        file_ln.write("GENERATOR DATA FOLLOWS \t \t %d ITEMS \n"%(nGen))
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
                    file_ln.write("%.4f\t%.4f\n" %(self.DSSGen.Kvar, self.DSSGen.kV))
                    cap_ind += 1
        file_ln.write("-999 \n")

        # Fill CAPACITOR data in txt file
        file_ln.write("CAPACITOR DATA FOLLOWS \t \t %d ITEMS \n"%(nCap))
        for line in Elements:
            if line.strip():                               
                row = line.split()
                element = row[0].strip('"')
                attr = element.split('.')                                    
                # Store distribution lines from the network model             
                if attr[0] == 'Capacitor': 
                    el_name = attr[1]   
                    self.DSScap.Name = el_name
                    el_name = attr[1] 
                    row[1] = str(row[1])                
                    file_ln.write("%d\t" %(cap_ind) + "CAP " + "%s\t" %(self.DSScap.Name) + "\t" + row[1].ljust(len(longest)+5))
                    file_ln.write("%.4f\t%.4f\n" %(self.DSScap.kvar, self.DSScap.kV))
                    cap_ind += 1
        file_ln.write("-999 \n")
        
        # Fill XFMR data in txt file
        file_ln.write("TRANSFORMER DATA FOLLOWS \t \t %d ITEMS \n"%(nXfrm))
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
    # d1 = Data(fr"C:\Users\Auser\Desktop\Shiva\D-OPF\IEEE-123-Bus-3Phase\IEEE123Master.dss")
    f1 = open("ieee123_Elements.txt","r")
    f2 = open("ieee123_Buses.txt","r")
    f3 = open("ieee123_EXP_LOADS.csv","r")

    d1.Extract_Data(f1, f2, f3)