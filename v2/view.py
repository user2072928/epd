import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd  
from db import db
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, Image
from reportlab.lib.pagesizes import  A4
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import numpy as np
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.charts.barcharts import VerticalBarChart



# Input
class InputFrame(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        # define variable
        self.root = master
        self.ProductName = tk.StringVar()
        self.TotalWeight = tk.StringVar()
        self.EnergyUse = tk.StringVar()
        self.PowerConsumption = tk.StringVar()
        self.ServiceLife = tk.StringVar()
        self.TransporttoFactoryl = tk.StringVar()
        self.TransporttoFactoryr = tk.StringVar()
        self.TransporttoFactorys = tk.StringVar()
        self.TransporttoSite = tk.StringVar()
        self.TransporttoLandfill = tk.StringVar()
        
        self.ABS = tk.StringVar()
        self.Aluminium = tk.StringVar()
        self.Brass = tk.StringVar()
        self.CastIron = tk.StringVar()
        self.Ceramic = tk.StringVar()
        self.Copper = tk.StringVar()
        self.ElectronicComponent = tk.StringVar()
        self.ExpandedPolystyrene = tk.StringVar()
        self.Glass = tk.StringVar()        
        self.Insulation = tk.StringVar()
        self.Iron = tk.StringVar()
        self.Lithium = tk.StringVar()
        self.Plastics = tk.StringVar()
        self.Polyamide = tk.StringVar()
        self.Polycarbonate = tk.StringVar()
        self.Polyethylene  = tk.StringVar()          
        self.PolyurethaneFoam = tk.StringVar()
        self.PrintedBoard = tk.StringVar()
        self.PVCPipe = tk.StringVar()
        self.PVC = tk.StringVar()
        self.Rubber  = tk.StringVar()
        self.Silicon = tk.StringVar()
        self.StainlessSteel = tk.StringVar()
        self.Steel = tk.StringVar()
        self.Zinc  = tk.StringVar()

        self.ABS1 = tk.StringVar()
        self.Aluminium1 = tk.StringVar()
        self.Brass1 = tk.StringVar()
        self.CastIron1 = tk.StringVar()
        self.Ceramic1 = tk.StringVar()
        self.Copper1 = tk.StringVar()
        self.ElectronicComponent1 = tk.StringVar()
        self.ExpandedPolystyrene1 = tk.StringVar()
        self.Glass1 = tk.StringVar()        
        self.Insulation1 = tk.StringVar()
        self.Iron1 = tk.StringVar()
        self.Lithium1 = tk.StringVar()
        self.Plastics1 = tk.StringVar()
        self.Polyamide1 = tk.StringVar()
        self.Polycarbonate1 = tk.StringVar()
        self.Polyethylene1  = tk.StringVar()          
        self.PolyurethaneFoam1 = tk.StringVar()
        self.PrintedBoard1 = tk.StringVar()
        self.PVCPipe1 = tk.StringVar()
        self.PVC1 = tk.StringVar()
        self.Rubber1  = tk.StringVar()
        self.Silicon1 = tk.StringVar()
        self.StainlessSteel1 = tk.StringVar()
        self.Steel1 = tk.StringVar()
        self.Zinc1  = tk.StringVar()            
        
        self.status = tk.StringVar()
        self.create_page()
        
    def create_page(self):
        # place label and entry
        tk.Label(self).grid(row=0, stick=tk.W, pady=2)
        
        tk.Label(self, text="Product Information ").grid(row=1, padx=20, stick=tk.W, pady=2) 
        
        tk.Label(self, text="Product Name: ").grid(row=2, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.ProductName).grid(row=2, padx=20, column=1, stick=tk.E)

        tk.Label(self, text="Total Weight: ").grid(row=3, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.TotalWeight).grid(row=3, padx=20, column=1, stick=tk.E)

        tk.Label(self, text="Energy Use Pre Product: ").grid(row=4, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.EnergyUse).grid(row=4, padx=20, column=1, stick=tk.E)
        
        tk.Label(self, text="Power Consumption: ").grid(row=5, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.PowerConsumption).grid(row=5, padx=20, column=1, stick=tk.E)
        
        tk.Label(self, text="Service Life: ").grid(row=6, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.ServiceLife).grid(row=6, padx=20, column=1, stick=tk.E)
 
        tk.Label(self, text="Transport to Factory(Local): ").grid(row=7, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.TransporttoFactoryl).grid(row=7, padx=20, column=1, stick=tk.E)
        
        tk.Label(self, text="Transport to Factory(Road): ").grid(row=8, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.TransporttoFactoryr).grid(row=8, padx=20, column=1, stick=tk.E)
        
        tk.Label(self, text="Transport to Factory(Sea): ").grid(row=9, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.TransporttoFactorys).grid(row=9, padx=20, column=1, stick=tk.E)
        
        tk.Label(self, text="Transport to Site: ").grid(row=10, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.TransporttoSite).grid(row=10, padx=20, column=1, stick=tk.E)

        tk.Label(self, text="Transport to Landfill: ").grid(row=11, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.TransporttoLandfill).grid(row=11, padx=20, column=1, stick=tk.E) 


        tk.Label(self, text="Component Information ").grid(row=1, padx=20, column=2, stick=tk.W, pady=2) 
        tk.Label(self, text="Remanufactured Product ").grid(row=1, padx=20, column=3, stick=tk.W, pady=2) 
        tk.Label(self, text="New Product ").grid(row=1, padx=20, column=4, stick=tk.W, pady=2) 
        
        tk.Label(self, text="ABS %: ").grid(row=2, padx=20, column=2,stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.ABS).grid(row=2, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.ABS1).grid(row=2, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Aluminium %: ").grid(row=3, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Aluminium).grid(row=3, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Aluminium1).grid(row=3, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Brass %: ").grid(row=4, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Brass).grid(row=4, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Brass1).grid(row=4, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Cast iron %: ").grid(row=5, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.CastIron).grid(row=5, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.CastIron1).grid(row=5, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Ceramic %: ").grid(row=6, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Ceramic).grid(row=6, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Ceramic1).grid(row=6, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Copper %: ").grid(row=7, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Copper).grid(row=7, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Copper1).grid(row=7, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Electronic component %: ").grid(row=8, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.ElectronicComponent).grid(row=8, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.ElectronicComponent1).grid(row=8, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Expanded polystyrene %: ").grid(row=9, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.ExpandedPolystyrene).grid(row=9, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.ExpandedPolystyrene1).grid(row=9, padx=20, column=4, stick=tk.E)
       
        tk.Label(self, text="Glass %: ").grid(row=10, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Glass).grid(row=10, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Glass1).grid(row=10, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Insulation (general) %: ").grid(row=11, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Insulation).grid(row=11, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Insulation1).grid(row=11, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Iron %: ").grid(row=12, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Iron).grid(row=12, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Iron1).grid(row=12, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Lithium %: ").grid(row=13, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Lithium).grid(row=13, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Lithium1).grid(row=13, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Plastics (general) %: ").grid(row=14, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Plastics).grid(row=14, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Plastics1).grid(row=14, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Polyamide %: ").grid(row=15, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Polyamide).grid(row=15, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Polyamide1).grid(row=15, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Polycarbonate %: ").grid(row=16, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Polycarbonate).grid(row=16, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Polycarbonate1).grid(row=16, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Polyethylene %: ").grid(row=17, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Polyethylene).grid(row=17, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Polyethylene1).grid(row=17, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Polyurethane foam %: ").grid(row=18, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.PolyurethaneFoam).grid(row=18, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.PolyurethaneFoam1).grid(row=18, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Printed wiring board, mixed mounted %: ").grid(row=19, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.PrintedBoard).grid(row=19, padx=20, column=3, stick=tk.E) 
        tk.Entry(self, textvariable=self.PrintedBoard1).grid(row=19, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="PVC pipe %: ").grid(row=20, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.PVCPipe).grid(row=20, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.PVCPipe1).grid(row=20, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="PVC %: ").grid(row=21, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.PVC).grid(row=21, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.PVC1).grid(row=21, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Rubber %: ").grid(row=22, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Rubber).grid(row=22, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Rubber1).grid(row=22, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Silicon %: ").grid(row=23, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Silicon).grid(row=23, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Silicon1).grid(row=23, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Stainless steel %: ").grid(row=24, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.StainlessSteel).grid(row=24, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.StainlessSteel1).grid(row=24, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Steel (general or galvanised) %: ").grid(row=25, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Steel).grid(row=25, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Steel1).grid(row=25, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Zinc %: ").grid(row=26, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Zinc).grid(row=26, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Zinc1).grid(row=26, padx=20, column=4, stick=tk.E)
                 
        tk.Button(self, text="Record", command=self.record_product).grid(row=27, column=1, stick=tk.E, pady=2)
        tk.Label(self, textvariable=self.status).grid(row=28, column=1, stick=tk.E, pady=2 )

    def record_product(self):
        product = {
            "Product Name": self.ProductName.get(),
            "Total Weight": self.TotalWeight.get(),
            "Energy Use Pre Product": self.EnergyUse.get(),  
            "Power Consumption": self.PowerConsumption.get(),
            "Service Life": self.ServiceLife.get(),
            "Transport to Factory(Local)": self.TransporttoFactoryl.get(),
            "Transport to Factory(Road)": self.TransporttoFactoryr.get(),
            "Transport to Factory(Sea)": self.TransporttoFactorys.get(),            
            "Transport to Site": self.TransporttoSite.get(),
            "Transport to Landfill": self.TransporttoLandfill.get(), 
            
            "ABS %": self.ABS.get(),
            "Aluminium %": self.Aluminium.get(),
            "Brass %": self.Brass.get(),
            "Cast iron %": self.CastIron.get(),  
            "Ceramic %": self.Ceramic.get(),
            "Copper %": self.Copper.get(),
            "Electronic component %": self.ElectronicComponent.get(),
            "Expanded polystyrene %": self.ExpandedPolystyrene.get(),
            "Glass %": self.Glass.get(),           
            "Insulation (general) %": self.Insulation.get(),
            "Iron %": self.Iron.get(),
            "Lithium %": self.Lithium.get(),
            "Plastics (general) %": self.Plastics.get(),  
            "Polyamide %": self.Polyamide.get(),
            "Polycarbonate %": self.Polycarbonate.get(),
            "Polyethylene %": self.Polyethylene.get(),
            "Polyurethane foam %": self.PolyurethaneFoam.get(),
            "Printed wiring board, mixed mounted %": self.PrintedBoard.get(),             
            "PVC pipe %": self.PVCPipe.get(),
            "PVC %": self.PVC.get(),
            "Rubber %": self.Rubber.get(),
            "Silicon %": self.Silicon.get(),  
            "Stainless steel %": self.StainlessSteel.get(),
            "Steel (general or galvanised) %": self.Steel.get(),
            "Zinc %": self.Zinc.get(),
            
            "ABS(new) %": self.ABS1.get(),
            "Aluminium(new) %": self.Aluminium1.get(),
            "Brass(new) %": self.Brass1.get(),
            "Cast iron(new) %": self.CastIron1.get(),  
            "Ceramic(new) %": self.Ceramic1.get(),
            "Copper(new) %": self.Copper1.get(),
            "Electronic component(new) %": self.ElectronicComponent1.get(),
            "Expanded polystyrene(new) %": self.ExpandedPolystyrene1.get(),
            "Glass(new) %": self.Glass1.get(),           
            "Insulation (general)(new) %": self.Insulation1.get(),
            "Iron(new) %": self.Iron1.get(),
            "Lithium(new) %": self.Lithium1.get(),
            "Plastics (general)(new) %": self.Plastics1.get(),  
            "Polyamide(new) %": self.Polyamide1.get(),
            "Polycarbonate(new) %": self.Polycarbonate1.get(),
            "Polyethylene(new) %": self.Polyethylene1.get(),
            "Polyurethane foam(new) %": self.PolyurethaneFoam1.get(),
            "Printed wiring board, mixed mounted(new) %": self.PrintedBoard1.get(),             
            "PVC pipe(new) %": self.PVCPipe1.get(),
            "PVC(new) %": self.PVC1.get(),
            "Rubber(new) %": self.Rubber1.get(),
            "Silicon(new) %": self.Silicon1.get(),  
            "Stainless steel(new) %": self.StainlessSteel1.get(),
            "Steel (general or galvanised)(new) %": self.Steel1.get(),
            "Zinc(new) %": self.Zinc1.get()            
        }  # A product
        
        db.insert(product)
        db.save_data()

        self.status.set("Record Successfully")
        self.clear_data()

    # Clean
    def clear_data(self):
        self.ProductName.set("")
        self.TotalWeight.set("")
        self.EnergyUse.set("")
        self.PowerConsumption.set("")
        self.ServiceLife.set("")
        self.TransporttoFactoryl.set("")
        self.TransporttoFactoryr.set("")
        self.TransporttoFactorys.set("")
        self.TransporttoSite.set("")
        self.TransporttoLandfill.set("")
        
        self.ABS.set("")
        self.Aluminium.set("")
        self.Brass.set("")
        self.CastIron.set("")
        self.Ceramic.set("")
        self.Copper.set("")
        self.ElectronicComponent.set("")
        self.ExpandedPolystyrene.set("")
        self.Glass.set("")    
        self.Insulation.set("")
        self.Iron.set("")
        self.Lithium.set("")
        self.Plastics.set("")
        self.Polyamide.set("")
        self.Polycarbonate.set("")
        self.Polyethylene.set("")         
        self.PolyurethaneFoam.set("")
        self.PrintedBoard.set("")
        self.PVCPipe.set("")
        self.PVC.set("")
        self.Rubber.set("")
        self.Silicon.set("")
        self.StainlessSteel.set("")
        self.Steel.set("")
        self.Zinc.set("")  
        
        self.ABS1.set("")
        self.Aluminium1.set("")
        self.Brass1.set("")
        self.CastIron1.set("")
        self.Ceramic1.set("")
        self.Copper1.set("")
        self.ElectronicComponent1.set("")
        self.ExpandedPolystyrene1.set("")
        self.Glass1.set("")    
        self.Insulation1.set("")
        self.Iron1.set("")
        self.Lithium1.set("")
        self.Plastics1.set("")
        self.Polyamide1.set("")
        self.Polycarbonate1.set("")
        self.Polyethylene1.set("")         
        self.PolyurethaneFoam1.set("")
        self.PrintedBoard1.set("")
        self.PVCPipe1.set("")
        self.PVC1.set("")
        self.Rubber1.set("")
        self.Silicon1.set("")
        self.StainlessSteel1.set("")
        self.Steel1.set("")
        self.Zinc1.set("")   

# Upload
class UploadFrame(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.root = master
        
        frame1 = tk.LabelFrame(self, text="Product BoM")
        frame1.grid(row=0, padx=20, ipadx=700, column=0, pady=2, ipady=248, stick=tk.W)     
        
        file_frame = tk.LabelFrame(self, text="Open File")
        file_frame.grid(row=1, padx=500, ipadx=200, column=0, pady=50, ipady=50, stick=tk.W)

        b1 = tk.Button(file_frame, text='Upload File', command=lambda: file_dialog())
        b1.place(rely=0.5, relx=0.4)
         
        label_file = ttk.Label(file_frame, text="No File Selected")
        label_file.place(rely=0, relx=0)
                
        tv1 = ttk.Treeview(frame1)
        tv1.place(relheight=1, relwidth=1)
        
        # set x y scrollbar        
        treescrolly = tk.Scrollbar(frame1, orient="vertical", command=tv1.yview)
        treescrollx = tk.Scrollbar(frame1, orient="horizontal", command=tv1.xview)
        tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
                
        treescrollx.pack(side="bottom", fill="x")
        treescrolly.pack(side="right", fill="y")
        
        def file_dialog():
            #open excel file
             filename = fd.askopenfilename(initialdir="/", title="Select A File", filetypes=(("xlsx file", "*.xlsx"),("All file", "*.*")))
            # read different sheets
             bom = pd.read_excel(filename, sheet_name="BOM", engine="openpyxl")
             spec = pd.read_excel(filename, sheet_name="Product Spec & Assumptions", engine="openpyxl")
             ref = pd.read_excel(filename, sheet_name="Reference", engine="openpyxl")
             stageA = pd.read_excel(filename, sheet_name="Stage A", engine="openpyxl")
             stageC = pd.read_excel(filename, sheet_name="Stage C", engine="openpyxl")
             mf = pd.read_excel(filename, sheet_name="Manufacturer Form", engine="openpyxl")
             
             #clean data
             bom.replace(" ","0",inplace = True)
             bom.replace("-","0",inplace = True)
             bom = bom.fillna("Delete")
             bom.replace("Delete","",inplace = True)
             bom.replace("'-"," ",inplace = True)

             product = {
                 "Product Name": spec.iloc[2,2],
                 "Total Weight": spec.iloc[21,2],
                 "Energy Use Pre Product": spec.iloc[24,2],  
                 "Power Consumption": spec.iloc[20,2],
                 "Service Life": spec.iloc[23,2],
                 "Transport to Factory(Local)": ref.iloc[51,2],
                 "Transport to Factory(Road)": ref.iloc[54,2],
                 "Transport to Factory(Sea)": ref.iloc[54,3],            
                 "Transport to Site": stageA.iloc[55,2],
                 "Transport to Landfill": stageC.iloc[6,2], 
                 
                 "ABS %": mf.iloc[16,3]*100,
                 "Aluminium %": mf.iloc[17,3]*100,
                 "Brass %": mf.iloc[18,3]*100,
                 "Cast iron %": mf.iloc[19,3]*100,  
                 "Ceramic %": mf.iloc[20,3]*100,
                 "Copper %": mf.iloc[21,3]*100,
                 "Electronic component %": mf.iloc[22,3]*100,
                 "Expanded polystyrene %": mf.iloc[23,3]*100,
                 "Glass %": mf.iloc[24,3]*100,           
                 "Insulation (general) %": mf.iloc[25,3]*100,
                 "Iron %": mf.iloc[26,3]*100,
                 "Lithium %": mf.iloc[27,3]*100,
                 "Plastics (general) %": mf.iloc[28,3]*100,  
                 "Polyamide %": mf.iloc[29,3]*100,
                 "Polycarbonate %": mf.iloc[30,3]*100,
                 "Polyethylene %": mf.iloc[31,3]*100,
                 "Polyurethane foam %": mf.iloc[32,3]*100,
                 "Printed wiring board, mixed mounted %":mf.iloc[33,3]*100,             
                 "PVC pipe %": mf.iloc[34,3]*100,
                 "PVC %": mf.iloc[35,3]*100,
                 "Rubber %": mf.iloc[36,3]*100,
                 "Silicon %": mf.iloc[37,3]*100,  
                 "Stainless steel %": mf.iloc[38,3]*100,
                 "Steel (general or galvanised) %":mf.iloc[39,3]*100,
                 "Zinc %": mf.iloc[40,3]*100,
                 
                 "ABS(new) %": mf.iloc[16,5]*100,
                 "Aluminium(new) %": mf.iloc[17,5]*100,
                 "Brass(new) %": mf.iloc[18,5]*100,
                 "Cast iron(new) %": mf.iloc[19,5]*100,  
                 "Ceramic(new) %": mf.iloc[20,5]*100,
                 "Copper(new) %": mf.iloc[21,5]*100,
                 "Electronic component(new) %": mf.iloc[22,5]*100,
                 "Expanded polystyrene(new) %": mf.iloc[23,5]*100,
                 "Glass(new) %": mf.iloc[24,5]*100,           
                 "Insulation (general)(new) %": mf.iloc[25,5]*100,
                 "Iron(new) %": mf.iloc[26,5]*100,
                 "Lithium(new) %": mf.iloc[27,5]*100,
                 "Plastics (general)(new) %": mf.iloc[28,5]*100,  
                 "Polyamide(new) %": mf.iloc[29,5]*100,
                 "Polycarbonate(new) %": mf.iloc[30,5]*100,
                 "Polyethylene(new) %": mf.iloc[31,5]*100,
                 "Polyurethane foam(new) %": mf.iloc[32,5]*100,
                 "Printed wiring board, mixed mounted(new) %": mf.iloc[33,5]*100,             
                 "PVC pipe(new) %": mf.iloc[34,5]*100,
                 "PVC(new) %": mf.iloc[35,5]*100,
                 "Rubber(new) %": mf.iloc[36,5]*100,
                 "Silicon(new) %": mf.iloc[37,5]*100,  
                 "Stainless steel(new) %": mf.iloc[38,5]*100,
                 "Steel (general or galvanised)(new) %": mf.iloc[39,5]*100,
                 "Zinc(new) %": mf.iloc[40,5]*100            
             }  # A product
             
             db.insert(product)
             db.save_data()

             clear_data()
             
             # add data in diaplay box
             tv1["column"] = list(bom.columns)
             tv1["show"] = "headings"
             for column in tv1["column"]:
                    tv1.heading(column,text=column)
             bom_rows = bom.to_numpy().tolist()
             for row in bom_rows:
                  tv1.insert("", "end", values=row)

        def clear_data():
            tv1.delete(*tv1.get_children())    

# Query
class QueryFrame(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)

        self.table_view = tk.LabelFrame(text="Product Database")
        self.table_view.pack()
        self.status = tk.StringVar()
        self.qu_name = tk.StringVar()
        self.create_page()

    def create_page(self):

        columns = ("Product Name",)
        self.tree_view = ttk.Treeview(self,show = "headings", columns = columns)
        self.tree_view.column("Product Name",width =500, anchor="center")
        self.tree_view.heading("Product Name",text="Product Name")



        self.tree_view.pack(fill =tk.BOTH, expand = True)
        self.fresh_data()
    
                
        treescrolly = tk.Scrollbar(self, orient="vertical", command=self.tree_view.yview)
        treescrollx = tk.Scrollbar(self, orient="horizontal", command=self.tree_view.xview)
        self.tree_view.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
                
        treescrollx.pack(side="bottom", fill="x")
        treescrolly.pack(side="right", fill="y")
      
        tk.Button(self, text='Refresh Product Database', command=self.fresh_data).pack()
        tk.Label(self, text="Query EPD by Name").pack()
        e1 = tk.Entry(self, textvariable=self.qu_name)
        e1.pack()

        tk.Button(self, text='Query', command=self.epd).pack()
        tk.Label(self, textvariable=self.status).pack()
        
        
    def fresh_data(self):
        self.tree_view.delete(*self.tree_view.get_children()) 
        products = db.all()
        for product in products:
            self.tree_view.insert("","end", values=(product["Product Name"],))


    def epd(self):
        ProductName = self.qu_name.get()
        product = db.search_by_name(ProductName)

        if product:
            self.ProductName = tk.StringVar()
            self.TotalWeight = tk.StringVar()
            self.EnergyUse = tk.StringVar()
            self.PowerConsumption = tk.StringVar()
            self.ServiceLife = tk.StringVar()
            self.TransporttoFactoryl = tk.StringVar()
            self.TransporttoFactoryr = tk.StringVar()
            self.TransporttoFactorys = tk.StringVar()
            self.TransporttoSite = tk.StringVar()
            self.TransporttoLandfill = tk.StringVar()
            
            self.ABS = tk.StringVar()
            self.Aluminium = tk.StringVar()
            self.Brass = tk.StringVar()
            self.CastIron = tk.StringVar()
            self.Ceramic = tk.StringVar()
            self.Copper = tk.StringVar()
            self.ElectronicComponent = tk.StringVar()
            self.ExpandedPolystyrene = tk.StringVar()
            self.Glass = tk.StringVar()        
            self.Insulation = tk.StringVar()
            self.Iron = tk.StringVar()
            self.Lithium = tk.StringVar()
            self.Plastics = tk.StringVar()
            self.Polyamide = tk.StringVar()
            self.Polycarbonate = tk.StringVar()
            self.Polyethylene  = tk.StringVar()          
            self.PolyurethaneFoam = tk.StringVar()
            self.PrintedBoard = tk.StringVar()
            self.PVCPipe = tk.StringVar()
            self.PVC = tk.StringVar()
            self.Rubber  = tk.StringVar()
            self.Silicon = tk.StringVar()
            self.StainlessSteel = tk.StringVar()
            self.Steel = tk.StringVar()
            self.Zinc  = tk.StringVar()
            
            self.ABS1 = tk.StringVar()
            self.Aluminium1 = tk.StringVar()
            self.Brass1 = tk.StringVar()
            self.CastIron1 = tk.StringVar()
            self.Ceramic1 = tk.StringVar()
            self.Copper1 = tk.StringVar()
            self.ElectronicComponent1 = tk.StringVar()
            self.ExpandedPolystyrene1 = tk.StringVar()
            self.Glass1 = tk.StringVar()        
            self.Insulation1 = tk.StringVar()
            self.Iron1 = tk.StringVar()
            self.Lithium1 = tk.StringVar()
            self.Plastics1 = tk.StringVar()
            self.Polyamide1 = tk.StringVar()
            self.Polycarbonate1 = tk.StringVar()
            self.Polyethylene1  = tk.StringVar()          
            self.PolyurethaneFoam1 = tk.StringVar()
            self.PrintedBoard1 = tk.StringVar()
            self.PVCPipe1 = tk.StringVar()
            self.PVC1 = tk.StringVar()
            self.Rubber1  = tk.StringVar()
            self.Silicon1 = tk.StringVar()
            self.StainlessSteel1 = tk.StringVar()
            self.Steel1 = tk.StringVar()
            self.Zinc1  = tk.StringVar()
            
            self.ProductName.set(product["Product Name"])
            self.TotalWeight.set(product["Total Weight"])
            self.EnergyUse.set(product["Energy Use Pre Product"])
            self.PowerConsumption.set(product["Power Consumption"])
            self.ServiceLife.set(product["Service Life"])
            self.TransporttoFactoryl.set(product["Transport to Factory(Local)"])
            self.TransporttoFactoryr.set(product["Transport to Factory(Road)"])
            self.TransporttoFactorys.set(product["Transport to Factory(Sea)"])
            self.TransporttoSite.set(product["Transport to Site"])
            self.TransporttoLandfill.set(product["Transport to Landfill"])
            
            self.ABS.set(product["ABS %"])
            self.Aluminium.set(product["Aluminium %"])
            self.Brass.set(product["Brass %"])
            self.CastIron.set(product["Cast iron %"])
            self.Ceramic.set(product["Ceramic %"])
            self.Copper.set(product["Copper %"])
            self.ElectronicComponent.set(product["Electronic component %"])
            self.ExpandedPolystyrene.set(product["Expanded polystyrene %"])
            self.Glass.set(product["Glass %"])    
            self.Insulation.set(product["Insulation (general) %"])
            self.Iron.set(product["Iron %"])
            self.Lithium.set(product["Lithium %"])
            self.Plastics.set(product["Plastics (general) %"])
            self.Polyamide.set(product["Polyamide %"])
            self.Polycarbonate.set(product["Polycarbonate %"])
            self.Polyethylene.set(product["Polyethylene %"])         
            self.PolyurethaneFoam.set(product["Polyurethane foam %"])
            self.PrintedBoard.set(product["Printed wiring board, mixed mounted %"])
            self.PVCPipe.set(product["PVC pipe %"])
            self.PVC.set(product["PVC %"])
            self.Rubber.set(product["Rubber %"])
            self.Silicon.set(product["Silicon %"])
            self.StainlessSteel.set(product["Stainless steel %"])
            self.Steel.set(product["Steel (general or galvanised) %"])
            self.Zinc.set(product["Zinc %"])
            
            self.ABS1.set(product["ABS(new) %"])
            self.Aluminium1.set(product["Aluminium(new) %"])
            self.Brass1.set(product["Brass(new) %"])
            self.CastIron1.set(product["Cast iron(new) %"])
            self.Ceramic1.set(product["Ceramic(new) %"])
            self.Copper1.set(product["Copper(new) %"])
            self.ElectronicComponent1.set(product["Electronic component(new) %"])
            self.ExpandedPolystyrene1.set(product["Expanded polystyrene(new) %"])
            self.Glass1.set(product["Glass(new) %"])    
            self.Insulation1.set(product["Insulation (general)(new) %"])
            self.Iron1.set(product["Iron(new) %"])
            self.Lithium1.set(product["Lithium(new) %"])
            self.Plastics1.set(product["Plastics (general)(new) %"])
            self.Polyamide1.set(product["Polyamide(new) %"])
            self.Polycarbonate1.set(product["Polycarbonate(new) %"])
            self.Polyethylene1.set(product["Polyethylene(new) %"])         
            self.PolyurethaneFoam1.set(product["Polyurethane foam(new) %"])
            self.PrintedBoard1.set(product["Printed wiring board, mixed mounted(new) %"])
            self.PVCPipe1.set(product["PVC pipe(new) %"])
            self.PVC1.set(product["PVC(new) %"])
            self.Rubber1.set(product["Rubber(new) %"])
            self.Silicon1.set(product["Silicon(new) %"])
            self.StainlessSteel1.set(product["Stainless steel(new) %"])
            self.Steel1.set(product["Steel (general or galvanised)(new) %"])
            self.Zinc1.set(product["Zinc(new) %"])
            
            self.status.set('Query Successfully')
            self.qu_name.set("")
            
            self.create_tree()
        else:
            self.status.set('EPD does not exist')
        
    def create_tree(self):

        top1=tk.Tk()
        top1.title("EPD Page")
        
        A1 = float(self.TotalWeight.get())*(float(self.ABS.get())*3.76/100+float(self.Aluminium.get())*13.1/100
                                            +float(self.Brass.get())*4.8/100+float(self.CastIron.get())*1.52/100
                                            +float(self.Ceramic.get())*0.7/100+float(self.Copper.get())*3.81/100
                                            +float(self.ElectronicComponent.get())*49/100+float(self.ExpandedPolystyrene.get())*3.43/100
                                            +float(self.Glass.get())*1.44/100+float(self.Insulation.get())*1.86/100
                                            +float(self.Iron.get())*2.03/100+float(self.Lithium.get())*5.3/100
                                            +float(self.Plastics.get())*3.31/100+float(self.Polyamide.get())*9.14/100
                                            +float(self.Polycarbonate.get())*7.62/100+float(self.Polyethylene.get())*2.54/100
                                            +float(self.PolyurethaneFoam.get())*4.55/100+float(self.PrintedBoard.get())*154/100
                                            +float(self.PVCPipe.get())*3.23/100+float(self.PVC.get())*3.1/100
                                            +float(self.Rubber.get())*2.85/100+float(self.Silicon.get())*13.8/100
                                            +float(self.StainlessSteel.get())*4.4/100+float(self.Steel.get())*2.97/100
                                            +float(self.Zinc.get())*4.18/100)
        
        A2 = float(self.TotalWeight.get())/1000*((0.132*float(self.TransporttoFactoryr.get())+0.019*float(self.TransporttoFactorys.get()))
                                            *(float(self.ABS.get())/100+float(self.Aluminium.get())/100
                                            +float(self.Brass.get())/100+float(self.CastIron.get())/100
                                            +float(self.Ceramic.get())/100+float(self.Copper.get())/100
                                            +float(self.ElectronicComponent.get())/100+float(self.ExpandedPolystyrene.get())/100
                                            +float(self.Glass.get())/100+float(self.Insulation.get())/100
                                            +float(self.Iron.get())/100+float(self.Lithium.get())/100
                                            +float(self.Plastics.get())/100+float(self.Polyamide.get())/100
                                            +float(self.Polycarbonate.get())/100+float(self.Polyethylene.get())/100
                                            +float(self.PolyurethaneFoam.get())/100+float(self.PrintedBoard.get())/100
                                            +float(self.PVCPipe.get())/100+float(self.PVC.get())/100
                                            +float(self.Rubber.get())/100+float(self.Silicon.get())/100
                                            +float(self.StainlessSteel.get())/100+float(self.Steel.get())/100
                                            +float(self.Zinc.get())/100)+0.132*float(self.TransporttoFactoryl.get())
                                            *(float(self.ABS1.get())/100+float(self.Aluminium1.get())/100
                                            +float(self.Brass1.get())/100+float(self.CastIron1.get())/100
                                            +float(self.Ceramic1.get())/100+float(self.Copper1.get())/100
                                            +float(self.ElectronicComponent1.get())/100+float(self.ExpandedPolystyrene1.get())/100
                                            +float(self.Glass1.get())/100+float(self.Insulation1.get())/100
                                            +float(self.Iron1.get())/100+float(self.Lithium1.get())/100
                                            +float(self.Plastics1.get())/100+float(self.Polyamide1.get())/100
                                            +float(self.Polycarbonate1.get())/100+float(self.Polyethylene1.get())/100
                                            +float(self.PolyurethaneFoam1.get())/100+float(self.PrintedBoard1.get())/100
                                            +float(self.PVCPipe1.get())/100+float(self.PVC1.get())/100
                                            +float(self.Rubber1.get())/100+float(self.Silicon1.get())/100
                                            +float(self.StainlessSteel1.get())/100+float(self.Steel1.get())/100
                                            +float(self.Zinc1.get())/100-(float(self.ABS.get())/100+float(self.Aluminium.get())/100
                                            +float(self.Brass.get())/100+float(self.CastIron.get())/100
                                            +float(self.Ceramic.get())/100+float(self.Copper.get())/100
                                            +float(self.ElectronicComponent.get())/100+float(self.ExpandedPolystyrene.get())/100
                                            +float(self.Glass.get())/100+float(self.Insulation.get())/100
                                            +float(self.Iron.get())/100+float(self.Lithium.get())/100
                                            +float(self.Plastics.get())/100+float(self.Polyamide.get())/100
                                            +float(self.Polycarbonate.get())/100+float(self.Polyethylene.get())/100
                                            +float(self.PolyurethaneFoam.get())/100+float(self.PrintedBoard.get())/100
                                            +float(self.PVCPipe.get())/100+float(self.PVC.get())/100
                                            +float(self.Rubber.get())/100+float(self.Silicon.get())/100
                                            +float(self.StainlessSteel.get())/100+float(self.Steel.get())/100
                                            +float(self.Zinc.get())/100)))
        
        A3 = float(self.EnergyUse.get())*0.29
        A4 = float(self.TotalWeight.get())*float(self.TransporttoSite.get())/1000*0.132
        C2 = float(self.TotalWeight.get())*float(self.TransporttoLandfill.get())/1000*0.132
        C3 = float(self.EnergyUse.get())*0.29
        C4 =float(self.TotalWeight.get())*0.55*0.0089
        S1 = A1+A2+A3+A4+C2+C3+C4
        S2 = 1.3*S1
        R1 = S2
        B6 = float(self.PowerConsumption.get())/1000*float(self.ServiceLife.get())*10*365*0.29
        A1S = str(round(A1, 2))
        A2S = str(round(A2, 2))
        A3S = str(round(A3, 2))
        A4S = str(round(A4, 2))
        C2S = str(round(C2, 2))
        C3S = str(round(C3, 2))
        C4S = str(round(C4, 2))
        S1S = str(round(S1, 2))
        S2S = str(round(S2, 2))
        R1S = str(round(R1, 2))
        B6S = str(round(B6, 2))
        
        
        
        A1n = float(self.TotalWeight.get())*(float(self.ABS1.get())*3.76/100+float(self.Aluminium1.get())*13.1/100
                                            +float(self.Brass1.get())*4.8/100+float(self.CastIron1.get())*1.52/100
                                            +float(self.Ceramic1.get())*0.7/100+float(self.Copper1.get())*3.81/100
                                            +float(self.ElectronicComponent1.get())*49/100+float(self.ExpandedPolystyrene1.get())*3.43/100
                                            +float(self.Glass1.get())*1.44/100+float(self.Insulation1.get())*1.86/100
                                            +float(self.Iron1.get())*2.03/100+float(self.Lithium1.get())*5.3/100
                                            +float(self.Plastics1.get())*3.31/100+float(self.Polyamide1.get())*9.14/100
                                            +float(self.Polycarbonate1.get())*7.62/100+float(self.Polyethylene1.get())*2.54/100
                                            +float(self.PolyurethaneFoam1.get())*4.55/100+float(self.PrintedBoard1.get())*154/100
                                            +float(self.PVCPipe1.get())*3.23/100+float(self.PVC1.get())*3.1/100
                                            +float(self.Rubber1.get())*2.85/100+float(self.Silicon1.get())*13.8/100
                                            +float(self.StainlessSteel1.get())*4.4/100+float(self.Steel1.get())*2.97/100
                                            +float(self.Zinc1.get())*4.18/100)
        
        A2n = float(self.TotalWeight.get())*(0.132*float(self.TransporttoFactoryr.get())+0.019*float(self.TransporttoFactorys.get()))/1000
        A3n = float(self.EnergyUse.get())*0.84
        A4n = float(self.TotalWeight.get())*float(self.TransporttoSite.get())/1000*0.132
        C2n = float(self.TotalWeight.get())*float(self.TransporttoLandfill.get())/1000*0.132
        C3n = float(self.EnergyUse.get())*0.29
        C4n =float(self.TotalWeight.get())*0.55*0.0089
        S1n = A1n+A2n+A3n+A4n+C2n+C3n+C4n
        S2n = 1.3*S1n
        R1n = S2n
        A1nS = str(round(A1n, 2))
        A2nS = str(round(A2n, 2))
        A3nS = str(round(A3n, 2))
        A4nS = str(round(A4n, 2))
        C2nS = str(round(C2n, 2))
        C3nS = str(round(C3n, 2))
        C4nS = str(round(C4n, 2))
        S1nS = str(round(S1n, 2))
        S2nS = str(round(S2n, 2))
        R1nS = str(round(R1n, 2))
        MD = (float(self.ABS1.get())/100+float(self.Aluminium1.get())/100+float(self.Brass1.get())/100+
              float(self.CastIron1.get())/100+float(self.Ceramic1.get())/100+float(self.Copper1.get())/100+
              float(self.ElectronicComponent1.get())/100+float(self.ExpandedPolystyrene1.get())/100+
              float(self.Glass1.get())/100+float(self.Insulation1.get())/100+float(self.Iron1.get())/100+
              float(self.Lithium1.get())/100 +float(self.Plastics1.get())/100+float(self.Polyamide1.get())/100+
              float(self.Polycarbonate1.get())/100+float(self.Polyethylene1.get())/100+float(self.PolyurethaneFoam1.get())/100+
              float(self.PrintedBoard1.get())/100+float(self.PVCPipe1.get())/100+float(self.PVC1.get())/100+
              float(self.Rubber1.get())/100+float(self.Silicon1.get())/100+float(self.StainlessSteel1.get())/100+
              float(self.Steel1.get())/100+float(self.Zinc1.get())/100)
        if MD < 0.95:
            MD = "N"
        else:
            MD = "Y"
             
        
        


        tk.Label(top1, text="Mid-level calculation",  justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300, font=('microsoft yahei', 8, 'bold'), bg = "LightSteelBlue").grid(row=1, column=0)        
        tk.Label(top1, text="Remanufactured Product", width=25, borderwidth=1, relief="groove", wraplength = 180,font=('microsoft yahei', 8, 'bold'), bg = "LightSteelBlue").grid(row=1, column=1)
        tk.Label(top1, text="New Product", width=25, borderwidth=1, relief="groove", wraplength = 180, font=('microsoft yahei', 8, 'bold'), bg = "LightSteelBlue").grid(row=1, column=2)
        tk.Label(top1, text="Notes/source",  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450, font=('microsoft yahei', 8, 'bold'), bg = "LightSteelBlue").grid(row=1, column=3)

        tk.Label(top1, text="Date of assessment", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=2, column=0)        
        tk.Label(top1, text="", width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=2, column=1)
        tk.Label(top1, text="", width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=2, column=2)
        tk.Label(top1, text="", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=2, column=3)
        
        tk.Label(top1, text="Name of assessor and assessor organisation", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=3, column=0)        
        tk.Label(top1, text="EGG Lighting", width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=3, column=1)
        tk.Label(top1, text="", width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=3, column=2)
        tk.Label(top1, text="", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=3, column=3)
        
        tk.Label(top1, text="Contact details of assessor", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=4, column=0)        
        tk.Label(top1, text="circular@egglighting.com", width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=4, column=1)
        tk.Label(top1, text="", width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=4, column=2)
        tk.Label(top1, text="", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=4, column=3)
        
        tk.Label(top1, text="Product information", anchor = 'w', justify = 'left', width=155, borderwidth=1, relief="groove", wraplength = 300 ,font=('microsoft yahei', 8, 'bold'), bg = "LightSteelBlue").grid(row=5,  column=0, columnspan=4)  
        
        tk.Label(top1, text="Product name", height=3, anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=6, column=0)        
        tk.Label(top1, text=self.ProductName.get(), height=3, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=6, column=1)
        tk.Label(top1, text="", height=3, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=6, column=2)
        tk.Label(top1, text="ASD Luminaire (Edinburgh)", anchor = 'w',  justify = 'left', height=3, width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=6, column=3)
        
        tk.Label(top1, text="Power Consumption", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=7, column=0)        
        tk.Label(top1, text=self.PowerConsumption.get(), width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=7, column=1)
        tk.Label(top1, text="", width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=7, column=2)
        tk.Label(top1, text="W", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=7, column=3)
        
        tk.Label(top1, text="Product weight (kg)", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=8, column=0)        
        tk.Label(top1, text=self.TotalWeight.get(), width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=8, column=1)
        tk.Label(top1, text="", width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=8, column=2)
        tk.Label(top1, text="kg", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=8, column=3)
        
        tk.Label(top1, text="Material % breakdown for at least 95% of the product weight? (Y/N)", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=9, column=0)        
        tk.Label(top1, text=MD, width=25, height=2, borderwidth=1, relief="groove", wraplength = 180).grid(row=9, column=1)
        tk.Label(top1, text="", width=25, height=2, borderwidth=1, relief="groove", wraplength = 180).grid(row=9, column=2)
        tk.Label(top1, text="97%", anchor = 'w',  justify = 'left', width=64, height=2, borderwidth=1, relief="groove", wraplength = 450).grid(row=9, column=3)
        
        tk.Label(top1, text="Service life of the product (years)", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=10, column=0)        
        tk.Label(top1, text=self.ServiceLife.get(), width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=10, column=1)
        tk.Label(top1, text="", width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=10, column=2)
        tk.Label(top1, text="Life depends on usage", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=10, column=3)

        tk.Label(top1, text="Energy consumption of the factory per unit of product", height=3, anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=11, column=0)        
        tk.Label(top1, text=self.EnergyUse.get(), height=3, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=11, column=1)
        tk.Label(top1, text="", height=3, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=11, column=2)
        tk.Label(top1, text="", height=3, anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=11, column=3)
        
        tk.Label(top1, text="Location of manufacture", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=12, column=0)        
        tk.Label(top1, text="Glasgow, Scotland", width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=12, column=1)
        tk.Label(top1, text="", width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=12, column=2)
        tk.Label(top1, text="", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=12, column=3)
        
        tk.Label(top1, text="Embodied carbon results (kg CO2e) breakdown", anchor = 'w', justify = 'left', width=155, borderwidth=1, relief="groove", wraplength = 300, font=('microsoft yahei', 8, 'bold'), bg = "LightSteelBlue").grid(row=13, column=0, columnspan=4) 
        
        tk.Label(top1, text="A1: Material extraction", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=14, column=0)        
        tk.Label(top1, text=A1S, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=14, column=1)
        tk.Label(top1, text=A1nS, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=14, column=2)
        tk.Label(top1, text="", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=14, column=3) 
        
        tk.Label(top1, text="A2: Transport to factory", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=15, column=0)        
        tk.Label(top1, text=A2S, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=15, column=1)
        tk.Label(top1, text=A2nS, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=15, column=2)
        tk.Label(top1, text="", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=15, column=3)        
        
        tk.Label(top1, text="A3: Manufacturing", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=16, column=0)        
        tk.Label(top1, text=A3S, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=16, column=1)
        tk.Label(top1, text=A3nS, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=16, column=2)
        tk.Label(top1, text="", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=16, column=3)        
        
        tk.Label(top1, text="A4: Transport to site", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove" , wraplength = 300).grid(row=17, column=0)        
        tk.Label(top1, text=A4S, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=17, column=1)
        tk.Label(top1, text=A4nS, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=17, column=2)
        tk.Label(top1, text="", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=17, column=3)        

        tk.Label(top1, text="A5: Construction", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove" , wraplength = 300, bg = "DarkGray").grid(row=18, column=0)        
        tk.Label(top1, text="0.00", width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=18, column=1)
        tk.Label(top1, text="", width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=18, column=2)
        tk.Label(top1, text="Assume no installation emissions", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450, bg = "DarkGray").grid(row=18, column=3)                

        tk.Label(top1, text="B1: Use - refrigerant leakage", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove" , wraplength = 300, bg = "DarkGray").grid(row=19, column=0)        
        tk.Label(top1, text="-", width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=19, column=1)
        tk.Label(top1, text="", width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=19, column=2)
        tk.Label(top1, text="Excluded as considered negligible", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450, bg = "DarkGray").grid(row=19, column=3) 
        
        tk.Label(top1, text="B2: Maintenance", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove" , wraplength = 300, bg = "DarkGray").grid(row=20, column=0)        
        tk.Label(top1, text="-", width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=20, column=1)
        tk.Label(top1, text="", width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=20, column=2)
        tk.Label(top1, text="Not applicable", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450, bg = "DarkGray").grid(row=20, column=3)        
        
        tk.Label(top1, text="B3: Repair", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove" , wraplength = 300, bg = "DarkGray").grid(row=21, column=0)        
        tk.Label(top1, text="-", width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=21, column=1)
        tk.Label(top1, text="", width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=21, column=2)
        tk.Label(top1, text="Commercial LED luminaires not typically suitable for repair", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450, bg = "DarkGray").grid(row=21, column=3)        
        
        tk.Label(top1, text="B4: Replacement", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove" , wraplength = 300, bg = "DarkGray").grid(row=22, column=0)        
        tk.Label(top1, text="-", width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=22, column=1)
        tk.Label(top1, text="", width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=22, column=2)
        tk.Label(top1, text="Only applicable during \"Service Life\". Remanufacture is a stage D operation", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450, bg = "DarkGray").grid(row=22, column=3)        
        
        tk.Label(top1, text="B5: Refurbishment", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove" , wraplength = 300, bg = "DarkGray").grid(row=23, column=0)        
        tk.Label(top1, text="-", width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=23, column=1)
        tk.Label(top1, text="", width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=23, column=2)
        tk.Label(top1, text="Only applicable during \"Service Life\". Remanufacture is a stage D operation", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450, bg = "DarkGray").grid(row=23, column=3)        
        
        tk.Label(top1, text="B6: Operational energy", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove" , wraplength = 300, bg = "DarkGray").grid(row=24, column=0)        
        tk.Label(top1, text=B6S, width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=24, column=1)
        tk.Label(top1, text="", width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=24, column=2)
        tk.Label(top1, text="Outside of scope for embodied carbon", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450, bg = "DarkGray").grid(row=24, column=3)       
        
        tk.Label(top1, text="B7: Operational water", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove" , wraplength = 300, bg = "DarkGray").grid(row=25, column=0)        
        tk.Label(top1, text="-", width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=25, column=1)
        tk.Label(top1, text="", width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=25, column=2)
        tk.Label(top1, text="Negligible water used", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450, bg = "DarkGray").grid(row=25, column=3)        
        
        tk.Label(top1, text="C1: Deconstruction including refridgerant leakage", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove" , wraplength = 300, bg = "DarkGray").grid(row=26, column=0)        
        tk.Label(top1, text="-", width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=26, column=1)
        tk.Label(top1, text="", width=25, borderwidth=1, relief="groove", wraplength = 180, bg = "DarkGray").grid(row=26, column=2)
        tk.Label(top1, text="Excluded as considered negligible", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450, bg = "DarkGray").grid(row=26, column=3)                 

        tk.Label(top1, text="C2: Transport to waste disposal", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=27, column=0)        
        tk.Label(top1, text=C2S, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=27, column=1)
        tk.Label(top1, text=C2nS, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=27, column=2)
        tk.Label(top1, text="Assuming transport to national waste processing", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=27, column=3) 
        
        tk.Label(top1, text="C3: Waste processing energy use", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=28, column=0)        
        tk.Label(top1, text=C3S, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=28, column=1,)
        tk.Label(top1, text=C3nS, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=28, column=2)
        tk.Label(top1, text="Assuming same energy use as final assembly", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=28, column=3)
        
        tk.Label(top1, text="C4: Disposal to landfill", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=29, column=0)     
        tk.Label(top1, text=C4S, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=29, column=1)
        tk.Label(top1, text=C4nS, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=29, column=2)
        tk.Label(top1, text="Assume 55% sent to landfill as per CIBSE TM65 assumption", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=29, column=3)
        
        tk.Label(top1, text="Embodied carbon results (kg CO2e) — without refrigerant leakage", anchor = 'w', justify = 'left', width=155, borderwidth=1, relief="groove", wraplength = 500, font=('microsoft yahei', 8, 'bold'), bg = "LightSteelBlue").grid(row=30, column=0, columnspan=4) 
        
        tk.Label(top1, text="A1–C4 (excluding A5-C1)", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=31, column=0)        
        tk.Label(top1, text=S1S, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=31, column=1)
        tk.Label(top1, text=S1nS, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=31, column=2) 
        tk.Label(top1, text="", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=31, column=3)
        
        tk.Label(top1, text="A1–C4 (excluding A5-C1) with buffer factor", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=32, column=0)        
        tk.Label(top1, text=S2S, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=32, column=1)
        tk.Label(top1, text=S2nS, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=32, column=2)
        tk.Label(top1, text="Assuming 1.3 buffer", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=32, column=3)
        
        tk.Label(top1, text="Embodied carbon result with ‘mid-level calculation’ method (kg CO2e) — total", anchor = 'w', justify = 'left', width=155, borderwidth=1, relief="groove", wraplength = 500, font=('microsoft yahei', 8, 'bold'), bg = "LightSteelBlue").grid(row=33, column=0, columnspan=4) 
        
        tk.Label(top1, text="Result of ‘mid-level’ calculation", anchor = 'w', justify = 'left', width=40, borderwidth=1, relief="groove", wraplength = 300).grid(row=34, column=0)        
        tk.Label(top1, text=R1S, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=34, column=1)
        tk.Label(top1, text=R1nS, width=25, borderwidth=1, relief="groove", wraplength = 180).grid(row=34, column=2) 
        tk.Label(top1, text="", anchor = 'w',  justify = 'left', width=64, borderwidth=1, relief="groove", wraplength = 450).grid(row=34, column=3)
        
        tk.Label(top1).grid(row=35)
        
        tk.Button(top1, text='Download PDF', command = lambda:self.saveepd()).grid(row=36, column=1) 
        
        tk.Label(top1).grid(row=37)

        top1.mainloop()
    def saveepd(self):
        A1 = float(self.TotalWeight.get())*(float(self.ABS.get())*3.76/100+float(self.Aluminium.get())*13.1/100
                                            +float(self.Brass.get())*4.8/100+float(self.CastIron.get())*1.52/100
                                            +float(self.Ceramic.get())*0.7/100+float(self.Copper.get())*3.81/100
                                            +float(self.ElectronicComponent.get())*49/100+float(self.ExpandedPolystyrene.get())*3.43/100
                                            +float(self.Glass.get())*1.44/100+float(self.Insulation.get())*1.86/100
                                            +float(self.Iron.get())*2.03/100+float(self.Lithium.get())*5.3/100
                                            +float(self.Plastics.get())*3.31/100+float(self.Polyamide.get())*9.14/100
                                            +float(self.Polycarbonate.get())*7.62/100+float(self.Polyethylene.get())*2.54/100
                                            +float(self.PolyurethaneFoam.get())*4.55/100+float(self.PrintedBoard.get())*154/100
                                            +float(self.PVCPipe.get())*3.23/100+float(self.PVC.get())*3.1/100
                                            +float(self.Rubber.get())*2.85/100+float(self.Silicon.get())*13.8/100
                                            +float(self.StainlessSteel.get())*4.4/100+float(self.Steel.get())*2.97/100
                                            +float(self.Zinc.get())*4.18/100)
        
        A2 = float(self.TotalWeight.get())/1000*((0.132*float(self.TransporttoFactoryr.get())+0.019*float(self.TransporttoFactorys.get()))
                                            *(float(self.ABS.get())/100+float(self.Aluminium.get())/100
                                            +float(self.Brass.get())/100+float(self.CastIron.get())/100
                                            +float(self.Ceramic.get())/100+float(self.Copper.get())/100
                                            +float(self.ElectronicComponent.get())/100+float(self.ExpandedPolystyrene.get())/100
                                            +float(self.Glass.get())/100+float(self.Insulation.get())/100
                                            +float(self.Iron.get())/100+float(self.Lithium.get())/100
                                            +float(self.Plastics.get())/100+float(self.Polyamide.get())/100
                                            +float(self.Polycarbonate.get())/100+float(self.Polyethylene.get())/100
                                            +float(self.PolyurethaneFoam.get())/100+float(self.PrintedBoard.get())/100
                                            +float(self.PVCPipe.get())/100+float(self.PVC.get())/100
                                            +float(self.Rubber.get())/100+float(self.Silicon.get())/100
                                            +float(self.StainlessSteel.get())/100+float(self.Steel.get())/100
                                            +float(self.Zinc.get())/100)+0.132*float(self.TransporttoFactoryl.get())
                                            *(float(self.ABS1.get())/100+float(self.Aluminium1.get())/100
                                            +float(self.Brass1.get())/100+float(self.CastIron1.get())/100
                                            +float(self.Ceramic1.get())/100+float(self.Copper1.get())/100
                                            +float(self.ElectronicComponent1.get())/100+float(self.ExpandedPolystyrene1.get())/100
                                            +float(self.Glass1.get())/100+float(self.Insulation1.get())/100
                                            +float(self.Iron1.get())/100+float(self.Lithium1.get())/100
                                            +float(self.Plastics1.get())/100+float(self.Polyamide1.get())/100
                                            +float(self.Polycarbonate1.get())/100+float(self.Polyethylene1.get())/100
                                            +float(self.PolyurethaneFoam1.get())/100+float(self.PrintedBoard1.get())/100
                                            +float(self.PVCPipe1.get())/100+float(self.PVC1.get())/100
                                            +float(self.Rubber1.get())/100+float(self.Silicon1.get())/100
                                            +float(self.StainlessSteel1.get())/100+float(self.Steel1.get())/100
                                            +float(self.Zinc1.get())/100-(float(self.ABS.get())/100+float(self.Aluminium.get())/100
                                            +float(self.Brass.get())/100+float(self.CastIron.get())/100
                                            +float(self.Ceramic.get())/100+float(self.Copper.get())/100
                                            +float(self.ElectronicComponent.get())/100+float(self.ExpandedPolystyrene.get())/100
                                            +float(self.Glass.get())/100+float(self.Insulation.get())/100
                                            +float(self.Iron.get())/100+float(self.Lithium.get())/100
                                            +float(self.Plastics.get())/100+float(self.Polyamide.get())/100
                                            +float(self.Polycarbonate.get())/100+float(self.Polyethylene.get())/100
                                            +float(self.PolyurethaneFoam.get())/100+float(self.PrintedBoard.get())/100
                                            +float(self.PVCPipe.get())/100+float(self.PVC.get())/100
                                            +float(self.Rubber.get())/100+float(self.Silicon.get())/100
                                            +float(self.StainlessSteel.get())/100+float(self.Steel.get())/100
                                            +float(self.Zinc.get())/100)))
        
        A3 = float(self.EnergyUse.get())*0.29
        A4 = float(self.TotalWeight.get())*float(self.TransporttoSite.get())/1000*0.132
        C2 = float(self.TotalWeight.get())*float(self.TransporttoLandfill.get())/1000*0.132
        C3 = float(self.EnergyUse.get())*0.29
        C4 =float(self.TotalWeight.get())*0.55*0.0089
        S1 = A1+A2+A3+A4+C2+C3+C4
        S2 = 1.3*S1
        R1 = S2
        A1S = str(round(A1, 2))
        A2S = str(round(A2, 2))
        A3S = str(round(A3, 2))
        A4S = str(round(A4, 2))
        C2S = str(round(C2, 2))
        C3S = str(round(C3, 2))
        C4S = str(round(C4, 2))
        S1S = str(round(S1, 2))
        S2S = str(round(S2, 2))
        R1S = str(round(R1, 2))
        
        
        A1n = float(self.TotalWeight.get())*(float(self.ABS1.get())*3.76/100+float(self.Aluminium1.get())*13.1/100
                                            +float(self.Brass1.get())*4.8/100+float(self.CastIron1.get())*1.52/100
                                            +float(self.Ceramic1.get())*0.7/100+float(self.Copper1.get())*3.81/100
                                            +float(self.ElectronicComponent1.get())*49/100+float(self.ExpandedPolystyrene1.get())*3.43/100
                                            +float(self.Glass1.get())*1.44/100+float(self.Insulation1.get())*1.86/100
                                            +float(self.Iron1.get())*2.03/100+float(self.Lithium1.get())*5.3/100
                                            +float(self.Plastics1.get())*3.31/100+float(self.Polyamide1.get())*9.14/100
                                            +float(self.Polycarbonate1.get())*7.62/100+float(self.Polyethylene1.get())*2.54/100
                                            +float(self.PolyurethaneFoam1.get())*4.55/100+float(self.PrintedBoard1.get())*154/100
                                            +float(self.PVCPipe1.get())*3.23/100+float(self.PVC1.get())*3.1/100
                                            +float(self.Rubber1.get())*2.85/100+float(self.Silicon1.get())*13.8/100
                                            +float(self.StainlessSteel1.get())*4.4/100+float(self.Steel1.get())*2.97/100
                                            +float(self.Zinc1.get())*4.18/100)
        
        A2n = float(self.TotalWeight.get())*(0.132*float(self.TransporttoFactoryr.get())+0.019*float(self.TransporttoFactorys.get()))/1000
        A3n = float(self.EnergyUse.get())*0.84
        A4n = float(self.TotalWeight.get())*float(self.TransporttoSite.get())/1000*0.132
        C2n = float(self.TotalWeight.get())*float(self.TransporttoLandfill.get())/1000*0.132
        C3n = float(self.EnergyUse.get())*0.29
        C4n =float(self.TotalWeight.get())*0.55*0.0089
        S1n = A1n+A2n+A3n+A4n+C2n+C3n+C4n
        S2n = 1.3*S1n
        R1n = S2n
        
        
        A1nS = str(round(A1n, 2))
        A2nS = str(round(A2n, 2))
        A3nS = str(round(A3n, 2))
        A4nS = str(round(A4n, 2))
        C2nS = str(round(C2n, 2))
        C3nS = str(round(C3n, 2))
        C4nS = str(round(C4n, 2))
        S1nS = str(round(S1n, 2))
        S2nS = str(round(S2n, 2))
        R1nS = str(round(R1n, 2))
        MD = (float(self.ABS1.get())/100+float(self.Aluminium1.get())/100+float(self.Brass1.get())/100+
              float(self.CastIron1.get())/100+float(self.Ceramic1.get())/100+float(self.Copper1.get())/100+
              float(self.ElectronicComponent1.get())/100+float(self.ExpandedPolystyrene1.get())/100+
              float(self.Glass1.get())/100+float(self.Insulation1.get())/100+float(self.Iron1.get())/100+
              float(self.Lithium1.get())/100 +float(self.Plastics1.get())/100+float(self.Polyamide1.get())/100+
              float(self.Polycarbonate1.get())/100+float(self.Polyethylene1.get())/100+float(self.PolyurethaneFoam1.get())/100+
              float(self.PrintedBoard1.get())/100+float(self.PVCPipe1.get())/100+float(self.PVC1.get())/100+
              float(self.Rubber1.get())/100+float(self.Silicon1.get())/100+float(self.StainlessSteel1.get())/100+
              float(self.Steel1.get())/100+float(self.Zinc1.get())/100)
        if MD < 0.95:
            MD = "N"
        else:
            MD = "Y"
                                 
        doc = SimpleDocTemplate("EPD.pdf")
        styles = getSampleStyleSheet()
        style = styles['Normal']

        story =[]

        P1 = Image("D:\\Onedrive\\OneDrive - University of Edinburgh\\Desktop\\论文\\v2\\v2\\logo1.png")
        P2 = Image("D:\\Onedrive\\OneDrive - University of Edinburgh\\Desktop\\论文\\v2\\v2\\logo2.png")
        P1.drawHeight = 70
        P1.drawWidth = 148
        P2.drawHeight = 70
        P2.drawWidth = 68

        data1 = [[P1, P2, "Embodied Carbon Calculation (TM65: 2020)"],
                 ["", "", "This calculation has been carried out by EGG Lighting according to the CIBSE \nTM65: 2020\"Embodied carbon in building services: a calculation methodology\""],
                 ]
        t1 = Table(data1,colWidths=None, rowHeights=None,style=[
        ('SPAN',(0,0),(0,1)),
        ('SPAN',(1,0),(1,1)),
        ("VALIGN",(2,0),(2,0),"TOP"),
        ('FONTSIZE',(2,0),(2,0), 14),
        ('FONTSIZE',(2,1),(2,1), 8),
        ('FONTWEIGHT',(2,0),(2,0), "bold"),
        ])
        
        data2 = [["For any questions about this document or product please contact: circular@egglighting.com"]]
        t2 = Table(data2,colWidths=522, rowHeights=None, spaceAfter=2)
        
        data3 = [
                 ["PRODUCT SPEC", "", ""],
                 ["Product name", self.ProductName.get(), ""],
                 ["Power consumption (W)", self.PowerConsumption.get(), ""],
                 ["Service life of the product (years)", self.ServiceLife.get(), ""]
                ]
        t3 = Table(data3,colWidths=174, rowHeights=None, spaceAfter=2, style=[
        ('FONTSIZE',(0,0),(0,0), 12),
        ('FONTSIZE',(0,1),(0,3), 8),
        ('FONTSIZE',(1,1),(1,3), 8),
        ('BACKGROUND',(0,1),(2,3), colors.lavender),
        ])
        
        data4 = [["RESULTS SUMMARY", ""],
                 ["Total embodied carbon - remanufactured luminaire (kgCO2e)", R1S],
                 ["Total embodied carbon - equivalent new luminaire (kgCO2e)", R1nS]]

        t4 = Table(data4,colWidths=261, rowHeights=None, spaceAfter=4 ,style=[
        ('FONTSIZE',(0,0),(0,0), 12),
        ('FONTSIZE',(0,1),(0,2), 8),
        ('FONTSIZE',(1,1),(1,2), 8),
        ('BACKGROUND',(0,1),(1,2), colors.lavender),
        ])
        
        data5 = [["ASSESSMENT: REMANUFACTURED"]]
        t5 = Table(data5,colWidths=522, rowHeights=None,style=[
        ('FONTSIZE',(0,0),(0,0), 12),
        ('ALIGNMENT',(0,0),(0,0), "CENTER"),
        ])
    

        d1 = Drawing(260, 260)
        pc = Pie()
        pc.x = 129
        pc.y = 20
        pc.width = 180
        pc.height = 180
        pc.sideLabels = 1
        pc.slices.fontSize = 9
        pc.simpleLabels = 0
        pc.data = [A1n,A2n,A3n,A4n,C2n,C3n,C4n]
        pc.labels = ['A1: Material extraction','A2: Transport to factory','A3: Manufacturing',
                     'A4: Transport to site','C2: Transport to waste disposal',
                     'C3: Waste processing energy use',"C4: Disposal to landfill"]


        pc.slices[1].label_angle = 90
        pc.slices[2].label_angle = 90
        pc.slices[3].label_angle = 90
        pc.slices[4].label_angle = 90
        pc.slices[5].label_angle = 90
        pc.slices[6].label_angle = 90
        pc.slices[1].label_dx = 10
        pc.slices[2].label_dx = 20
        pc.slices[3].label_dx = 30
        pc.slices[4].label_dx = 40
        pc.slices[5].label_dx = 50
        pc.slices[6].label_dx = 60

        d1.add(pc)
        
        data6 = [["ASSESSMENT: REMANUFACTURED VS NEW EQUIVALENT"]]
        t6 = Table(data6,colWidths=522, rowHeights=None,style=[
        ('FONTSIZE',(0,0),(0,0), 12),
        ('ALIGNMENT',(0,0),(0,0), "CENTER"),
        ])
        
        
        d2 = Drawing(400, 150)
        data = [
                (A1n,A2n,A3n,A4n,C2n,C3n,C4n,S1n,S2n),
                ((A1,A2,A3,A4,C2,C3,C4,S1,S2))
               ]
        bc = VerticalBarChart()
        bc.x = 80
        bc.y = 10
        bc.height = 125
        bc.width = 300
        bc.data = data
        bc.strokeColor = colors.black
        bc.valueAxis.valueMin = 0
        bc.valueAxis.valueMax = 120
        bc.valueAxis.valueStep = 10
        bc.categoryAxis.labels.boxAnchor = 'ne'
        bc.categoryAxis.labels.dx = 8
        bc.categoryAxis.labels.dy = -2
        bc.categoryAxis.labels.angle = 0
        bc.categoryAxis.categoryNames = ['A1','A2','A3','A4','C2','C3','C4','A1-C4',"Total"]
        d2.add(bc)
        

        a = self.ProductName.get()[0:56]

        
        



        data= [["Mid-level calculation","", "Remanufactured Product", "New Product","",""],
        ["Name of assessor and assessor organisation","", "EGG Lighting", "EGG Lighting","",""],
        ["Contact details of assessor", "","circular@egglighting.com", "circular@egglighting.com","",""],
        ["Product information","", "", "","",""],
        ["Product name","", a, "","ASD Luminaire (Edinburgh)",""],
        ["Power Consumption", "",self.PowerConsumption.get(), self.PowerConsumption.get(),"W",""],
        ["Product weight (kg)","", self.TotalWeight.get(), self.TotalWeight.get(),"kg",""],
        ["Material % breakdown for at least 95% of the \nproduct weight? (Y/N)","", MD, MD,"97%",""],
        ["Service life of the product (years)","", self.ServiceLife.get(), self.ServiceLife.get(),"Life depends on usage",""],
        ["Location of manufacture","", "Glasgow, Scotland", "Glasgow, Scotland","",""],
        ["Embodied carbon results (kg CO2e) breakdown","", '', '',"",""],
        ["A1: Material extraction","", A1S, A1nS,"",""],
        ["A2: Transport to factory","", A2S, A2nS,"",""],
        ["A3: Manufacturing","", A3S, A3nS,"",""],
        ["A4: Transport to site","", A4S, A4nS,"",""],
        ["C2: Transport to waste disposal","", C2S, C2nS,"Assuming transport to national waste processing",""],
        ["C3: Waste processing energy use","", C3S, C3nS,"Assuming same energy use as final assembly",""],
        ["C4: Disposal to landfill", "",C4S, C4nS,"Assume 55% sent to landfill as per CIBSE TM65 \nassumption",""],
        ["Embodied carbon results (kg CO2e) — without refrigerant leakage","", "", "","",""],
        ["A1–C4 (excluding A5-C1)","", S1S, S1nS,"",""],
        ["A1–C4 (excluding A5-C1) with buffer factor","", S2S, S2nS,"Assuming 1.3 buffer - scaleup factor is assumed to \nbe 1.3 in the mid level calculation",""],
        ["Embodied carbon result with ‘mid-level calculation’ method (kg CO2e) — total","", "", "","",""],
        ["Result of ‘mid-level’ calculation", "",R1S, R1nS,"",""]]
        
        t=Table(data,colWidths=90, rowHeights=None, spaceBefore=10,style=[
        ('GRID',(0,0),(-1,-1),1,colors.grey),
        ('BACKGROUND', (0, 0), (5, 0), colors.lavender),
        ('BACKGROUND', (0, 3), (5, 3), colors.lavender),
        ('BACKGROUND', (0, 10), (5, 10), colors.lavender),
        ('BACKGROUND', (0, 18), (5, 18), colors.lavender),
        ('BACKGROUND', (0, 21), (5, 21), colors.lavender),
        ('SPAN',(0,3),(5,3)),
        ('SPAN',(0,10),(5,10)),
        ('SPAN',(0,18),(5,18)),
        ('SPAN',(0,21),(5,21)),
        ('SPAN',(0,0),(1,0)),
        ('SPAN',(0,1),(1,1)),
        ('SPAN',(0,2),(1,2)),
        ('SPAN',(0,4),(1,4)),
        ('SPAN',(0,5),(1,5)),
        ('SPAN',(0,6),(1,6)),
        ('SPAN',(0,7),(1,7)),
        ('SPAN',(0,8),(1,8)),
        ('SPAN',(0,9),(1,9)),
        ('SPAN',(0,11),(1,11)),
        ('SPAN',(0,12),(1,12)),
        ('SPAN',(0,13),(1,13)),
        ('SPAN',(0,14),(1,14)),
        ('SPAN',(0,15),(1,15)),
        ('SPAN',(0,16),(1,16)),
        ('SPAN',(0,17),(1,17)),
        ('SPAN',(0,19),(1,19)),
        ('SPAN',(0,20),(1,20)),
        ('SPAN',(0,22),(1,22)),
        ('SPAN',(4,0),(5,0)),
        ('SPAN',(4,1),(5,1)),
        ('SPAN',(4,2),(5,2)),
        ('SPAN',(4,4),(5,4)),
        ('SPAN',(4,5),(5,5)),
        ('SPAN',(4,6),(5,6)),
        ('SPAN',(4,7),(5,7)),
        ('SPAN',(4,8),(5,8)),
        ('SPAN',(4,9),(5,9)),
        ('SPAN',(4,11),(5,11)),
        ('SPAN',(4,12),(5,12)),
        ('SPAN',(4,13),(5,13)),
        ('SPAN',(4,14),(5,14)),
        ('SPAN',(4,15),(5,15)),
        ('SPAN',(4,16),(5,16)),
        ('SPAN',(4,17),(5,17)),
        ('SPAN',(4,19),(5,19)),
        ('SPAN',(4,20),(5,20)),
        ('SPAN',(4,22),(5,22)),
        ('SPAN',(2,4),(3,4)),
        ('ALIGN',(2,0),(3,22),'CENTER'),
        ('FONTSIZE',(0,0),(5,22), 7.5),
        ])
        
        story.append(t1)
        story.append(t2)
        story.append(t3)
        story.append(t4)
        story.append(t5)
        story.append(d1) 
        story.append(t6)
        story.append(d2)
        story.append(t)
        doc.build(story)


class DeleteFrame(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        tk.Label(self, text='Delete Product').pack()
        self.status = tk.StringVar()
        self.de_name = tk.StringVar()
        self.ProductName = tk.StringVar()
        self.create_page()


    def create_page(self):
        tk.Label(self, text="Delete by Name").pack(anchor=tk.W, padx=20)
        e1 = tk.Entry(self, textvariable=self.de_name)
        e1.pack(side=tk.LEFT, padx=20, pady=5)

        tk.Button(self, text='Delete', command=self.delete).pack(side=tk.RIGHT)
        tk.Label(self, textvariable=self.status).pack()


    def delete(self):
        name = self.de_name.get()
        print(name)
        result = db.delete_by_name(name)
        if result:
            db.save_data()
            self.status.set('Delete Successfully')
            self.de_name.set("")
        else:
            self.status.set('Product does not exist')


class ChangeFrame(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.root = master


        self.ProductName = tk.StringVar()
        self.ProductNumber = tk.StringVar()
        self.TotalWeight = tk.StringVar()
        self.EnergyUse = tk.StringVar()
        self.PowerConsumption = tk.StringVar()
        self.ServiceLife = tk.StringVar()
        self.TransporttoFactoryl = tk.StringVar()
        self.TransporttoFactoryr = tk.StringVar()
        self.TransporttoFactorys = tk.StringVar()        
        self.TransporttoSite = tk.StringVar()
        self.TransporttoLandfill = tk.StringVar()
        
        self.ABS = tk.StringVar()
        self.Aluminium = tk.StringVar()
        self.Brass = tk.StringVar()
        self.CastIron = tk.StringVar()
        self.Ceramic = tk.StringVar()
        self.Copper = tk.StringVar()
        self.ElectronicComponent = tk.StringVar()
        self.ExpandedPolystyrene = tk.StringVar()
        self.Glass = tk.StringVar()        
        self.Insulation = tk.StringVar()
        self.Iron = tk.StringVar()
        self.Lithium = tk.StringVar()
        self.Plastics = tk.StringVar()
        self.Polyamide = tk.StringVar()
        self.Polycarbonate = tk.StringVar()
        self.Polyethylene  = tk.StringVar()          
        self.PolyurethaneFoam = tk.StringVar()
        self.PrintedBoard = tk.StringVar()
        self.PVCPipe = tk.StringVar()
        self.PVC = tk.StringVar()
        self.Rubber  = tk.StringVar()
        self.Silicon = tk.StringVar()
        self.StainlessSteel = tk.StringVar()
        self.Steel = tk.StringVar()
        self.Zinc  = tk.StringVar() 

        self.ABS1 = tk.StringVar()
        self.Aluminium1 = tk.StringVar()
        self.Brass1 = tk.StringVar()
        self.CastIron1 = tk.StringVar()
        self.Ceramic1 = tk.StringVar()
        self.Copper1 = tk.StringVar()
        self.ElectronicComponent1 = tk.StringVar()
        self.ExpandedPolystyrene1 = tk.StringVar()
        self.Glass1 = tk.StringVar()        
        self.Insulation1 = tk.StringVar()
        self.Iron1 = tk.StringVar()
        self.Lithium1 = tk.StringVar()
        self.Plastics1 = tk.StringVar()
        self.Polyamide1 = tk.StringVar()
        self.Polycarbonate1 = tk.StringVar()
        self.Polyethylene1  = tk.StringVar()          
        self.PolyurethaneFoam1 = tk.StringVar()
        self.PrintedBoard1 = tk.StringVar()
        self.PVCPipe1 = tk.StringVar()
        self.PVC1 = tk.StringVar()
        self.Rubber1  = tk.StringVar()
        self.Silicon1 = tk.StringVar()
        self.StainlessSteel1 = tk.StringVar()
        self.Steel1 = tk.StringVar()
        self.Zinc1  = tk.StringVar()            
        
        self.status = tk.StringVar()
        self.create_page()

    def create_page(self):
        
        tk.Label(self, text="Product Information ").grid(row=1, padx=20, stick=tk.W, pady=2) 
        
        tk.Label(self, text="Product Name: ").grid(row=2, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.ProductName).grid(row=2, padx=20, column=1, stick=tk.E)

        tk.Label(self, text="Total Weight: ").grid(row=3, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.TotalWeight).grid(row=3, padx=20, column=1, stick=tk.E)

        tk.Label(self, text="Energy Use Pre Product: ").grid(row=4, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.EnergyUse).grid(row=4, padx=20, column=1, stick=tk.E)
        
        tk.Label(self, text="Power Consumption: ").grid(row=5, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.PowerConsumption).grid(row=5, padx=20, column=1, stick=tk.E)
        
        tk.Label(self, text="Service Life: ").grid(row=6, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.ServiceLife).grid(row=6, padx=20, column=1, stick=tk.E)
 
        tk.Label(self, text="Transport to Factory(Local): ").grid(row=7, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.TransporttoFactoryl).grid(row=7, padx=20, column=1, stick=tk.E)
        
        tk.Label(self, text="Transport to Factory(Road): ").grid(row=8, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.TransporttoFactoryr).grid(row=8, padx=20, column=1, stick=tk.E)
        
        tk.Label(self, text="Transport to Factory(Sea): ").grid(row=9, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.TransporttoFactorys).grid(row=9, padx=20, column=1, stick=tk.E)
        
        tk.Label(self, text="Transport to Site: ").grid(row=10, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.TransporttoSite).grid(row=10, padx=20, column=1, stick=tk.E)

        tk.Label(self, text="Transport to Landfill: ").grid(row=11, padx=20, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.TransporttoLandfill).grid(row=11, padx=20, column=1, stick=tk.E) 


        tk.Label(self, text="Component Information ").grid(row=1, padx=20, column=2, stick=tk.W, pady=2) 
        tk.Label(self, text="Remanufactured Product ").grid(row=1, padx=20, column=3, stick=tk.W, pady=2) 
        tk.Label(self, text="New Product ").grid(row=1, padx=20, column=4, stick=tk.W, pady=2) 
        
        tk.Label(self, text="ABS %: ").grid(row=2, padx=20, column=2,stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.ABS).grid(row=2, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.ABS1).grid(row=2, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Aluminium %: ").grid(row=3, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Aluminium).grid(row=3, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Aluminium1).grid(row=3, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Brass %: ").grid(row=4, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Brass).grid(row=4, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Brass1).grid(row=4, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Cast iron %: ").grid(row=5, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.CastIron).grid(row=5, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.CastIron1).grid(row=5, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Ceramic %: ").grid(row=6, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Ceramic).grid(row=6, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Ceramic1).grid(row=6, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Copper %: ").grid(row=7, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Copper).grid(row=7, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Copper1).grid(row=7, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Electronic component %: ").grid(row=8, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.ElectronicComponent).grid(row=8, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.ElectronicComponent1).grid(row=8, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Expanded polystyrene %: ").grid(row=9, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.ExpandedPolystyrene).grid(row=9, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.ExpandedPolystyrene1).grid(row=9, padx=20, column=4, stick=tk.E)
       
        tk.Label(self, text="Glass %: ").grid(row=10, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Glass).grid(row=10, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Glass1).grid(row=10, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Insulation (general) %: ").grid(row=11, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Insulation).grid(row=11, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Insulation1).grid(row=11, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Iron %: ").grid(row=12, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Iron).grid(row=12, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Iron1).grid(row=12, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Lithium %: ").grid(row=13, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Lithium).grid(row=13, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Lithium1).grid(row=13, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Plastics (general) %: ").grid(row=14, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Plastics).grid(row=14, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Plastics1).grid(row=14, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Polyamide %: ").grid(row=15, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Polyamide).grid(row=15, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Polyamide1).grid(row=15, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Polycarbonate %: ").grid(row=16, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Polycarbonate).grid(row=16, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Polycarbonate1).grid(row=16, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Polyethylene %: ").grid(row=17, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Polyethylene).grid(row=17, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Polyethylene1).grid(row=17, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Polyurethane foam %: ").grid(row=18, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.PolyurethaneFoam).grid(row=18, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.PolyurethaneFoam1).grid(row=18, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Printed wiring board, mixed mounted %: ").grid(row=19, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.PrintedBoard).grid(row=19, padx=20, column=3, stick=tk.E) 
        tk.Entry(self, textvariable=self.PrintedBoard1).grid(row=19, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="PVC pipe %: ").grid(row=20, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.PVCPipe).grid(row=20, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.PVCPipe1).grid(row=20, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="PVC %: ").grid(row=21, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.PVC).grid(row=21, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.PVC1).grid(row=21, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Rubber %: ").grid(row=22, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Rubber).grid(row=22, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Rubber1).grid(row=22, padx=20, column=4, stick=tk.E)

        tk.Label(self, text="Silicon %: ").grid(row=23, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Silicon).grid(row=23, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Silicon1).grid(row=23, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Stainless steel %: ").grid(row=24, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.StainlessSteel).grid(row=24, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.StainlessSteel1).grid(row=24, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Steel (general or galvanised) %: ").grid(row=25, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Steel).grid(row=25, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Steel1).grid(row=25, padx=20, column=4, stick=tk.E)
        
        tk.Label(self, text="Zinc %: ").grid(row=26, padx=20, column=2, stick=tk.W, pady=2)
        tk.Entry(self, textvariable=self.Zinc).grid(row=26, padx=20, column=3, stick=tk.E)
        tk.Entry(self, textvariable=self.Zinc1).grid(row=26, padx=20, column=4, stick=tk.E)
        
        tk.Button(self, text="Query", command=self._search).grid(row=27, padx=10,column=1, stick=tk.W, pady=2)
        tk.Button(self, text="Change", command=self._change).grid(row=27, padx=10, column=2, stick=tk.E, pady=2)
        
        tk.Label(self, textvariable=self.status).grid(row=28, padx=10, column=2, stick=tk.E, pady=2)
        


    def _search(self):
        ProductName = self.ProductName.get()
        product = db.search_by_name(ProductName)
        if product:
            self.TotalWeight.set(product["Total Weight"])
            self.EnergyUse.set(product["Energy Use Pre Product"])
            self.PowerConsumption.set(product["Power Consumption"])
            self.ServiceLife.set(product["Service Life"])
            self.TransporttoFactoryl.set(product["Transport to Factory(Local)"])
            self.TransporttoFactoryr.set(product["Transport to Factory(Road)"])
            self.TransporttoFactorys.set(product["Transport to Factory(Sea)"])            
            self.TransporttoSite.set(product["Transport to Site"])
            self.TransporttoLandfill.set(product["Transport to Landfill"])
            
            self.ABS.set(product["ABS %"])
            self.Aluminium.set(product["Aluminium %"])
            self.Brass.set(product["Brass %"])
            self.CastIron.set(product["Cast iron %"])
            self.Ceramic.set(product["Ceramic %"])
            self.Copper.set(product["Copper %"])
            self.ElectronicComponent.set(product["Electronic component %"])
            self.ExpandedPolystyrene.set(product["Expanded polystyrene %"])
            self.Glass.set(product["Glass %"])    
            self.Insulation.set(product["Insulation (general) %"])
            self.Iron.set(product["Iron %"])
            self.Lithium.set(product["Lithium %"])
            self.Plastics.set(product["Plastics (general) %"])
            self.Polyamide.set(product["Polyamide %"])
            self.Polycarbonate.set(product["Polycarbonate %"])
            self.Polyethylene.set(product["Polyethylene %"])         
            self.PolyurethaneFoam.set(product["Polyurethane foam %"])
            self.PrintedBoard.set(product["Printed wiring board, mixed mounted %"])
            self.PVCPipe.set(product["PVC pipe %"])
            self.PVC.set(product["PVC %"])
            self.Rubber.set(product["Rubber %"])
            self.Silicon.set(product["Silicon %"])
            self.StainlessSteel.set(product["Stainless steel %"])
            self.Steel.set(product["Steel (general or galvanised) %"])
            self.Zinc.set(product["Zinc %"]) 
            
            self.ABS1.set(product["ABS(new) %"])
            self.Aluminium1.set(product["Aluminium(new) %"])
            self.Brass1.set(product["Brass(new) %"])
            self.CastIron1.set(product["Cast iron(new) %"])
            self.Ceramic1.set(product["Ceramic(new) %"])
            self.Copper1.set(product["Copper(new) %"])
            self.ElectronicComponent1.set(product["Electronic component(new) %"])
            self.ExpandedPolystyrene1.set(product["Expanded polystyrene(new) %"])
            self.Glass1.set(product["Glass(new) %"])    
            self.Insulation1.set(product["Insulation (general)(new) %"])
            self.Iron1.set(product["Iron(new) %"])
            self.Lithium1.set(product["Lithium(new) %"])
            self.Plastics1.set(product["Plastics (general)(new) %"])
            self.Polyamide1.set(product["Polyamide(new) %"])
            self.Polycarbonate1.set(product["Polycarbonate(new) %"])
            self.Polyethylene1.set(product["Polyethylene(new) %"])         
            self.PolyurethaneFoam1.set(product["Polyurethane foam(new) %"])
            self.PrintedBoard1.set(product["Printed wiring board, mixed mounted(new) %"])
            self.PVCPipe1.set(product["PVC pipe(new) %"])
            self.PVC1.set(product["PVC(new) %"])
            self.Rubber1.set(product["Rubber(new) %"])
            self.Silicon1.set(product["Silicon(new) %"])
            self.StainlessSteel1.set(product["Stainless steel(new) %"])
            self.Steel1.set(product["Steel (general or galvanised)(new) %"])
            self.Zinc1.set(product["Zinc(new) %"])
   
            self.status.set('Query Successfully')
        else:
            self.status.set('No Information of This Product')


    def _change(self):
        pdt = {
            "Product Name": self.ProductName.get(),
            "Total Weight": self.TotalWeight.get(),
            "Energy Use Pre Product": self.EnergyUse.get(),  
            "Power Consumption": self.PowerConsumption.get(),
            "Service Life": self.ServiceLife.get(),
            "Transport to Factory(Local)": self.TransporttoFactoryl.get(),
            "Transport to Factory(Road)": self.TransporttoFactoryr.get(),
            "Transport to Factory(Sea)": self.TransporttoFactorys.get(),            
            "Transport to Site": self.TransporttoSite.get(),
            "Transport to Landfill": self.TransporttoLandfill.get(), 
            
            "ABS %": self.ABS.get(),
            "Aluminium %": self.Aluminium.get(),
            "Brass %": self.Brass.get(),
            "Cast iron %": self.CastIron.get(),  
            "Ceramic %": self.Ceramic.get(),
            "Copper %": self.Copper.get(),
            "Electronic component %": self.ElectronicComponent.get(),
            "Expanded polystyrene %": self.ExpandedPolystyrene.get(),
            "Glass %": self.Glass.get(),           
            "Insulation (general) %": self.Insulation.get(),
            "Iron %": self.Iron.get(),
            "Lithium %": self.Lithium.get(),
            "Plastics (general) %": self.Plastics.get(),  
            "Polyamide %": self.Polyamide.get(),
            "Polycarbonate %": self.Polycarbonate.get(),
            "Polyethylene %": self.Polyethylene.get(),
            "Polyurethane foam %": self.PolyurethaneFoam.get(),
            "Printed wiring board, mixed mounted %": self.PrintedBoard.get(),             
            "PVC pipe %": self.PVCPipe.get(),
            "PVC %": self.PVC.get(),
            "Rubber %": self.Rubber.get(),
            "Silicon %": self.Silicon.get(),  
            "Stainless steel %": self.StainlessSteel.get(),
            "Steel (general or galvanised) %": self.Steel.get(),
            "Zinc %": self.Zinc.get(),
            
            "ABS(new) %": self.ABS1.get(),
            "Aluminium(new) %": self.Aluminium1.get(),
            "Brass(new) %": self.Brass1.get(),
            "Cast iron(new) %": self.CastIron1.get(),  
            "Ceramic(new) %": self.Ceramic1.get(),
            "Copper(new) %": self.Copper1.get(),
            "Electronic component(new) %": self.ElectronicComponent1.get(),
            "Expanded polystyrene(new) %": self.ExpandedPolystyrene1.get(),
            "Glass(new) %": self.Glass1.get(),           
            "Insulation (general)(new) %": self.Insulation1.get(),
            "Iron(new) %": self.Iron1.get(),
            "Lithium(new) %": self.Lithium1.get(),
            "Plastics (general)(new) %": self.Plastics1.get(),  
            "Polyamide(new) %": self.Polyamide1.get(),
            "Polycarbonate(new) %": self.Polycarbonate1.get(),
            "Polyethylene(new) %": self.Polyethylene1.get(),
            "Polyurethane foam(new) %": self.PolyurethaneFoam1.get(),
            "Printed wiring board, mixed mounted(new) %": self.PrintedBoard1.get(),             
            "PVC pipe(new) %": self.PVCPipe1.get(),
            "PVC(new) %": self.PVC1.get(),
            "Rubber(new) %": self.Rubber1.get(),
            "Silicon(new) %": self.Silicon1.get(),  
            "Stainless steel(new) %": self.StainlessSteel1.get(),
            "Steel (general or galvanised)(new) %": self.Steel1.get(),
            "Zinc(new) %": self.Zinc1.get()            
        }  # A product
        
        r = db.update(pdt)
        db.save_data()
        if r:
            self.status.set("Change Successfully")
        else:
            self.status.set("Change Unsuccessfully")


class AboutFrame(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.root = master
        self.create_page()

    def create_page(self):
        
        tk.Label(self, text='Author: Shaojiang').pack()
        tk.Label(self, text='CopyRight: Egglighting Company').pack()

