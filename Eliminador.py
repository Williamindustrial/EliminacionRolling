# -*- coding: utf-8 -*-
"""
Created on Mon Mar  3 16:27:40 2025

@author: williamtorres
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os
import win32com.client
import pyperclip
import time
import tkinter as tk
from tkinter import messagebox

class Eliminador:

    def __init__(self,file_path, campañaIniciorollingCorp, campañaIniciorollingPR, campañaFin):
        self.file_path= file_path
        self.inicioRollingCorporativo= float(campañaIniciorollingCorp)
        self.inicioRollingPR=float(campañaIniciorollingPR)
        self.campañaFin=float(campañaFin)
        self.CDL_filtrado= None
        self.centros= [
                        ('CO03', 'COLOMBIA'),
                        ('PE03', 'PERU'),
                        ('MX03', 'MEXICO'),
                        ('CL03', 'CHILE'),
                        ('GT23', 'GUATEMALA'),
                        ('SV13', 'EL SALVADOR'),
                        ('CR03', 'COSTA RICA'),
                        ('DO03', 'REPUBLICA DOMINICANA'),
                        ('PA33', 'PANAMA'),
                        ('BO03', 'BOLIVIA'),
                        ('EC03', 'ECUADOR'),
                        ('PR03', 'PUERTO RICO'),
                        ('US03', 'USA')
                    ]
        self.CampaniaDescontinuacionVector= None
        self.VectorCDP=None
        self.concantenarData()
        self.crearPandasEliminacionCDPCAMPAÑA()
        
    def actualizar_campania(self,row):
        if row['CampaniaDescontinuacion'] < self.inicioRollingCorporativo and row['CDP']!='PR03':
            row['CampaniaDescontinuacion'] = self.inicioRollingCorporativo
        elif row['CampaniaDescontinuacion'] < self.inicioRollingPR and row['CDP']=='PR03':
            row['CampaniaDescontinuacion'] = self.inicioRollingPR
        elif row['CampaniaDescontinuacion'] >= self.inicioRollingCorporativo and row['CDP']!='PR03':
            campaña=row['CampaniaDescontinuacion'] %100
            if(campaña>=18):
                primeras_cuatro_cifras = row['CampaniaDescontinuacion'] // 100
                campañaBorrado= str(int(primeras_cuatro_cifras))+"01"
                campañaBorrado=int(campañaBorrado)
            else:
                campañaBorrado=row['CampaniaDescontinuacion']+1
            row['CampaniaDescontinuacion']=campañaBorrado
        elif row['CampaniaDescontinuacion'] >= self.inicioRollingPR and row['CDP']=='PR03':
            campaña=row['CampaniaDescontinuacion'] %100
            if(campaña>=13):
                primeras_cuatro_cifras = row['CampaniaDescontinuacion'] // 100
                campañaBorrado= str(int(primeras_cuatro_cifras))+"01"
                campañaBorrado=int(campañaBorrado)
            else:
                campañaBorrado=row['CampaniaDescontinuacion']+1
        return row


    def concantenarData(self):
        CDP = pd.DataFrame(self.centros, columns=['CDP', 'Pais'])
        añoActual = datetime.now().year
        CDL_AÑOANTERIOR_a= pd.read_excel( self.file_path, sheet_name=str(añoActual-2))
        CDL_AÑOANTERIOR= pd.read_excel( self.file_path, sheet_name=str(añoActual-1))
        CDL_AÑOACTUAL= pd.read_excel( self.file_path, sheet_name=str(añoActual))
        CDL_AÑOACTUAL_1= pd.read_excel(self.file_path, sheet_name=str(añoActual+1))
        # Concatenar los dos DataFrames
        CDL_CONCATENADO = pd.concat([CDL_AÑOACTUAL, CDL_AÑOACTUAL_1], ignore_index=True)
        CDL_CONCATENADO = pd.concat([CDL_CONCATENADO, CDL_AÑOANTERIOR_a], ignore_index=True)
        CDL_CONCATENADO = pd.concat([CDL_CONCATENADO, CDL_AÑOANTERIOR], ignore_index=True)
        try:
            CDL_AÑOACTUAL_2 = pd.read_excel(self.file_path, sheet_name=str(añoActual+2))
            # Concatenar los dos DataFrames
            CDL_CONCATENADO = pd.concat([CDL_CONCATENADO, CDL_AÑOACTUAL_2], ignore_index=True)
        except ValueError as e:
                print( "No existe la hoja "+str(añoActual+2))
        CDL_CONCATENADO = pd.merge(CDL_CONCATENADO, CDP, left_on="Pais", right_on="Pais", how="left")
        CDL_CONCATENADO = CDL_CONCATENADO[CDL_CONCATENADO['CampaniaDescontinuacion'].notna()]
        CDL_CONCATENADO.head()
        self.CDL_filtrado = CDL_CONCATENADO.apply(self.actualizar_campania, axis=1)
        self.CampaniaDescontinuacionVector= self.CDL_filtrado["CampaniaDescontinuacion"].unique()
        self.CampaniaDescontinuacionVector.sort()
        self.VectorCDP= self.CDL_filtrado["CDP"].unique()
        self.VectorCDP.sort()
        
        
    def mostrar_mensaje_error(self, mensaje):
        """Muestra un mensaje de error en una ventana emergente."""
    #root = tk.Tk()
    #root.withdraw()  # Ocultar la ventana principal
        messagebox.showerror("Error", mensaje)
        
    
    def conectarSAP(self):
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)  # Conexión activa
            session = connection.Children(0)  # Primera sesión activa
            
            return session
        except Exception as e:
            print(e)
            self.mostrar_mensaje_error(f"No esta logeado en SAP")
            return False
        
        
    def conectarzpp259(self,session):
        session.findById("wnd[0]/tbar[0]/okcd").text = "zpp259"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)  # Esperar 1 segundos antes de continuar
        session.findById("wnd[0]/usr/radP_B").setFocus()
        session.findById("wnd[0]/usr/radP_B").select()
        time.sleep(1)  # Esperar 1 segundos antes de continuar
        
        
    def pegarData(self, session, listaCodigos):
        session.findById("wnd[0]/usr/btn%_S_CPROD_%_APP_%-VALU_PUSH").press()
        # Convertir la lista en un string con saltos de línea
        texto_a_pegar=""
        for i in range(len(listaCodigos)):
            texto_a_pegar= texto_a_pegar+"\r\n"+str(listaCodigos[i])
        #texto_a_pegar = "\r\n".join(listaCodigos)  # "\r\n" es un salto de línea en SAP
        # Copiar al portapapeles
        pyperclip.copy(texto_a_pegar)
        # Hacer clic en la primera celda de la tabla para asegurarse de que está activa
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,0]").setFocus()
        # Pegar con Ctrl + V (sendVKey 86)
        session.findById("wnd[0]/usr/btn%_S_CPROD_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        # Confirmar la entrada
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    
    
    def cambioData(self, session):
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/usr/btn%_S_CPROD_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        
    
    def añadirCDPcampaña(self,session, CDP, campañaEliminacion):
        campañaEliminacion= str(int(campañaEliminacion))
        session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = str(CDP)
        session.findById("wnd[0]/usr/txtS_COMCAM-LOW").Text = campañaEliminacion
        session.findById("wnd[0]/usr/txtS_COMCAM-HIGH").Text = str(int(self.campañaFin))
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
    
    def crearPandasEliminacionCDPCAMPAÑA(self):
        session= self.conectarSAP()
        self.conectarzpp259(session)
        for campaña in self.CampaniaDescontinuacionVector:
            for centro in self.VectorCDP:
                CDL_Centro_Campaña= None
                # Filter the DataFrame using both 'centro' and 'campaña' as conditions
                CDL_Centro_Campaña = self.CDL_filtrado[(self.CDL_filtrado['CDP'] == centro) & (self.CDL_filtrado['CampaniaDescontinuacion'] == campaña)]
                codigos= CDL_Centro_Campaña["CodigoSAP"].unique()
                if(len(codigos)>0):
                    self.pegarData(session,codigos)
                    self.añadirCDPcampaña(session, centro, campaña)
                    self.cambioData(session)
        messagebox.showinfo("Éxito", "¡La operación se completó correctamente!")
    
    
        
    
    
    
