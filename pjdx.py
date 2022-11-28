from sre_constants import AT_NON_BOUNDARY
import pandas as pd
import glob
from pandas import ExcelWriter
from kivymd.app import MDApp
from kivymd.uix.label import MDLabel
from kivymd.uix.screen import MDScreen
from kivymd.uix.button import MDRectangleFlatButton
from kivy.core.text import LabelBase
from kivymd.font_definitions import theme_font_styles
from kivy.lang import Builder
import os
from kivy.core.window import Window
from kivymd.uix.filemanager import MDFileManager
from kivymd.toast import toast
from pathlib import Path
from configparser import ConfigParser
configuracion = ConfigParser(allow_no_value=True)
diccionario = ConfigParser(allow_no_value=True)
#print(os.path.dirname(__file__))
cfg_path = os.path.join(os.path.dirname(__file__),'config.cfg')
diccionario_path = os.path.join(os.path.dirname(__file__),'diccionarios.cfg')

configuracion.optionxform = lambda option: option
configuracion.read(cfg_path,encoding="utf8")

diccionario.optionxform = lambda option: option
diccionario.read(diccionario_path,encoding="utf8")


RUTA=""
import warnings
import re

warnings.simplefilter(action='ignore', category=FutureWarning)
mes_sel = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio',
          'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
KV = '''
MDBoxLayout:
    orientation: "vertical"

    MDTopAppBar:
        title: "Peajes de Distribución"
        elevation: 3

    MDFloatLayout:

        MDFillRoundFlatIconButton:
            text: "Seleccione directorio PJDX"
            icon: "folder"
            pos_hint: {"center_x": .5, "center_y": .6}
            on_release: app.file_manager_open()
            
        MDFillRoundFlatIconButton:
            id: botonRUN
            text: " Iniciar  Cálculo  de  PJDX  "
            icon: "android"
            pos_hint: {"center_x": .5, "center_y": .5}
            on_release: app.get_sub_folders()
            disabled: True
            
        
        
'''

def limpieza_RUT(rut_revisar):
    #print('3')
    rut=""
    #print('--1--')
    #for rutx in rut_revisar:
    #rut=str(rut_revisar)
    #i=0
    #print(rut_revisar)
    #print(rut)
    #print(rutx)
    #print(i)
    #print('--2--')
    #print(rut)
    rut=rut_revisar
    if not pd.isnull(rut):
        rut=rut.strip() #elimina espacio al principio o final
        #print('--3--')
        if rut[-1]=='K'or rut[-1]=='k':
            #print('--4--')
            rut = re.sub(r'[^0-9]', '', rut)+'K'
            #print('--5--')
        else:
            #print('--6--')
            rut = re.sub(r'[^0-9]', '', rut) #elimina caracteres no numerico dejando la "K" para rut que terminan en k o K  
            #print('--7--')
        if len(rut) > 1:
            #print('--7.5--')
            #print(rut)
            #print(len(rut))
            #print(rut[len(rut)-1])
            #print(rut[:(len(rut)-1)])
            
            rut=rut[:(len(rut)-1)]+"-"+rut[len(rut)-1]
            #print(rut)
            #print('--8--')
        if len(rut) > 5:        
            rut=rut[:(len(rut)-5)]+"."+rut[(len(rut)-5):]
            #print(rut)
            #print('--9--')
        if len(rut) > 9:
            rut=rut[:(len(rut)-9)]+"."+rut[(len(rut)-9):]
            #print(rut)
            #print('--10--')
        if rut!='' and rut!='K':
            rut_revisar=rut
            #print('--11--')
            #rut_revisar.replace(rut_revisarrut)
            #df1_total.replace("", nan_value, inplace=True) 
            #return rut
            #print(rut)
            #print('--12--')
            #rut_revisar.replace(rut_revisar,rut, inplace=True)
        #i=i+1
        #print(rut)
        #print(rutx)
    return rut_revisar
   
def aplicar_strip(x):
    x=x.strip()
   
        

def pjdx_mensual(mes,ruta,anno):
    print('-------1-------')
    nan_value = float("NaN") 
    if mes!="Abril":
        exit()
    print('-------2-------')
    hoja1='1_Cobro_Peajes'
    hoja2='2_Pago_Peajes'
    hoja3='3_Cambio_Regimen'
    archivos = glob.glob(ruta + "\*.xlsx")
    df1_total = pd.DataFrame()
    df2_total = pd.DataFrame()
    df3_total = pd.DataFrame()
    print('-------3-------')
    for file in archivos:
        print('-------4-------')
        if file.endswith('.xlsx'):
            print('-------5-------')
            excel_file = pd.ExcelFile(file)
            df1 = pd.read_excel(file, sheet_name=hoja1)
            df1 = df1.iloc[df1[df1.iloc[:, 0].eq('Id_Cliente')].index[0]:, :].reset_index(drop=True)
            df1.columns = df1.iloc[0]
       
            
            df1.replace("", nan_value, inplace=True) 
            df1.dropna(how='all', axis=0, inplace=True)
            
            df1['ifc_mes']=mes
            df1['ifc_año']=anno
            df1 = df1.drop(0).reset_index(drop=True)
        
            df1_total = df1_total.append(df1)
            df2 = excel_file.parse(sheet_name = hoja2, header=0)
            df2.rename({'N°Cliente': 'Id_Cliente'}, axis=0,inplace=True)
            
            df2.replace("", nan_value, inplace=True) 
            df2.dropna(how='all', axis=0, inplace=True)
            
            df2['ifc_mes']=mes
            df2['ifc_año']=anno
            df2_total = df2_total.append(df2)
            df3 = excel_file.parse(sheet_name = hoja3, header=0)
            
            df3_total.replace("", nan_value, inplace=True) 
            df3_total.dropna(how='all', axis=1, inplace=True)
             
            df3['ifc_mes']=mes
            df3['ifc_año']=anno
            df3_total = df3_total.append(df3)
            print('-------6-------')
    print('-------7-------')
    
  
    
    df1_total['RUT Cliente']=df1_total['RUT Cliente'].apply(limpieza_RUT)
    df1_total['RUT Suministrador']=df1_total['RUT Suministrador'].apply(limpieza_RUT)
    df1_total['RUT Receptor']=df1_total['RUT Receptor'].apply(limpieza_RUT)
    df1_total['RUT Distribuidor']=df1_total['RUT Distribuidor'].apply(limpieza_RUT)
    df2_total['RUT DISTRIBUIDORA']=df2_total['RUT DISTRIBUIDORA'].apply(limpieza_RUT)
    df3_total['RUT Distribuidor']=df3_total['RUT Distribuidor'].apply(limpieza_RUT)
    df3_total['RUT Cliente']=df3_total['RUT Cliente'].apply(limpieza_RUT) 
    
    # se aplica diccionario
    #df1_total=df1_total.apply(lambda a: a.strip())
    #df2_total=df2_total.apply(lambda a: a.strip())
    #df3_total=df3_total.apply(lambda a: a.strip())
    
    for section in diccionario.sections():
        if section !='Tipo Proceso':
            for elemento in diccionario[section]:
                print(elemento)
            #    print(section)
                print(diccionario[section][elemento])
                df1_total[section]=df1_total[section].replace(elemento, diccionario[section][elemento])
                if section =='Distribuidor':
                    df3_total[section]=df3_total[section].replace(elemento, diccionario[section][elemento])
    # fin de diccionario 
      
    print('-------8-------') 
    archivo_errores = open(RUTA+"\errores_pjdx_"+mes+"_"+anno+".txt", "w")
    for section in configuracion.sections():
        tablas_maestras=[]
        errores=[]
        errores3=[]
        for elemento in configuracion[section]:
            tablas_maestras.append(elemento)
        archivo_errores.write("["+section+"]" + "\n")
        if section != 'Tipo Regimen':
            errores=set(list(df1_total[section]))-set(tablas_maestras)
        if section == 'Distribuidor':
            errores3=set(list(df3_total[section]))-set(tablas_maestras)
            errores=errores|errores3   
        for item in errores:
            archivo_errores.write("%s\n" % item)
        archivo_errores.write("\n\n")
        #print(set(list(df1_total[section])))
        #print(set(tablas_maestras))
    archivo_errores.close() 
    with ExcelWriter(RUTA+"\pjdx_"+mes+"_"+anno+".xlsx") as writer:
        df1_total.to_excel(writer, hoja1, index=False)
        df2_total.to_excel(writer, hoja2, index=False)
        df3_total.to_excel(writer, hoja3, index=False)
    
    
    print("Fin programa") 
    

class PJDX_AA(MDApp):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        Window.bind(on_keyboard=self.events)
        self.manager_open = False
        self.file_manager = MDFileManager(
            exit_manager=self.exit_manager, select_path=self.select_path
        )

    def build(self):
        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "Blue"
        return Builder.load_string(KV)

    def file_manager_open(self):
        self.file_manager.show(os.path.dirname(os.path.dirname(__file__)))  # carpeta de usuario os.path.expanduser("~")
        self.manager_open = True

    def select_path(self, path: str):
        
        global RUTA
        '''
        It will be called when you click on the file name
        or the catalog selection button.

        :param path: path to the selected directory or file;
        '''
        RUTA=path
        self.exit_manager()
        toast(path)
        self.root.ids.botonRUN.disabled=False
        

    def exit_manager(self, *args):
        '''Called when the user reaches the root of the directory tree.'''

        self.manager_open = False
        self.file_manager.close()

    def events(self, instance, keyboard, keycode, text, modifiers):
        '''Called when buttons are pressed on the mobile device.'''

        if keyboard in (1001, 27):
            if self.manager_open:
                self.file_manager.back()
        return True

    def get_sub_folders(self):
        p=RUTA
        for path in os.listdir(p):
            for subpath in os.listdir(p+"\\"+path):
                try:
                    #print(mes_sel[mes_sel.index(subpath)]+" existe")
                    pjdx_mensual(subpath,p+"\\"+path+"\\"+subpath,path)
                except:
                    pass


PJDX_AA().run()





