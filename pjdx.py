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
from kivy.core.window import Window
configuracion = ConfigParser(allow_no_value=True)
diccionario = ConfigParser(allow_no_value=True)
#print(os.path.dirname(__file__))
cfg_path = os.path.join(os.path.dirname(__file__),'config.cfg')
diccionario_path = os.path.join(os.path.dirname(__file__),'diccionarios.cfg')
Window.size = (470,700)
configuracion.optionxform = lambda option: option
configuracion.read(cfg_path,encoding="utf8")

diccionario.optionxform = lambda option: option
diccionario.read(diccionario_path,encoding="utf8")

df1_final = pd.DataFrame()
df2_final = pd.DataFrame()
df3_final = pd.DataFrame()



RUTA=""
import warnings
import re


dict_h1={'id_cliente_cli':'Id_Cliente',
'cliente_cli':'Cliente',
'rut_cliente_cli':'RUT Cliente',
'direccion_cli':'Dirección',
'potencia_conectada_cli':'Potencia Conectada',
'suministrador_cli':'Suministrador',
'rut_suministrador_cli':'RUT Suministrador',
'nombre_receptor_cli':'Nombre Receptor',
'rut_receptor_cli':'RUT Receptor',
'distribuidor_fac':'Distribuidor',
'rut_distribuidor_fac':'RUT Distribuidor',
'tipo_proceso_fac':'Tipo Proceso',
'tipodte_fac':'TipoDTE',
'n_dte_fac':'N° DTE',
'n_dte_original_fac':'N° DTE Original',
'fecha_emision_factura_fac':'Fecha emisión factura',
'fecha_vencimiento_factura_fac':'Fecha de vencimiento factura',
'peajedx_cargo_fijo_fac':'PeajeDx (Cargo Fijo)',
'peajedx_energia_fac':'PeajeDx (Energía)',
'peajedx_dda_max_pot_sum_fac':'PeajeDx (Dda Max Pot Sum)',
'peajedx_pot_fact_comp_fac':'PeajeDx (Pot Fact Comp)',
'peajedx_dda_max_pot_lhp_fac':'PeajeDx (Dda Max Pot LHP)',
'otros_cargos_fac':'Otros Cargos',
'otros_cargos_exento_fac':'Otros Cargos EXENTO',
'tarifa_tar':'Tarifa',
'empresa_dx_tar':'Empresa_Dx',
'comuna_tar':'Comuna',
'sistema_transmision_tar':'Sistema_Transmisión',
'tipo_suministro_tar':'Tipo_Suministro',
'subestacion_primaria_tar':'Subestación_Primaria',
'inicio_lectura_cons':'Inicio Lectura',
'fin_lectura_cons':'Fin Lectura',
'energia_cons':'Energía',
'dda_max_pot_sum_cons':'Dda Max Pot Sum',
'pot_fact_comp_cons':'Pot Fact Comp',
'dda_max_pot_lhp_cons':'Dda Max Pot LHP'
}

dict_h2={
'rut_distribuidora':'RUT DISTRIBUIDORA',
'n_dte':'Nº DTE',
'fecha_pago':'Fecha de pago total',
'fecha_pago_parcial':'fecha pago parcial',
'saldo_adeudado':'saldo adeudado',
'comentario':'Comentario'
}


dict_h3={
'distribuidor':'Distribuidor',
'rut_distribuidor':'RUT Distribuidor',
'cliente':'Cliente',
'rut_cliente':'RUT Cliente',
'potencia_conectada':'Potencia Conectada',
'direccion':'Dirección',
'regimen_actual':'RegimenActual',
'regimen_futuro':'RegimenFuturo',
'fecha_aviso':'Fecha_Aviso',
'fecha_cambio':'Fecha_Cambio'
}

warnings.simplefilter(action='ignore', category=UserWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)
mes_sel = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio',
          'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

mes_rev=mes_sel.reverse()

KV = '''
MDScreen:

    MDSpinner:
        id: spinner
        size_hint: None, None
        size: dp(46), dp(46)
        pos_hint: {'center_x': .5, 'center_y': .5}
        active: False

    MDBoxLayout:
        orientation: "vertical"

        MDTopAppBar:
            title: "Peajes de Distribución"
            elevation: 3

        MDFloatLayout:

            MDFillRoundFlatIconButton:
                id: dirbtn
                text: "Seleccione directorio PJDX"
                icon: "folder"
                pos_hint: {"center_x": .5, "center_y": .6}
                on_release: app.file_manager_open()
                disabled: False
        
                
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
   
        

def pjdx_mensual(self,mes,ruta,anno):

    global df1_final
    global df2_final
    global df3_final

    #if mes=='Junio':
    #    print('1')

    #print('-------1-------')
    nan_value = float("NaN") 
    #if mes!="Abril":
    #    exit()
    #print('-------2-------')
    hoja1='1_Cobro_Peajes'
    hoja2='2_Pago_Peajes'
    hoja3='3_Cambio_Regimen'
    archivos = glob.glob(ruta + "\*.xlsx")
    df1_total = pd.DataFrame()
    df2_total = pd.DataFrame()
    df3_total = pd.DataFrame()
    #print('-------3-------')
    for file in archivos:
        #if mes=='Junio':
        #    print('2')
        #print('-------4-------')
        if file.endswith('.xlsx'):
            #print('-------5-------')

            #if mes=='Junio':
            #    print('3')
            #    print(file)

            excel_file = pd.ExcelFile(file)
            df1 = pd.read_excel(file, sheet_name=hoja1)
            df1.rename(columns=dict_h1, inplace=True)

            if not(df1.columns.values[0]=='Id_Cliente' and df1.columns.values[1]=='Cliente' and df1.columns.values[2]=='RUT Cliente'):
                #if mes=='Junio':
                #    print(df1.head())  
                df1 = df1.iloc[df1[df1.iloc[:, 0].eq('Id_Cliente')].index[0]:, :].reset_index(drop=True)
                #if mes=='Junio':
                #    print(df1.head())  
                df1.columns = df1.iloc[0]
                #if mes=='Junio':
                #    print(df1.head())  
            
            df1.replace("", nan_value, inplace=True) 
            df1.dropna(how='all', axis=0, inplace=True)
            
            df1['ifc_mes']=mes
            df1['ifc_año']=anno
            #if mes=='Junio':
            #    print('4')
            df1['Suministrador']=df1['Suministrador'].str.strip()
            #df1['Suministrador']=df1['Suministrador'].astype(str)
            #df1['Suministrador']=df1['Suministrador'].apply(lambda x: x.str.strip())
                                                     
            

            df1 = df1.drop(0).reset_index(drop=True)
        
            df1_total = df1_total.append(df1)
            df2 = excel_file.parse(sheet_name = hoja2, header=0)
            df2.rename({'N°Cliente': 'Id_Cliente'}, axis=0,inplace=True)
            df2.rename(columns=dict_h2, inplace=True)

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
            df3.rename(columns=dict_h3, inplace=True)
            df3_total = df3_total.append(df3)
            #if mes=='Junio':
            #    print('------------------------5')
            

            #print('-------6-------')
    #print('-------7-------')
    
    #if mes=='Junio':
    #    print('6')
    
    df1_total['RUT Cliente']=df1_total['RUT Cliente'].apply(limpieza_RUT)
    df1_total['RUT Suministrador']=df1_total['RUT Suministrador'].apply(limpieza_RUT)
    df1_total['RUT Receptor']=df1_total['RUT Receptor'].apply(limpieza_RUT)
    df1_total['RUT Distribuidor']=df1_total['RUT Distribuidor'].apply(limpieza_RUT)
    df2_total['RUT DISTRIBUIDORA']=df2_total['RUT DISTRIBUIDORA'].apply(limpieza_RUT)
    df3_total['RUT Distribuidor']=df3_total['RUT Distribuidor'].apply(limpieza_RUT)
    df3_total['RUT Cliente']=df3_total['RUT Cliente'].apply(limpieza_RUT) 
    #df1_total['Suministrador']=df1_total['Suministrador'].astype(str)
    #df1_total['Suministrador']=df1_total['Suministrador'].apply(lambda a: a.str.strip())
    
    # se aplica diccionario
    #df1_total=df1_total.apply(lambda a: a.strip())
    #df2_total=df2_total.apply(lambda a: a.strip())
    #df3_total=df3_total.apply(lambda a: a.strip())
    
    #if mes=='Junio':
    #    print('7')

    for section in diccionario.sections():
        #if section !='Tipo Proceso':
        for elemento in diccionario[section]:
            #print(elemento)
        #    print(section)
        #    if mes=='Junio':
        #        print('8')

            #print(diccionario[section][elemento])
            df1_total[section]=df1_total[section].replace(elemento, diccionario[section][elemento])
            if section =='Distribuidor':
                df3_total[section]=df3_total[section].replace(elemento, diccionario[section][elemento])
    # fin de diccionario 
      
    #print('-------8-------') 
    archivo_errores = open(RUTA+"\errores_pjdx_"+mes+"_"+anno+".txt", "w")
    for section in configuracion.sections():
        tablas_maestras=[]
        errores=[]
        errores3=[]
        #if mes=='Junio':
        #    print('9')
        for elemento in configuracion[section]:
            tablas_maestras.append(elemento)
        archivo_errores.write("["+section+"]" + "\n")
        if section != 'Tipo Regimen':
            errores=set(list(df1_total[section]))-set(tablas_maestras)
        if section == 'Distribuidor':
            errores3=set(list(df3_total[section]))-set(tablas_maestras)
            errores=errores|errores3 
        if section == 'RUT Distribuidor':
            errores3=set(list(df2_total['RUT DISTRIBUIDORA']))-set(tablas_maestras)
            errores=errores|errores3 
        for item in errores:
            archivo_errores.write("%s\n" % item)
        archivo_errores.write("\n\n")
        #print(set(list(df1_total[section])))
        #print(set(tablas_maestras))
    archivo_errores.close() 

    #if mes=='Junio':
    #    print('10')

    df1_total = df1_total.astype({'ifc_año':'int'})
    df2_total = df2_total.astype({'ifc_año':'int'})
    df3_total = df3_total.astype({'ifc_año':'int'})

   



    df1_total = df1_total.sort_values(by='ifc_año',ascending=False, ignore_index=True)
    df2_total = df2_total.sort_values(by='ifc_año',ascending=False, ignore_index=True)
    df3_total = df3_total.sort_values(by='ifc_año',ascending=False, ignore_index=True)

    with ExcelWriter(RUTA+"\pjdx_"+mes+"_"+anno+".xlsx") as writer:
        df1_total.to_excel(writer, hoja1, index=False)
        df2_total.to_excel(writer, hoja2, index=False)
        df3_total.to_excel(writer, hoja3, index=False)

    df1_final = df1_final.append(df1_total)
    df2_final = df2_final.append(df2_total)
    df3_final = df3_final.append(df3_total)



    self.root.ids.spinner.active=False
    print("Fin "+mes+" "+anno) 
    

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
        hoja1='1_Cobro_Peajes'
        hoja2='2_Pago_Peajes'
        hoja3='3_Cambio_Regimen'
        
        self.root.ids.spinner.active=True
        p=RUTA
        for path in os.listdir(p):
            for subpath in os.listdir(p+"\\"+path):
                try:
                    #print(mes_sel[mes_sel.index(subpath)]+" existe")
                    pjdx_mensual(self,subpath,p+"\\"+path+"\\"+subpath,path)
                except:
                    pass


        #df1_final['ifc_mes'] = pd.Categorical(df1_final['ifc_mes'], ordered=True, categories=mes_rev)
        #df1_final = df1_final.sort_values('ifc_mes')
        #df2_final['ifc_mes'] = pd.Categorical(df2_final['ifc_mes'], ordered=True, categories=mes_rev)
        #df2_final = df2_final.sort_values('ifc_mes')
        #df3_final['ifc_mes'] = pd.Categorical(df3_final['ifc_mes'], ordered=True, categories=mes_rev)
        #df3_final = df3_final.sort_values('ifc_mes')


                    
        #df1_final = df1_final.sort_values(by='ifc_año',ascending=False, ignore_index=True)
        #df2_final = df2_final.sort_values(by='ifc_año',ascending=False, ignore_index=True)
        #df3_final = df3_final.sort_values(by='ifc_año',ascending=False, ignore_index=True)

        with ExcelWriter(RUTA+"\pjdx_completo.xlsx") as writer:
            df1_final.to_excel(writer, hoja1, index=False)
            df2_final.to_excel(writer, hoja2, index=False)
            df3_final.to_excel(writer, hoja3, index=False)
        print("Fin programa") 
               

                 



PJDX_AA().run()





