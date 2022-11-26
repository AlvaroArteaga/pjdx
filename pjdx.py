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
print(os.path.dirname(__file__))
cfg_path = os.path.join(os.path.dirname(__file__),'config.cfg')
print(cfg_path)
configuracion.read(cfg_path,encoding="utf8")


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
            text: " Iniciar  Cálculo  de  PJDX  "
            icon: "android"
            pos_hint: {"center_x": .5, "center_y": .5}
            
        
        
'''

def get_sub_folders(p):
    for path in os.listdir(p):
        for subpath in os.listdir(p+"\\"+path):
            try:
                print(mes_sel[mes_sel.index(subpath)]+" existe")
                pjdx_mensual(subpath,p+"\\"+path+"\\"+subpath)
            except:
                pass
            

def pjdx_mensual(mes,ruta):
    
    hoja1='1_Cobro_Peajes'
    hoja2='2_Pago_Peajes'
    hoja3='3_Cambio_Regimen'
    
    archivos = glob.glob(ruta + "\*.xlsx")

    df1_total = pd.DataFrame()
    df2_total = pd.DataFrame()
    df3_total = pd.DataFrame()
    
    for seccion in configuracion.sections():
        print(seccion)  

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
        '''
        It will be called when you click on the file name
        or the catalog selection button.

        :param path: path to the selected directory or file;
        '''

        self.exit_manager()
        toast(path)
        get_sub_folders(path)

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

PJDX_AA().run()





