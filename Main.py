

# region <<<<=====================================[Application Requirements]=====================================>>>>

# =======================================================|
#  Import (Config) to control APP Configuration Settings.|
# =======================================================|
from kivy.config import Config

# =============================================================================================================|
# Setting the APP to have Fixed Configuration (by putting False) which makes the user can't change Screen Size,|
# to keep the APP Organized.                                                                                   |
# =============================================================================================================|
Config.set('graphics', 'resizable', False)

# ================================|
# Import (MDApp) to Build the APP.|
# ================================|
from kivymd.app import MDApp

# ==================================================|
# Import (Window) to control the APP Window Setting.|
# ==================================================|
from kivy.core.window import Window

# ==================================================|
# Import (AsyncImage) to set APP Image from Website.|
# ==================================================|
from kivy.uix.image import AsyncImage

# =============================================================|
# Import (Builder) to create the KV file (APP Elements Layout).|
# =============================================================|
from kivy.lang.builder import Builder

# ==========================================================================|
# Import (ScreenManager) and (Screen) to create APP Screens and Manege them.|
# ==========================================================================|
from kivy.uix.screenmanager import ScreenManager, Screen

# ==========================================|
# Import (MDLabel) to Show some APP's texts.|
# ==========================================|
from kivymd.uix.label import MDLabel

# ===============================================================|
# Import (MDRaisedButton) as a button to execute the APP's actions.|
# ===============================================================|
from kivymd.uix.button import MDRaisedButton

# ===================================================|
# Import (MDBoxLayout) to contain all APP's Elements.|
# ===================================================|
from kivymd.uix.boxlayout import MDBoxLayout

# ==================================================|
# Import (BoxLayout) to contain some APP's Elements.|
# ==================================================|
from kivy.uix.boxlayout import BoxLayout

# ====================================================================|
# Import (MDTextField) and (TextInput) to Enter an Inputs for the APP.|
# ====================================================================|
from kivymd.uix.textfield import MDTextField
from kivy.uix.textinput import TextInput

# ======================================================================================|
# Import (MDDialog) as a Dialog window to inform the user about tasks and takes decisions.|
# ======================================================================================|
from kivymd.uix.dialog import MDDialog

# ================================================================|
# Import (OneLineAvatarListItem) to Create some APP's Lists Items.|
# ================================================================|
from kivymd.uix.list import OneLineAvatarListItem

# =======================================================================|
# Import (math) Library to use some Mathematics function (sin,cos...etc).|
# =======================================================================|
import math

# ====================================================================================|
# Import (glob) <built_in function in python> to search and find files inside folders.|
# ====================================================================================|
import glob

# ========================================================================================|
# Import (subprocess) <built_in function in python> to start and open an Application in__ |
# __Windows Operative System (Ex: CIMCO, Microsoft Word...Etc)                            |
# ========================================================================================|
import subprocess

# ===============================================================|
# Import (pandas) Library to Read and Write data of Excel Sheets.|
# ===============================================================|
import pandas as pd

# ===============================================|
# Import (openpyxl) Library to Load Excel Sheets.|
# ===============================================|
import openpyxl

# ===================================================|
# Import (smtplib) Library to Send and Manege Emails.|
# ===================================================|
import smtplib

# =============================================|
# Import (keyring) Library to Manege Passwords.|
# =============================================|
import keyring

# ======================================================|
# Import (date) Library to set and Manege Date and time.|
# ======================================================|
from datetime import date
today = date.today()
today_date = today.strftime("%m/%d/%Y")

# =============================================================================|
# Import (pyodbc) Library to connect the APP with the DataBase and Manege Data.|
# =============================================================================|
import pyodbc

# endregion <<<<====================================[Application Requirements]====================================>>>>


# region <<<<========================================[Screen Builder KV]=========================================>>>>

# region <<<<======================================[Screen Builder Notes]======================================>>>>

# ====================================================================================================|
#  All Notes of (Screens_Builder) Section will be above the code because can't include comment inside.|
# ====================================================================================================|

# region <<<<=======================================[General Notes]=========================================>>>>

# ===========================================================================================|
# [#] (ScreenManager) contains all APP Screens, and needs to add any new screen for it.      |
# [#] Each Screen has a specific name to be able to access it.                               |
# [#] Each Screen has a (MDLabel) to show texts of the screen.                               |
# [#] Each Screen has a (MDTextField) <with Specific id> to take Inputs.                     |
# [#] Each Screen has a (MDRaisedButton) to call Functions or to move between APP's Screens. |
#     Example of Button used to move between APP's Screens :                                 |
#                           on_press: root.manager.current = 'SettingScreen'                 |
#     Example of Button used to call Functions (Function MUST be Inside Screen Scope):       |
#                           on_press : root.create_program_for_old_horizontal_machine(object)|
# [#] Each Element has Certain Position, text format, color, and setting.                    |
# ===========================================================================================|

# endregion <<<<====================================[General Notes]=========================================>>>>

# region <<<<=======================================[Specific Notes]=========================================>>>>

# ================================================================================================================|
# [#] In LoginScreen: MDTextField of (id: Email) use specific text sitting of                                     |
#     (self.text.lower() if self.text is not None else '') : it used to make all input <small> letter.            |
# [#] In HomeScreen: Use Specific function (on_pre_enter) : it used to call function (inside HomeScreen Scope)    |
#     once entered the screen.                                                                                    |
# [#] In HomeScreen: MDLabel of (id: UserName) leave the 'text' empty to fill it by UserName when function called.|
# [#] In OldHorizontalScreen: MDTextField of (id: JobNumberForOldHorizontalMachine) use specific text sitting of  |
#     (self.text.upper() if self.text is not None else '' ) : it used to make all input <Capital> letter.         |
# ================================================================================================================|

# endregion <<<<====================================[Specific Notes]=========================================>>>>

# endregion <<<<======================================[Screen Builder Notes]======================================>>>>

Screens_Builder = """
ScreenManager:
    LoginScreen:
    HomeScreen:
    PinBoreScreen:
    OldHorizontalScreen:
    NewHorizontalScreen:
    SettingScreen:
    PinBoreSettingScreen:
    OldHorizontalSettingScreen:
    AppSettingScreen:
    UserSettingScreen:
    AddNewUserScreen:

    
<LoginScreen>:
    name: 'LoginScreen'
    MDLabel:
        text: 'Wiseco Programs Maker'
        pos_hint: {'center_x':0.78,'center_y':0.85}
        font_size: '34sp'
        bold: True
        italic: True
        theme_text_color: "Primary"
    MDLabel:
        text: 'An App to create CNC Programs according to the type of Operation and Machine'
        pos_hint: {'center_x':0.58,'center_y':0.75}
        font_size: '20sp'
        bold: True
        italic: True
        theme_text_color: "Secondary"     
    MDLabel:
        text: 'Enter Login Information If You are One of Programming Team'
        pos_hint: {'center_x':0.71,'center_y':0.57}
        font_size: '18sp'
        bold: True
        italic: True
        theme_text_color: "Hint"
    MDLabel:
        text: 'If You Are New User, Ask One of Programming Team to Add You'
        pos_hint: {'center_x':0.76,'center_y':0.50}
        font_size: '14sp'
        bold: True
        italic: True
        theme_text_color: "Hint"         
    MDTextField:
        id: Email
        text: self.text.lower() if self.text is not None else ''
        hint_text: "Enter Email Address"
        helper_text: "Use Active Email With Extension @rwbteam or @wiseco."
        helper_text_mode: "on_focus"
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.40}
        size_hint_x:None
        width:300
        height:10      
    MDTextField:
        id: Password
        hint_text: "Enter Password"
        helper_text: "Use The Shared Password You Got by Email."
        helper_text_mode: "on_focus"
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        password: True
        pos_hint: {'center_x': 0.50, 'center_y': 0.28}
        size_hint_x:None
        width:300
        height:10    
    MDRaisedButton:                                                                         
        text: 'LOGIN'
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        pos_hint: {'center_x':0.5,'center_y':0.18}
        on_press : 
            root.login_check()     
    MDRaisedButton:                                                                         
        text: 'Version 1.0.0'
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        pos_hint: {'center_x':0.07,'center_y':0.03}  
        on_press : 
            root.application_version_features()        
    MDLabel:
        text: 'Created by: Moemen Alatweh'
        pos_hint: {'center_x':1.32,'center_y':0.04}
        font_style: 'Caption'
        theme_text_color: "Custom"
        text_color: 175/255.0, 0/255.0, 0/255.0, 1
    MDLabel:
        text: 'malatweh@rwbteam.com'
        pos_hint: {'center_x':1.33,'center_y':0.01}
        font_style: 'Caption'
        theme_text_color: "Custom"
        text_color: 175/255.0, 0/255.0, 0/255.0, 1           
    
      
<HomeScreen>:
    name: 'HomeScreen'
    on_pre_enter: root.set_user_name()
    MDLabel:
        text: 'Welcome to Wiseco Programs Maker'
        pos_hint: {'center_x':0.70,'center_y':0.8}
        font_size: '32sp'
        bold: True
        italic: True
        theme_text_color: "Primary"
    MDLabel:
        id: UserName
        text: ''
        halign:'center'
        valign: 'middle'
        pos_hint: {'center_y':0.7}
        font_size: '24sp'
        bold: True
        italic: True
        theme_text_color: "Error"  
    MDLabel:
        text: 'Choose Type of Program'
        pos_hint: {'center_x':0.885,'center_y':0.55}
        font_size: '18sp'
        bold: True
        italic: True
        theme_text_color: "Secondary"    
    MDRaisedButton:
        text: 'Pin Bore'
        pos_hint: {'center_x':0.5,'center_y':0.45}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: root.manager.current = 'PinBoreScreen'    
    MDRaisedButton:
        text: 'Setting'
        halign:'left'
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: root.manager.current = 'SettingScreen'
    MDRaisedButton:
        text: 'LOGOUT'
        pos_hint: {'center_x':0.945,'center_y':0.035}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press : 
            root.logout()         
       
          
<PinBoreScreen>:
    name: 'PinBoreScreen'
    MDLabel:
        text: 'Pin Bore Programs Maker'
        pos_hint: {'center_x':0.80,'center_y':0.8}
        font_size: '30sp'
        bold: True
        italic: True
        theme_text_color: "Primary"
    MDLabel:
        text: 'Choose Machine That Need Program'
        pos_hint: {'center_x':0.83,'center_y':0.6}
        font_size: '18sp'
        bold: True
        italic: True
        theme_text_color: "Secondary"    
    MDRaisedButton:
        text: 'Old Horizontal Machines (28,29,32)'
        pos_hint: {'center_x':0.5,'center_y':0.5}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: root.manager.current = 'OldHorizontalScreen'
    MDRaisedButton:
        text: 'New Horizontal Machine (127)'
        pos_hint: {'center_x':0.5,'center_y':0.4}                    
        md_bg_color: 120/255, 0/255, 0/255, 1   
        font_size: "15sp"
        on_press : root.still_work_on_it(object) 
    MDRaisedButton:
        text: 'Back'
        pos_hint: {'center_x':0.5,'center_y':0.1}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: root.manager.current = 'HomeScreen'    
    
       
<OldHorizontalScreen>:
    name: 'OldHorizontalScreen'
    on_pre_enter:
    MDLabel:
        text: 'Old Horizontal Machines (28,29,32)'
        pos_hint: {'center_x':0.73,'center_y':0.8}
        font_size: '30sp'
        bold: True
        italic: True
        theme_text_color: "Primary"
    MDTextField:
        id: JobNumberForOldHorizontalMachine
        text: self.text.upper() if self.text is not None else ''
        hint_text: "Enter Job Number"
        helper_text: "Job number MUST match with number in Spec."
        helper_text_mode: "on_focus"
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.60}
        size_hint_x:None
        width:300
        height:10      
    MDRaisedButton:
        text: 'Submit'
        pos_hint: {'center_x':0.5,'center_y':0.45}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press : root.create_program_for_old_horizontal_machine(object)    
    MDRaisedButton:
        text: 'Back'
        pos_hint: {'center_x':0.6,'center_y':0.1}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: 
            root.manager.current = 'PinBoreScreen'   
            root.reset_old_horizontal_screen_fields()     
    MDRaisedButton:
        text: 'Home'
        pos_hint: {'center_x':0.4,'center_y':0.1}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: 
            root.manager.current = 'HomeScreen'        
            root.reset_old_horizontal_screen_fields() 
    MDRectangleFlatButton:
        id: LiveModeButton
        text: 'Live Mode'
        pos_hint: {'center_x':0.442,'center_y':0.28}
        font_size: "15sp"
        theme_text_color: "Custom"
        text_color: 0, 1, 0, 1
        line_color: 0, 1, 0, 1
        on_press : root.set_live_mode(object) 
    MDRectangleFlatButton:
        id: TestModeButton
        text: 'Test Mode'
        pos_hint: {'center_x':0.56,'center_y':0.28}     
        font_size: "15sp"    
        theme_text_color: "Custom"
        text_color: 1, 1, 1, 1
        line_color: 1, 1, 1, 1
        on_press : root.set_test_mode(object) 


<NewHorizontalScreen>:
    name: 'NewHorizontalScreen'
    MDLabel:
        text: 'New Horizontal Machine (127)'
        pos_hint: {'center_x':0.84,'center_y':0.8}
        font_size: '20sp'
        bold: True
        italic: True
        theme_text_color: "Primary"
    MDTextField:
        id: JobNumber
        text: self.text.upper() if self.text is not None else ''
        hint_text: "Enter Job Number"
        helper_text: "Job number MUST match with number in Spec."
        helper_text_mode: "on_focus"
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.60}
        size_hint_x:None
        width:300
        height:10      
    MDRaisedButton:
        text: 'Submit'
        pos_hint: {'center_x':0.5,'center_y':0.45}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press : root.create_program_for_new_horizontal_machine(object)
    MDRaisedButton:
        text: 'Back'
        pos_hint: {'center_x':0.6,'center_y':0.1}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: 
            root.manager.current = 'PinBoreScreen'   
            root.reset_new_horizontal_screen_fields() 
    MDRaisedButton:
        text: 'Home'
        pos_hint: {'center_x':0.4,'center_y':0.1}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: 
            root.manager.current = 'HomeScreen'        
            root.reset_new_horizontal_screen_fields()     


<SettingScreen>:
    name: 'SettingScreen'
    MDLabel:
        text: 'Setting'
        pos_hint: {'center_x':0.93,'center_y':0.9}
        font_size: '36sp'
        bold: True
        italic: True
        theme_text_color: "Primary"
    MDRaisedButton:
        text: 'Pin Bore Setting'
        pos_hint: {'center_x':0.5,'center_y':0.7}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: root.manager.current = 'PinBoreSettingScreen'
    MDRaisedButton:
        text: 'User Setting'
        pos_hint: {'center_x':0.9,'center_y':0.1}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: root.manager.current = 'UserSettingScreen'    
    MDRaisedButton:
        text: 'Application Setting'
        pos_hint: {'center_x':0.1,'center_y':0.1}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press : 
            root.admin_check() 
    MDRaisedButton:
        text: 'Back'
        pos_hint: {'center_x':0.5,'center_y':0.1}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: root.manager.current = 'HomeScreen'   
        

<PinBoreSettingScreen>:
    name: 'PinBoreSettingScreen'
    MDLabel:
        text: 'Pin Bore Setting'
        pos_hint: {'center_x':0.86,'center_y':0.9}
        font_size: '36sp'
        bold: True
        italic: True
        theme_text_color: "Primary"
    MDRaisedButton:
        text: 'Old Horizontal Setting Screen'
        pos_hint: {'center_x':0.5,'center_y':0.7}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: root.manager.current = 'OldHorizontalSettingScreen'
    MDRaisedButton:
        text: 'Home'
        pos_hint: {'center_x':0.4,'center_y':0.1}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: root.manager.current = 'HomeScreen'
    MDRaisedButton:
        text: 'Back'
        pos_hint: {'center_x':0.6,'center_y':0.1}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: root.manager.current = 'SettingScreen'


<OldHorizontalSettingScreen>:
    name: 'OldHorizontalSettingScreen'
    MDLabel:
        text: 'Old Horizontal Setting Screen'
        pos_hint: {'center_x':0.77,'center_y':0.9}
        font_size: '32sp'
        bold: True
        italic: True
        theme_text_color: "Primary"
    MDTextField:
        id: HorizontalTemplate
        hint_text: "Horizontal Template Path"
        text: "H:\CNC_Programming\Moe A\WisecoApplications\TestStuff\HorizontalTemplate 10-14-21.MIN"
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.75}
        size_hint_x:None
        width:800
        height:50       
    MDTextField:
        id: HorizontalToolList
        hint_text: "Horizontal Tool List Excel File Path"
        text: "H:\CNC_Programming\Moe A\WisecoApplications\TestStuff\HORIZONTAL_SHEETS_FOR_AUTOMATION.xlsx"
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.65}
        size_hint_x:None
        width:800
        height:50          
    MDTextField:
        id: ProbePrograms
        hint_text: "Probe Programs File Path"
        text: "H:\CNCProgs\HOREBORE\Probe Programs"
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.55}
        size_hint_x:None
        width:800
        height:50              
    MDTextField:
        id: RunningFolderPathOfOldHorizontalMachine
        hint_text: "Horizontal Programs (Running Folder) Path"
        text: "H:\CNC_Programming\Moe A\WisecoApplications\TestStuff\Horizontal"
        helper_text: "Folder that Use on Machine to Load the Program."
        helper_text_mode: "on_focus"
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.45}
        size_hint_x:None
        width:800
        height:50
    MDTextField:
        id: OriginalFolderPathOfOldHorizontalMachine
        hint_text: "Horizontal Programs (Original Folder) Path"
        text: "H:\CNC_Programming\Moe A\WisecoApplications\TestStuff\Horizontal Original"
        helper_text: "Folder that Use as Backup For Programs."
        helper_text_mode: "on_focus"
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.35}
        size_hint_x:None
        width:800
        height:50            
    MDTextField:
        id: RunningFolderPathOfOldHorizontalMachineTestMode
        hint_text: "Horizontal Programs (Running Folder) Path >> Test Mode"
        text: "H:\CNC_Programming\Moe A\WisecoApplications\TestStuff\Test_Mode_Horizontal"
        helper_text: "Test Mode >> Folder that Use on Machine to Load the Program."
        helper_text_mode: "on_focus"
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.25}
        size_hint_x:None
        width:800
        height:50     
    MDTextField:
        id: OriginalFolderPathOfOldHorizontalMachineTestMode
        hint_text: "Horizontal Programs (Original Folder) Path >> Test Mode"
        text: "H:\CNC_Programming\Moe A\WisecoApplications\TestStuff\Test_Mode_Horizontal_Original"
        helper_text: "Test Mode >> Folder that Use as Backup For Programs."
        helper_text_mode: "on_focus"
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.15}
        size_hint_x:None
        width:800
        height:50                             
    MDRaisedButton:
        text: 'Home'
        pos_hint: {'center_x':0.4,'center_y':0.05}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: root.manager.current = 'HomeScreen'
    MDRaisedButton:
        text: 'Back'
        pos_hint: {'center_x':0.6,'center_y':0.05}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: root.manager.current = 'PinBoreSettingScreen'      


<UserSettingScreen>:
    name: 'UserSettingScreen'
    MDLabel:
        text: 'User Setting'
        pos_hint: {'center_x':0.92,'center_y':0.9}
        font_size: '32sp'
        bold: True
        italic: True
        theme_text_color: "Primary"
    MDTextField:
        id: EmailPassword
        hint_text: "Enter Email Password"
        helper_text: "Email Password Probably be Same as Computer Login Password."
        helper_text_mode: "on_focus"
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        password: True
        pos_hint: {'center_x': 0.50, 'center_y': 0.75}
        size_hint_x:None
        width:800
        height:10  
    MDTextField:
        id: CimcoEditorPath
        hint_text: "CIMCO Editor Path"
        text: "C:\CIMCO\CIMCOEdit8\CIMCOEdit.EXE"
        helper_text: "Path Should be Where CIMCO App Installed in User Computer."
        helper_text_mode: "on_focus"
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.60}
        size_hint_x:None
        width:800
        height:50                              
    MDRaisedButton:
        text: 'Home'
        pos_hint: {'center_x':0.4,'center_y':0.1}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: root.manager.current = 'HomeScreen'
    MDRaisedButton:
        text: 'Back'
        pos_hint: {'center_x':0.6,'center_y':0.1}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: root.manager.current = 'SettingScreen'


<AppSettingScreen>:
    name: 'AppSettingScreen'
    MDLabel:
        text: 'Application Setting'
        pos_hint: {'center_x':0.87,'center_y':0.9}
        font_size: '32sp'
        bold: True
        italic: True
        theme_text_color: "Primary"
    MDTextField:
        id: EmailAddressList
        hint_text: "Email Address List File Path"
        text: "H:\CNC_Programming\Moe A\WisecoApplications\TestStuff\EMAIL_ADDRESS_LIST.xlsx"
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.75}
        size_hint_x:None
        width:800
        height:50   
    MDTextField:
        id: TrelloEmailAddress
        hint_text: "Trello Board Email Address"
        text: "moemenalatweh1+sqa4wcni54jz6erwnpbj@boards.trello.com"                   
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.65}
        size_hint_x:None
        width:800
        height:50       
    MDTextField:
        id: DatabaseServer
        hint_text: "Server of The Database"
        text: "us-men-app-sql1"                   
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.55}
        size_hint_x:None
        width:800
        height:50           
    MDTextField:
        id: DatabaseName
        hint_text: "Database Name"
        text: "EngineWorx"                   
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.45}
        size_hint_x:None
        width:800
        height:50               
    MDRaisedButton:
        text: 'Add New User'
        pos_hint: {'center_x':0.5,'center_y':0.25}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: root.manager.current = 'AddNewUserScreen'    
    MDRaisedButton:
        text: 'Home'
        pos_hint: {'center_x':0.4,'center_y':0.1}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: root.manager.current = 'HomeScreen'
    MDRaisedButton:
        text: 'Back'
        pos_hint: {'center_x':0.6,'center_y':0.1}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: root.manager.current = 'SettingScreen' 

                                                                
<AddNewUserScreen>:
    name: 'AddNewUserScreen'
    MDLabel:
        text: 'Add New User'
        pos_hint: {'center_x':0.90,'center_y':0.9}
        font_size: '32sp'
        bold: True
        italic: True
        theme_text_color: "Primary"   
    MDTextField:
        id: UserName
        hint_text: "Enter User Name"
        helper_text: "First Name, Last Name."
        helper_text_mode: "on_focus"
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.65}
        size_hint_x:None
        width:400
        height:10          
    MDTextField:
        id: NewRWBEmail
        hint_text: "Enter RWB Email Address"
        helper_text: "Email With Extension @rwbteam."
        helper_text_mode: "on_focus"
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.55}
        size_hint_x:None
        width:400
        height:10
    MDTextField:
        id: NewWisecoEmail
        hint_text: "Enter Wiseco Email Address"
        helper_text: "Email With Extension @wiseco, If Not Applicable. Leave It Blank."
        helper_text_mode: "on_focus"
        color_mode: 'custom'
        line_color_focus: 1, 1, 1, 1
        pos_hint: {'center_x': 0.50, 'center_y': 0.45}
        size_hint_x:None
        width:400
        height:10                  
    MDRaisedButton:
        text: 'SING UP'
        pos_hint: {'center_x':0.5,'center_y':0.30}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press : 
            root.add_new_user()        
    MDRaisedButton:
        text: 'Home'
        pos_hint: {'center_x':0.4,'center_y':0.1}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: 
            root.manager.current = 'HomeScreen'
            root.reset_new_user_screen_fields() 
    MDRaisedButton:
        text: 'Back'
        pos_hint: {'center_x':0.6,'center_y':0.1}
        md_bg_color: 120/255, 0/255, 0/255, 1
        font_size: "15sp"
        on_press: 
            root.manager.current = 'AppSettingScreen'  
            root.reset_new_user_screen_fields() 
     
"""


# endregion <<<<=======================================[Screen Builder KV]========================================>>>>


# region <<<<===========================================[Login Screen]===========================================>>>>

class LoginScreen(Screen):
    # ========================================================================|
    #  Create Function to define the APP, and show APP's Version and Features.|
    # ========================================================================|
    def application_version_features(self):
        print("(application_version_features) Function >> called")
        application_version_features_list = \
            ["An App to create CNC programs according to the type of operation and machine. ",
             "[color=ff1a1a]Version :[/color] 1.0.0"
             + "[color=ff1a1a]                                                                   "
               "Release Date :[/color] 06/28/2021",
             "--------------------------------------------------------------------------"
             "-----------------------------------------------------------------",
             "[color=ff1a1a]Features :[/color]", "[#] Create Pin Bore programs for Horizontal machines (28,29,32). "]
        close_button = MDRaisedButton(text='Close', on_release=self.close_login_screen_window, font_size=16)
        self.login_screen_message_window = MDDialog(title='[b][color=ffffff]Wiseco Programs Maker App[/color][/b]',
                                                    text=('[color=ffffff]' + '\n'.join(
                                                        application_version_features_list) + '[/color]'),
                                                    size_hint=(0.7, 1.0), buttons=[close_button], auto_dismiss=False)
        self.login_screen_message_window.open()

    # ==================================================================|
    #  Create Function to Set and Manege User Info and Password of Login|
    # ==================================================================|
    def login_check(self):
        print("(login_check) Function >> called")
        # ========================================================================================|
        # [#] To set and manege UserName and Password of Users we create Excel Sheet with name of |
        #     (EMAIL_ADDRESS_LIST) to contain Login Information of UserName and Password.         |
        # [#] Use (try/except) Blocks to Handle Error of not finding or accessing the Excel Sheet.|
        # ========================================================================================|
        try:
            # ============================================================================================|
            # [#] Define a Variable to set path of the Excel Sheet (EMAIL_ADDRESS_LIST).                  |
            # [#] Access MDTextField of (id: EmailAddressList) in (AppSettingScreen) from Screens_Builder |
            #     to get the File Path.                                                                   |
            # ============================================================================================|
            global email_address_list_file_path
            email_address_list_file_path = self.manager.get_screen('AppSettingScreen').ids["EmailAddressList"].text

            # ====================================================================|
            # [#] Define a Variable to read the Excel Sheet (EMAIL_ADDRESS_LIST). |
            # [#] Use Pandas library to read the Excel Sheet.                     |
            # [#] Use (sheet_name=None) to read all sheets inside the Excel Sheet.|
            # ====================================================================|
            global email_address_list_file
            email_address_list_file = pd.read_excel(email_address_list_file_path, sheet_name=None)

            # =======================================================================================|
            # [#] Define a List to store Users Names from the Excel Sheet.                           |
            # [#] Use for loop to iterate through Column of ['Users'] inside sheet of name ['Email'].|
            # [#] Add each Row (Users Name) to the List.                                             |
            # =======================================================================================|
            global users_name_list
            users_name_list = []
            for user in email_address_list_file['Email']['Users']:
                users_name_list.append(user)

            # ====================================================================================================|
            # [#] Define a List to store Users Email from the Excel Sheet.                                        |
            # [#] Use for loop to iterate through Column of ['Work Email Address'] inside sheet of name ['Email'].|
            # [#] Add each Row (Users Email) to the List.                                                         |
            # ====================================================================================================|
            global users_email_address_list
            users_email_address_list = []
            for user in email_address_list_file['Email']['Work Email Address']:
                users_email_address_list.append(user)

            # ===================================================================================================|
            # [#] Define a List to store Trello Users Names from the Excel Sheet.                                |
            # [#] Use for loop to iterate through Column of ['Trello Users Name'] inside sheet of name ['Email'].|
            # [#] Add each Row (Trello Users Name) to the List.                                                  |
            # ===================================================================================================|
            global trello_users_name_list
            trello_users_name_list = []
            for user in email_address_list_file['Email']['Trello Users Name']:
                trello_users_name_list.append(user)

            # ======================================================================================================|
            # [#] Define a List to store Guests Email from the Excel Sheet.                                         |
            # [#] Use for loop to iterate through Column of ['Guests Email Address'] inside sheet of name ['Email'].|
            # [#] Add each Row (Guests Email) to the List.                                                          |
            # [#] Guests Users can access and see the app , but can't make programs                                 |
            # ======================================================================================================|
            global guests_email_address_list
            guests_email_address_list = []
            for user in email_address_list_file['Email']['Guests Email Address']:
                guests_email_address_list.append(user)

            # =====================================|
            # Define a List to store the Passwords.|
            # =====================================|
            global password_list
            password_list = []
            for password in email_address_list_file['Email']['Pass']:
                password_list.append(password)

        # =======================================================================================================|
        # [#] Use (except) block to Handle Error of not finding or accessing the Excel Sheet and avoid APP crash.|
        # [#] Use (return) to stop the function (APP as a result) from executing the rest of the code.           |
        # =======================================================================================================|
        except Exception as error:
            close_button = MDRaisedButton(text='Close', on_release=self.close_login_screen_window, font_size=16)
            self.login_screen_message_window = MDDialog(title='[color=990000]Warning Message[/color]',
                                                        text=("[color=ffffff]Failed to Find, Load, or Access" +
                                                              '[b][u][color=ffffff] Email Address List [/color][/u][/b]'
                                                              + "\n" + "An Error has occurred :[/color]" + "\n" +
                                                              '[color=ff1a1a]' + str(error) + '[/color]' + "\n" +
                                                              "[color=ffffff]Double Check Network, "
                                                              "and File Location.[/color]"),
                                                        size_hint=(0.7, 1.0), buttons=[close_button],
                                                        auto_dismiss=False)
            self.login_screen_message_window.open()
            return

        # ===============================================================================================|
        # [#] Define a Variable to get the email address that user input.                                |
        # [#] Get the text of the MDTextField of id ["Email"] in (AppSettingScreen) from Screens_Builder.|
        # ===============================================================================================|
        global user_email_address
        user_email_address = self.ids["Email"].text

        # =========================================================================================================|
        # [#] Check Login Information by check if the inputs (Email and Password) that user entered is in the      |
        #      lists that getting from Excel sheet.                                                                |
        # [#] Move to 'HomeScreen' of APP if info is correct by passing screen name after (self.manager.current =).|
        # [#] Open Window to show message of wrong Login info if user entered inputs are not in lists that         |
        #     getting from Excel sheet.                                                                            |
        # =========================================================================================================|
        if (((self.ids["Email"].text in users_email_address_list and
              self.ids["Email"].text != 'malatweh@rwbteam.com')) and ((self.ids["Password"].text) in password_list)):
            self.manager.current = 'HomeScreen'
        elif (self.ids["Email"].text == 'malatweh@rwbteam.com' and
              (self.ids["Password"].text == 'moe' + password_list[6])):
            self.manager.current = 'HomeScreen'
        elif (self.ids["Email"].text in guests_email_address_list and
              ((self.ids["Password"].text) in password_list)):
            self.manager.current = 'HomeScreen'
        else:
            close_button = MDRaisedButton(text='Close', on_release=self.close_login_screen_window, font_size=16)
            self.login_screen_message_window = MDDialog(title='[color=990000]Warning Message[/color]', text=(
                '[color=ffffff]Wrong Email or Password, Try Again[/color]'), size_hint=(0.7, 1.0),
                                                        buttons=[close_button], auto_dismiss=False)
            self.login_screen_message_window.open()

    # ===============================================================================|
    #  Create Function to close screen message window when User click on Close Button|
    # ===============================================================================|
    def close_login_screen_window(self, obj):
        print("(close_login_screen_window) Function >> called")
        self.login_screen_message_window.dismiss()

# endregion <<<<==========================================[Login Screen]==========================================>>>>


# region <<<<===========================================[Home Screen]============================================>>>>

class HomeScreen(Screen):
    # ===============================================================================================================|
    # [#] Create Function that will be activated once enter the 'HomeScreen' to call other function (just by entering|
    #     the screen without click on Buttons), needs to add it to 'HomeScreen' on Screens_Builder as well.          |
    # ===============================================================================================================|
    def on_pre_enter(self):
        print("(on_pre_enter) Function >> called")
        # ====================================================================|
        # Call Function to set User info and display UserName on 'HomeScreen'.|
        # ====================================================================|
        self.set_user_name()

        # ===================================================|
        # Call Function to Connect the APP with the DataBase.|
        # ===================================================|
        self.database_connect()

        # ============================================================================================|
        # [#] Define Variable to set Cimco App path by enter 'HomeScreen' to use it for all operations|
        #     and machines to open the CNC programs after create them.                                |
        # [#] Access MDTextField of (id: CimcoEditorPath) in (UserSettingScreen) from Screens_Builder |
        #     to get Cimco application Path.                                                          |
        # ============================================================================================|
        global cimco_editor_path
        cimco_editor_path = self.manager.get_screen('UserSettingScreen').ids["CimcoEditorPath"].text

    # ============================================|
    # Create Function to Set and Manege User info.|
    # ============================================|
    def set_user_name(self):
        print("(set_user_name) Function >> called")
        # ==========================================================================================================|
        # [#] Steps To display UserName on 'HomeScreen':                                                            |
        #     -First, check if Email used to enter the app is in Lists of Users Email that getting from Excel Sheet.|
        #     -Find index (location) of user email in the list of 'users_email_address_list' to use it next.        |
        #     -Use the index to get the name of user from 'users_name_list' and set it as text of                   |
        #      the MDLabel of id ["UserName"] in (HomeScreen) in Screens_Builder to display name in'HomeScreen'.    |
        # [#] 'users_email_address_list' : List of Users can Access the APP and create programs.                    |
        # [#] 'guests_email_address_list' : List of Users can Access the APP only without create programs.          |
        # ==========================================================================================================|

        # =========================================================|
        # Define Variable to Set current user name who use the app.|
        # =========================================================|
        global connected_user_name

        # ========================================|
        # Define Variable to Set Trello user name.|
        # ========================================|
        global trello_user_name

        # ==========================================================================================================|
        # Access of MDTextField of (id: Email) in (LoginScreen) from Screens_Builder to check if it is in the lists.|
        # ==========================================================================================================|
        if (self.manager.get_screen('LoginScreen').ids["Email"].text in users_email_address_list):
            # =====================================|
            # Find index of user email in the list.|
            # =====================================|
            user_index = users_email_address_list.index(self.manager.get_screen('LoginScreen').ids["Email"].text)

            # ================================================================================================|
            # Get UserName from the list and set it as text of the MDLabel of id ["UserName"] in (HomeScreen).|
            # ================================================================================================|
            self.ids["UserName"].text = users_name_list[user_index]

            # ====================================================|
            # Store the user name in the variable to use it later.|
            # ====================================================|
            connected_user_name = self.ids["UserName"].text

            # ============================================================================|
            # Get the trello user name from the list by passing the index to use it later.|
            # ============================================================================|
            trello_user_name = trello_users_name_list[user_index]

        # ====================================================|
        # Same Steps above but in 'guests_email_address_list'.|
        # ====================================================|
        elif (self.manager.get_screen('LoginScreen').ids["Email"].text in guests_email_address_list):
            user_index = guests_email_address_list.index(self.manager.get_screen('LoginScreen').ids["Email"].text)
            self.ids["UserName"].text = (users_name_list[user_index])
            connected_user_name = self.ids["UserName"].text
            trello_user_name = trello_users_name_list[user_index]

    # ============================================================================================|
    # [#] Create Function to Connect the APP with the DataBase.                                   |
    # [#] Use (try/except) Blocks to Handle any Error may occur when connecting with the DataBase.|
    # ============================================================================================|
    def database_connect(self):
        print("(database_connect) Function >> called")

        # =========================================================================================|
        # [#] Define Variable to set the Server of the DataBase.                                   |
        # [#] Access MDTextField of (id: DatabaseServer) in (AppSettingScreen) from Screens_Builder|
        #     to get the Server of the DataBase.                                                   |
        # =========================================================================================|
        global engine_worx_database_server
        engine_worx_database_server = self.manager.get_screen('AppSettingScreen').ids["DatabaseServer"].text

        # =======================================================================================|
        # [#] Define Variable to set the DataBase Name.                                          |
        # [#] Access MDTextField of (id: DatabaseName) in (AppSettingScreen) from Screens_Builder|
        #     to get the DataBase Name.                                                          |
        # =======================================================================================|
        global engine_worx_database_name
        engine_worx_database_name = self.manager.get_screen('AppSettingScreen').ids["DatabaseName"].text

        try:
            # ===================================================================================================|
            # [#] Use (pyodbc) Library to connect the APP with the DataBase and Manege Data.                     |
            # [#] To be Able to access the DataBase, Needs permission and DataBase info From IT Department.      |
            # [#] Information needed to connect with the DateBase:                                               |
            #     - Type of Driver which is : {SQL Server}                                                       |
            #     - Name of Server which is : us-men-app-sql1                                                    |
            #     - Name of DataBase which is : EngineWorx                                                       |
            #     - User Authentication : Set it as 'Yes' while Authentication Information are Same as UserLogin |
            #                             Info for the Computer (Windows Authentication), If They are different  |
            #                             or set it to use (SQL Authentication) needs to add:                    |
            #                             ('Uid=WISECOMANF\\DomainUser;') and ('Pwd='UserPasswordToSQL;') with   |
            #                             ('Trusted_Connection=No;').                                            |
            # [#] Define Variable to create connection with DataBase by using DataBase Info.                     |
            # ===================================================================================================|
            global engine_worx_database_connect
            engine_worx_database_connect = pyodbc.connect('Driver={SQL Server};'
                                                          'Server=' + engine_worx_database_server + ';'
                                                          'Database=' + engine_worx_database_name + ';'
                                                          'Trusted_Connection=Yes;')

            # ================================================================================|
            # Define Variable to Create 'Cursor' to point and locate the data inside DataBase.|
            # ================================================================================|
            global engine_worx_database_cursor
            engine_worx_database_cursor = engine_worx_database_connect.cursor()

        # =====================================================================================================|
        # [#] Use (except) block to Handle any Error may occur when accessing the DataBase and avoid APP crash.|
        # [#] Use (return) to stop the function (APP as a result) from executing the rest of the code.         |
        # =====================================================================================================|
        except Exception as error:
            close_button = MDRaisedButton(text='Close', on_release=self.close_home_screen_window, font_size=16)
            self.home_screen_window = MDDialog(title='[color=990000]Warning Message[/color]',
                                               text=("[color=ffffff]Failed to Connect or Access" +
                                                     '[b][u][color=ffffff] EngineWorx [/color][/u][/b]'
                                                     + "DataBase." + "\n" + "An Error has occurred :[/color]" +
                                                     "\n" + '[color=ff1a1a]' + str(error) + '[/color]' + "\n" +
                                                     "[color=ffffff]Double Check Network and your DataBase "
                                                     "Authentication.[/color]"),
                                               size_hint=(0.7, 1.0), buttons=[close_button], auto_dismiss=False)
            self.home_screen_window.open()
            return

    # ==========================================================|
    # Create Function to let User be able to logout from the APP|
    # ==========================================================|
    def logout(self):
        print("(logout) Function >> called")
        # =========================================================================================|
        #  [#] Move back to 'LoginScreen' of APP when user click on LOGOUT Button by passing screen|
        #     name after (self.manager.current =).                                                 |
        # =========================================================================================|
        self.manager.current = 'LoginScreen'

        # =====================================================================================================|
        # Reset Login Info (Email and Password) in 'LoginScreen' to let User be able to enter login info again.|
        # =====================================================================================================|
        self.manager.get_screen('LoginScreen').ids["Email"].text = ""
        self.manager.get_screen('LoginScreen').ids["Password"].text = ""

    # ===============================================================================|
    #  Create Function to close screen message window when User click on Close Button|
    # ===============================================================================|
    def close_home_screen_window(self, obj):
        print("(close_home_screen_window) Function >> called")
        self.home_screen_window.dismiss()


# endregion <<<<==========================================[Home Screen]===========================================>>>>


# region <<<<==================================[Load Horizontal Sheets Function]=================================>>>>

# ====================================================================================================|
# [#] Create Function to Load Excel Sheet of (HORIZONTAL_SHEETS_FOR_AUTOMATION) that contain all tools|
#     and some info that used to create PinBore Programs.                                             |
# [#] Creating Excel Sheet that's contain all tools for different machines to make it easier for User |
#     to add or update any tool or info without needing to update the code itself (in most cases, but |
#     some adjustments in Excel file required to update the code).                                    |
# [#] Use this Function to load tools list and info for old and new PoiBore Machines                  |
# ====================================================================================================|
def load_horizontal_machine_tool_list_sheets(self):
    print("(load_horizontal_machine_tool_list_sheets) Function >> called")
    # ========================================================================================|
    # [#] Define a Variable to set path of the Excel Sheet (HORIZONTAL_SHEETS_FOR_AUTOMATION).|
    # [#] Access MDTextField of (id: HorizontalToolList) in (OldHorizontalSettingScreen) from |
    #     Screens_Builder to get the File Path.                                               |
    # ========================================================================================|
    horizontal_tool_list_file_path = self.manager.get_screen(
        'OldHorizontalSettingScreen').ids["HorizontalToolList"].text

    # =================================================================================|
    # [#] Define a Variable to read the Excel Sheet (HORIZONTAL_SHEETS_FOR_AUTOMATION).|
    # [#] Use Pandas library to read the Excel Sheet.                                  |
    # [#] Use (sheet_name=None) to read all sheets inside the Excel Sheet.             |
    # =================================================================================|
    global horizontal_tool_list_file
    horizontal_tool_list_file = pd.read_excel(horizontal_tool_list_file_path, sheet_name=None)

    # ============================================================================================|
    # [#] Define a List to store Finish Bore Tool List from the Excel Sheet.                      |
    # [#] Use for loop to iterate through Column of ['PIN_BORE_DIAMETER'] inside sheet of         |
    #     name ['FINISH_BORE_TOOL_LIST'].                                                         |
    # [#] Add each Row (Pin Bore Diameter Size) to the List.                                      |
    # [#] ['FINISH_BORE_TOOL_LIST']: The sheet that contains the tool list of Finish Bore.        |
    # [#] ['PIN_BORE_DIAMETER']: The column that contains the Pin Bore Diameter Sizes.            |
    # [#] Choose Column of ['PIN_BORE_DIAMETER'] to use it as index to access the tool list later.|
    # ============================================================================================|
    global finish_bore_tool_list
    finish_bore_tool_list = []
    for tool in horizontal_tool_list_file['FINISH_BORE_TOOL_LIST']['PIN_BORE_DIAMETER']:
        finish_bore_tool_list.append(tool)

    # =========================================================================================|
    # [#] Define a List to store Rough Bore Tool List from the Excel Sheet.                    |
    # [#] Use for loop to iterate through Column of ['DRILL_DIAMETER'] inside sheet of         |
    #     name ['ROUGH_BORE_TOOL_LIST'].                                                       |
    # [#] Add each Row (Drill Diameter Size) to the List.                                      |
    # [#] ['ROUGH_BORE_TOOL_LIST']: The sheet that contains the tool list of Rough Bore.       |
    # [#] ['DRILL_DIAMETER']: The column that contains the Drill Diameter Sizes.               |
    # [#] Choose Column of ['DRILL_DIAMETER'] to use it as index to access the tool list later.|
    # =========================================================================================|
    global rough_bore_tool_list
    rough_bore_tool_list = []
    for tool in horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST']['DRILL_DIAMETER']:
        rough_bore_tool_list.append(tool)

    # =====================================================================================================|
    # [#] Define a List to store LockRing and C/Fren Tool List from the Excel Sheet.                       |
    # [#] Use for loop to iterate through Column of ['TOOL_WIDTH'] inside sheet of                         |
    #     name ['LOCK_RING_AND_CFREN_TOOL_LIST'].                                                          |
    # [#] Add each Row (LockRing or C/Fren Width) to the List.                                             |
    # [#] ['LOCK_RING_AND_CFREN_TOOL_LIST']: The sheet that contains the tool list of LockRing and C/Fren. |
    # [#] ['TOOL_WIDTH']: The column that contains the LockRing or C/Fren Width.                           |
    # [#] Choose Column of ['TOOL_WIDTH'] to use it as index to access the tool list later.                |
    # =====================================================================================================|
    global lock_ring_and_cfren_tool_list
    lock_ring_and_cfren_tool_list = []
    for tool in horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST']['TOOL_WIDTH']:
        lock_ring_and_cfren_tool_list.append(tool)

    # =============================================================================================|
    # [#] Define a List to store Miscellaneous Tool List from the Excel Sheet.                     |
    # [#] Use for loop to iterate through Column of ['TOOL_USAGE'] inside sheet of                 |
    #     name ['MISCELLANEOUS_TOOL_LIST'].                                                        |
    # [#] Add each Row (Tool Usage) to the List.                                                   |
    # [#] ['MISCELLANEOUS_TOOL_LIST']: The sheet that contains the tool list of not specific tools.|
    # [#] ['TOOL_USAGE']: The column that contains the Usage of the tools.                         |
    # [#] Choose Column of ['TOOL_USAGE'] to use it as index to access the tool list later.        |
    # =============================================================================================|
    global miscellaneous_tool_list
    miscellaneous_tool_list = []
    for tool in horizontal_tool_list_file['MISCELLANEOUS_TOOL_LIST']['TOOL_USAGE']:
        miscellaneous_tool_list.append(tool)

    # ============================================================================================|
    # [#] Define a List to store Horizontal Slot Numbers(i,j, Radius) from the Excel Sheet.       |
    # [#] Use for loop to iterate through Column of ['PIN_BORE_DIAMETER'] inside sheet of         |
    #     name ['HORIZONTAL_SLOT_NUMBERS'].                                                       |
    # [#] Add each Row (Pin Bore Diameter Size) to the List.                                      |
    # [#] ['HORIZONTAL_SLOT_NUMBERS']: The sheet that contains the horiz Numbers(i,j, Radius).    |
    # [#] ['PIN_BORE_DIAMETER']: The column that contains the Pin Bore Diameter Sizes that        |
    #     have specific horiz Numbers(i,j, Radius) foe each size.                                 |
    # [#] Choose Column of ['PIN_BORE_DIAMETER'] to use it as index to access the tool list later.|
    # ============================================================================================|
    global horizontal_slot_numbers
    horizontal_slot_numbers = []
    for tool in horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS']['PIN_BORE_DIAMETER']:
        horizontal_slot_numbers.append(tool)

    # =============================================================================================================|
    # [#] Define a List to store Forged Forging Numbers from the Excel Sheet.                                      |
    # [#] Use for loop to iterate through Column of ['FORGING_NUMBER'] inside sheet of name ['FORGED_FORGING'].    |
    # [#] Add each Row (Forging Number) to the List.                                                               |
    # [#] ['FORGED_FORGING']: The sheet that contains the Forging Number that needs special Notes in horiz Program.|
    # [#] ['FORGING_NUMBER']: The column that contains the Forging Number of Forged Forging.                       |
    # [#] Choose Column of ['FORGING_NUMBER'] to use it as index to access the tool list later.                    |
    # =============================================================================================================|
    global forged_forging_list
    forged_forging_list = []
    for forging in horizontal_tool_list_file['FORGED_FORGING']['FORGING_NUMBER']:
        forged_forging_list.append(forging)

    # ================================================================================================================|
    # [#] Define a List to store all DoubleOilHolesSlots pin sizes (that used in horiz template) from the Excel Sheet.|
    # [#] Use this List to check if Pin Bore Diameter Size include in the logic of Horiz template of                  |
    #     DoubleOilHolesSlots, otherwise inform the user to add new DoubleOilHolesSlots Numbers to Horiz template.    |
    # [#] Use for loop to iterate through Column of ['PIN_BORE_DIAMETER'] inside sheet of name ['DOHS_PIN_SIZES'].    |
    # [#] Add each Row (Pin Bore Diameter Size) to the List.                                                          |
    # [#] ['DOHS_PIN_SIZES']: The sheet that contains the Pin Bore Diameter Sizes that included in the                |
    #     logic of Horiz template of DoubleOilHolesSlots.                                                             |
    # [#] ['PIN_BORE_DIAMETER']: The column that contains the DoubleOilHolesSlots Pin Sizes that                      |
    #     used in horiz template.                                                                                     |
    # [#] Choose Column of ['PIN_BORE_DIAMETER'] to use it as index to access the tool list later.                    |
    # ================================================================================================================|
    global double_oil_hole_slots_pin_sizes_list
    double_oil_hole_slots_pin_sizes_list = []
    for pin_size in horizontal_tool_list_file['DOHS_PIN_SIZES']['PIN_BORE_DIAMETER']:
        double_oil_hole_slots_pin_sizes_list.append(pin_size)

    # =====================================================================================|
    # ***Still Need to back and add Manual Horiz Slots that used in Horiz Template***      |
    #                                     <><><>                                           |
    # =====================================================================================|

# endregion <<<<=================================[Load Horizontal Sheets Function]================================>>>>


# region <<<<===================================[Four Cycle Pin Bore Function]===================================>>>>

# =====================================================================================|
#  Create Function to set all 4-Cycle PinBore Variables and get them from the DataBase.|
# =====================================================================================|
def four_cycle_pin_bore_variables(self):
    print("(four_cycle_pin_bore_variables) Function >> called")

    # region <<<<============================[Piston Information]=============================>>>>
    # =============================================================================================|
    # [#] Set all Variable that related on Piston Info and get them from the DataBase.             |
    # [#] Use (try/except) Blocks to Handle any Error may occur when connecting with the DataBase. |
    # [#] Use 'piston_id' (that getting in 'create_program_for_old_horizontal_machine' Function) to|
    #     find the data for each Variable                                                          |
    # [#] piston_id: is the unique ID for each Job (Piston) stored in the DataBase.                |
    # =============================================================================================|
    try:
        print("try BLOCK in (four_cycle_pin_bore_variables) Function <Spec Info>")

        # region  <<<<============================[Job Released Status]============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Job Released Status to check if the job is ready to program or not. |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston                                                                    |
        #     Column Name : Released_Y_N                                                                 |
        # ===============================================================================================|
        global job_released_status
        engine_worx_database_cursor.execute(
            'SELECT Released_Y_N FROM SpexPiston WHERE PistonID = ?', piston_id)

        # ===============================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                              |
        # [#] Set the value as a "None" if no data found, otherwise set the data as it is stored.        |
        # [#] Expected Data output Type : 'String'                                                       |
        # [#] Expected Data output value : 'YES' or 'NO'                                                 |
        # ===============================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if (data is None):
                job_released_status = "None"
            else:
                job_released_status = data
        print("[#]job_released_status FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", job_released_status)

        # ==================================================================================================|
        # [#] Check Job Released Status:                                                                    |
        #     if it is not released, call (failed_to_create_old_horizontal_machine_program) Function to warn|
        #     the user and stop running the code.                                                           |
        # [#] Add the message to the main Fail Messages and Email Messages to show them later.              |
        # [#] Use (return) to stop the function (APP as a result) from executing the rest of the code.      |
        # ==================================================================================================|
        if (job_released_status != "YES" and job_released_status != "Yes"):
            fail_messages_of_creating_old_horizontal_machine_program.append(
                "This Job Number " + '[b][u][color=0099ff]' + self.ids["JobNumberForOldHorizontalMachine"].text +
                '[/color][/u][/b]' + " is NOT Released" + " (Released Status: " + job_released_status + ")." +
                "\n" + "Work with Engineering to fix the Issue.")
            email_messages_of_creating_old_horizontal_machine_program.append(
                "This Job Number is NOT Released" + " (Released Status:" + job_released_status + ")" +
                "\n" + "Work with Engineering to fix the Issue." + "\n")
            failed_to_create_old_horizontal_machine_program(self)
            return

        # endregion  <<<<===========================[Job Released Status]===========================>>>>

        # region <<<<==============================[Pin Hole Diameter]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Pin Hole Diameter.                                                  |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : PistPinBoreDiameterIN                                                        |
        # ===============================================================================================|
        global pin_hole_diameter
        engine_worx_database_cursor.execute(
            'SELECT PistPinBoreDiameterIN FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ===============================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                              |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',       |
        #     otherwise, Round the Numerical Value for 4-Digits.                                         |
        # [#] Expected Data output Type : 'Numeric'                                                      |
        # [#] Expected Data output value : Numeric Value between [0.???? - 1.????]                       |
        # ===============================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                pin_hole_diameter = data
            else:
                pin_hole_diameter = round(float(data), 4)
        print("[#]pin_hole_diameter FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", pin_hole_diameter)

        # ====================================================================================================|
        # [#] pin_hole_diameter_verifying_status <Note>.                                                      |
        # [#] Some old jobs missing PinHoleDiameter value or sometimes it's forgotten by Engineers, because of|
        #     that needs to Define Variable to set Pin Hole Diameter Availability Status.                     |
        # [#] Set the Value to be 'False' as a default assuming the value is not missing.                     |
        # [#] Change the Value to be 'True' if it is missing to let user check it and corrected.              |
        # ====================================================================================================|
        global pin_hole_diameter_verifying_status
        pin_hole_diameter_verifying_status = False

        # endregion <<<<===========================[Pin Hole Diameter]=============================>>>>

        # region <<<<==============================[Pilot Diameter]==============================>>>>

        # region <<<<========================[Pilot Bore Availability Status]========================>>>>

        # ==============================================================================================|
        # [#] pilot_availability_status <Note>.                                                         |
        # [#] Define Variable to set Pilot Availability Status                                          |
        # [#] Set the Value to be "" as a default to use it in [VC119] Variable in Horizontal template. |
        # [#] Change the Value to be <1> when Pilot Diameter Detected, if status doesn't change, it will|
        #     ask user to choose the Pilot Diameter (Big or Small).                                     |
        # ==============================================================================================|
        global pilot_availability_status
        pilot_availability_status = ""

        # endregion <<<<=====================[Pilot Bore Availability Status]========================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Pilot Diameter.                                                     |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : PilotBoreDiameter                                                            |
        # ===============================================================================================|
        global pilot_diameter
        engine_worx_database_cursor.execute(
            'SELECT PilotBoreDiameter FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ===============================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                              |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',       |
        #     otherwise, Round the Numerical Value for 4-Digits.                                         |
        # [#] Expected Data output Type : 'Numeric'                                                      |
        # [#] Expected Data output value : Numeric Value: <2.2500> or <1.7000>                           |
        # ===============================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                pilot_diameter = data
            else:
                pilot_diameter = round(float(data), 4)
        print("[#]pilot_diameter FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", pilot_diameter)

        # endregion <<<<============================[Pilot Diameter]============================>>>>

        # region <<<<==============================[Pilot Bore Depth]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Pilot Bore Depth.                                                   |
        # [#] Define Variables to set PilotBoreDepthToDeck and PilotBoreDepthToDome to use them in logic |
        #     to Decide Pilot Bore Depth Value                                                           |
        # ===============================================================================================|
        global pilot_bore_depth
        global pilot_bore_depth_to_deck
        global pilot_bore_depth_to_dome

        # region <<<<==============================[Pilot Bore Depth To Deck]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Pilot Bore Depth To Deck.                                           |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : PilotBoreDepthToDeck                                                         |
        # ===============================================================================================|
        engine_worx_database_cursor.execute(
            'SELECT PilotBoreDepthToDeck FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ===============================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                              |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',       |
        #     otherwise, Round the Numerical Value for 4-Digits.                                         |
        # [#] Expected Data output Type : 'Numeric'                                                      |
        # [#] Expected Data output value : Numeric Value > [0.????]                                      |
        # ===============================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                pilot_bore_depth_to_deck = data
            else:
                pilot_bore_depth_to_deck = round(float(data), 4)
        print("[#]pilot_bore_depth TO DECK FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", pilot_bore_depth_to_deck)

        # endregion <<<<==============================[Pilot Bore Depth To Deck]==============================>>>>

        # region <<<<==============================[Pilot Bore Depth To Dome]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Pilot Bore Depth To Dome.                                           |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : PilotBoreDepthToDome                                                         |
        # ===============================================================================================|
        engine_worx_database_cursor.execute(
            'SELECT PilotBoreDepthToDome FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ===============================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                              |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',       |
        #     otherwise, Round the Numerical Value for 4-Digits.                                         |
        # [#] Expected Data output Type : 'Numeric'                                                      |
        # [#] Expected Data output value : Numeric Value > [0.????]                                      |
        # ===============================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                pilot_bore_depth_to_dome = data
            else:
                pilot_bore_depth_to_dome = round(float(data), 4)
        print("[#]pilot_bore_depth TO DOME FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", pilot_bore_depth_to_dome)

        # endregion <<<<==============================[Pilot Bore Depth To Dome]==============================>>>>

        # ===========================================================================================|
        # [#] pilot_bore_depth_verifying_status <Note>.                                              |
        # [#] Some old jobs missing PilotBoreDepth value or have it wrong with the correct value     |
        #     is locate in the 'Legacy Comments', because of that needs to Define Variable to set    |
        #     Pin Hole Diameter Verifying Status to check the value when have Legacy Comments.       |
        # [#] Set the Value to be 'False' as a default assuming the value is not missing and correct.|
        # [#] Change the Value to be 'True' After User check and verify the value.                   |
        # ===========================================================================================|
        global pilot_bore_depth_verifying_status
        pilot_bore_depth_verifying_status = False

        # endregion <<<<============================[Pilot Bore Depth]============================>>>>

        # region <<<<==============================[Pilot To Pin]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Pilot to Pin.                                                       |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : PilotBorePilotToPin                                                          |
        # ===============================================================================================|
        global pilot_to_pin
        engine_worx_database_cursor.execute(
            'SELECT PilotBorePilotToPin FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ===============================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                              |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',       |
        #     otherwise, Round the Numerical Value for 4-Digits.                                         |
        # [#] Expected Data output Type : 'Numeric'                                                      |
        # [#] Expected Data output value : Numeric Value < [0.????],ie: (Negative value)                 |
        # ===============================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                pilot_to_pin = data
            else:
                pilot_to_pin = round(float(data), 4)
        print("[#]pilot_to_pin FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", pilot_to_pin)

        # ===========================================================================================|
        # [#] pilot_to_pin_verifying_status <Note>
        # [#] Some old jobs missing PilotToPin value or have it wrong with the correct value is      |
        #     locate in the 'Legacy Comments', because of that needs to Define Variable to set       |
        #     Pilot To Pin Verifying Status to check the value when have Legacy Comments.            |
        # [#] Set the Value to be 'False' as a default assuming the value is not missing and correct.|
        # [#] Change the Value to be 'True' After User check and verify the value.                   |
        # ===========================================================================================|
        global pilot_to_pin_verifying_status
        pilot_to_pin_verifying_status = False

        # endregion <<<<============================[Pilot To Pin]============================>>>>

        # region <<<<======================[X_Distance From Origin To Pin Center]=========================>>>>

        # ================================================================================|
        # [#] Define Variable to set the X_distance from Piston Origin to Pin Hole Center.|
        # [#] The Value will be Calculated by Math Later.                                 |
        # ================================================================================|
        global X_distance_from_origin_to_pin_center
        X_distance_from_origin_to_pin_center = 0

        # endregion <<<<=====================[X_Distance From Origin To Pin Center]========================>>>>

        # region <<<<==============================[Offset Amount]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Offset Amount.                                                      |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : PistPinBoreOffsetAmount                                                      |
        # ===============================================================================================|
        global offset_amount
        engine_worx_database_cursor.execute(
            'SELECT PistPinBoreOffsetAmount FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # =========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                        |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as 'String',|
        #     otherwise, set it <0> (without digits) if it is Zero,                                |
        #     or Round the Numerical Value for 4-Digits.                                           |
        # [#] Expected Data output Type : 'Numeric'                                                |
        # [#] Expected Data output value : Numeric Value > [0.????]                                |
        # =========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                offset_amount = data
            else:
                if (data == 0.0):
                    offset_amount = math.floor(data)
                else:
                    offset_amount = round(float(data), 4)
        print("[#]offset_amount FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", offset_amount)

        # endregion <<<<============================[Offset Amount]============================>>>>

        # region <<<<==============================[Offset Direction]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Offset Direction.                                                   |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : PistPinBoreOffsetDirection                                                   |
        # ===============================================================================================|
        global offset_direction
        engine_worx_database_cursor.execute(
            'SELECT PistPinBoreOffsetDirection FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ===========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                          |
        # [#] Offset Direction Stored in the DataBase as 'Text' in different ways, because of that   |
        #     it needs to check many options to find out the offset direction.                       |
        # [#] Set the value to be "OFFSET To0" if indicate direction <To 0>,                         |
        # [#] Set the value to be "OFFSET To180" if indicate direction <To 180>,                     |
        # [#] Set the value to be "OFFSET EACH WAY" if indicate direction <1/2 EACH WAY>,            |
        # [#] Expected Data output Type : 'String'.                                                  |
        # [#] Expected Data output value : String Value indicate : <To 0>, <To 180> or <1/2 EACH WAY>|
        # ===========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data == "0") or (data == 'TO 0') or (data == 'TOWARD 0') or (data == 'TOWARDS 0') or
                    (data == 'INT') or (data == 'TO INTAKE') or (data == 'AWAY BUMP') or (data == '0 OFFSET') or
                    (data == 'OFFSET TOWARD 0') or (data == 'OFFSET TO INTAKE') or (data == 'TO ZERO') or
                    (data == 'TO INT.') or (data == 'TO 0 SIDE') or (data == '0               ')):
                offset_direction = "OFFSET To0"
            elif((data == '180') or (data == 'TO 180') or (data == 'TOWARD 180') or (data == 'TOWARDS 180') or
                 (data == 'EXT') or (data == 'TO EXHAUST') or (data == 'TOWARDS BUMP') or (data == '180 OFFSET') or
                 (data == 'TO BUMP') or (data == 'TO EXH.') or (data == 'TO 180 SIDE') or
                 (data == '180               ')):
                offset_direction = "OFFSET To180"
            elif ((data == '1/2 EACH WAY') or (data == '1/2 EACH WAY ') or (data == 'OFFSET EACHWAY') or
                  (data == 'OFFSET 1/2 EACH') or (data == '1/2 OFFSET EACHWAY') or (data == 'HALF EACH WAY') or
                  (data == 'EACH WAY') or (data == 'HALF WAY') or (data == '0 AND 180') or (data == 'OFFSET EACH WAY')
                  or (data == '1/2 EACH DIRECTION') or (data == 'OFFSET HALF WAY') or (data == 'OFFSET 1/2 WAY') or
                  (data == 'OFFSET 1/2 EACH WA') or (data == 'OFFSET 1/2 EACH WAY') or (data == 'OFFSET HALF EACH WAY')
                  or (data == 'OFFSET 1/2 EACHWAY') or (data == 'OFFSET 1/2 EACH WAY  ') or (data == 'offset 1/2 way')
                  or (data == 'OFFSET TO 0 FOR L PARTS, TO 180 FOR R PARTS') or
                  (data == 'OFFSET 1/2 EACH WAY RIGHT TOWARDS 180, LEFT TOWARD 0')):
                offset_direction = "OFFSET EACH WAY"

            # ========================================================================================================|
            # [#] if can't detect Offset Direction from the DataBase, check 'PinBoreNotes' if it has anything related |
            #    on offset direction. (Some old jobs missing offset direction and have it as a note in 'PinBoreNotes')|
            # [#] Check PinBoreNotes in te DataBase for Offset Direction.                                             |
            # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:         |
            #     Table Name : SpexPiston_PinBore                                                                     |
            #     Column Name : PistPinBoreNotes                                                                      |
            # ========================================================================================================|
            else:
                engine_worx_database_cursor.execute(
                    'SELECT PistPinBoreNotes FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

                # ===========================================================================================|
                # [#] Use for loop to iterate through the DateBase.                                          |
                # [#] Offset Direction Stored in the DataBase as 'Text' in different ways, because of that   |
                #     it needs to check many options to find out the offset direction.                       |
                # [#] Set the value to be "OFFSET To0" if indicate direction <To 0>,                         |
                # [#] Set the value to be "OFFSET To180" if indicate direction <To 180>,                     |
                # [#] Set the value to be "OFFSET EACH WAY" if indicate direction <1/2 EACH WAY>,            |
                # [#] Expected Data output Type : 'String'.                                                  |
                # [#] Expected Data output value : String Value indicate : <To 0>, <To 180> or <1/2 EACH WAY>|
                # ===========================================================================================|
                for data in engine_worx_database_cursor.fetchone():
                    if ((data == "0") or (data == 'TO 0') or (data == 'TOWARD 0') or (data == 'TOWARDS 0') or
                        (data == 'TOWARD INT') or (data == 'TOWARDS INT') or (data == 'INT') or
                        (data == 'TO INTAKE') or (data == 'AWAY BUMP') or (data == '0 OFFSET') or
                        (data == 'OFFSET TOWARD 0') or (data == 'OFFSET TO INTAKE') or (data == 'TO ZERO') or
                            (data == 'TO INT.') or (data == 'TO 0 SIDE')):
                        offset_direction = "OFFSET To0"
                    elif ((data == '180') or (data == 'TO 180') or (data == 'TOWARD 180') or (data == 'TOWARDS 180') or
                          (data == 'TOWARD EXH') or (data == 'TOWARDS EXH') or (data == 'EXT') or (data == 'TO EXHAUST')
                          or (data == 'TOWARDS BUMP') or (data == '180 OFFSET') or (data == 'TO BUMP') or
                          (data == 'TO EXH.') or (data == 'TO 180 SIDE')):
                        offset_direction = "OFFSET To180"
                    elif ((data == '1/2 EACH WAY') or (data == '1/2 EACH WAY ') or (data == 'OFFSET EACHWAY') or
                          (data == 'OFFSET 1/2 EACH') or (data == '1/2 OFFSET EACHWAY') or (data == 'HALF EACH WAY') or
                          (data == 'EACH WAY') or (data == 'HALF WAY') or (data == '0 AND 180') or
                          (data == 'OFFSET EACH WAY') or (data == '1/2 EACH DIRECTION') or (data == 'OFFSET HALF WAY')
                          or (data == 'OFFSET 1/2 WAY') or (data == 'OFFSET 1/2 EACH WA') or
                          (data == 'OFFSET 1/2 EACH WAY') or (data == 'OFFSET HALF EACH WAY') or
                          (data == 'OFFSET 1/2 EACH WAY  ') or (data == 'offset 1/2 way') or
                          (data == 'OFFSET 1/2 EACHWAY') or (data == 'OFFSET TO 0 FOR L PARTS, TO 180 FOR R PARTS') or
                          (data == 'OFFSET 1/2 EACH WAY RIGHT TOWARDS 180, LEFT TOWARD 0')):
                        offset_direction = "OFFSET EACH WAY"
                    else:
                        offset_direction = ""
        print("[#]offset_direction FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", offset_direction)

        # endregion <<<<============================[Offset Direction]============================>>>>

        # region <<<<==============================[Rough Bore Speed]==============================>>>>

        # ===================================================|
        # [#] Define Variable to set the Speed of Rough Bore.|
        # [#] The Value comes from horizontal template.      |
        # ===================================================|
        global rough_bore_speed
        rough_bore_speed = 8000

        # endregion <<<<===========================[Rough Bore Speed]==============================>>>>

        # region <<<<==============================[Rough Bore Feed]==============================>>>>

        # ==================================================|
        # [#] Define Variable to set the Feed of Rough Bore.|
        # [#] The Value comes from horizontal template.     |
        # ==================================================|
        global rough_bore_feed
        rough_bore_feed = 100

        # endregion <<<<===========================[Rough Bore Feed]=============================>>>>

        # region <<<<====================[Value Used In Z_Value Finish Bore Bottom]====================>>>>

        # ================================================================|
        # [#] Define Variable to set the Z_Value of Bottom of Finish Bore.|
        # [#] The Value comes from horizontal template.                   |
        # ================================================================|
        global value_used_in_Z_value_finish_bore_bottom
        value_used_in_Z_value_finish_bore_bottom = 1.156

        # endregion <<<<=================[Value Used In Z_Value Finish Bore Bottom]====================>>>>

        # region <<<<========================[Ledge Tool Diameter]========================>>>>

        # region <<<<========================[Ledge Cut Availability Status]========================>>>>

        # ==========================================================================================================|
        # [#] ledge_cut_availability_status <Note>.                                                                 |
        # [#] Define Variable to set Ledge Cut Availability Status.                                                 |
        # [#] Set the Value to be "" as a default, the value will change to <1> or <0>  if detect the status, If not|
        #     and value still equal "", it will ask user to choose the Ledge cut Status (Needs or Doesn't Need).    |
        # ==========================================================================================================|
        global ledge_cut_availability_status
        ledge_cut_availability_status = ""

        # endregion <<<<=====================[Ledge Cut Availability Status]========================>>>>

        # =====================================================================|
        # [#] Define Variable to set the Diameter of Ledge Tool.               |
        # [#] Set Value as <0.625> which is the diameter of Stander Ledge Tool.|
        # =====================================================================|
        global ledge_tool_diameter
        ledge_tool_diameter = 0.625

        # endregion <<<<=====================[Ledge Tool Diameter]========================>>>>

        # region <<<<==============================[Lock Ring Cutter Width]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set LockRing Width.                                                     |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : RetClipGrvWidth                                                              |
        # ===============================================================================================|
        global lock_ring_cutter_width
        engine_worx_database_cursor.execute(
            'SELECT RetClipGrvWidth FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'Numeric'                                               |
        # [#] Expected Data output value : Numeric Value > [0.????]                               |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                lock_ring_cutter_width = data
            else:
                lock_ring_cutter_width = round(float(data), 4)
        print("[#]lock_ring_cutter_width FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", lock_ring_cutter_width)

        # endregion <<<<============================[Lock Ring Cutter Width]============================>>>>

        # region <<<<==============================[Lock Ring ID Spacing]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set LockRing ID Spacing.                                                |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : RetClipGrvInnerDiameterSpace                                                 |
        # ===============================================================================================|
        global lock_ring_ID_spacing
        engine_worx_database_cursor.execute(
            'SELECT RetClipGrvInnerDiameterSpace FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'Numeric'                                               |
        # [#] Expected Data output value : Numeric Value > [0.????]                               |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                lock_ring_ID_spacing = data
            else:
                lock_ring_ID_spacing = round(float(data), 4)
        print("[#]lock_ring_ID_spacing FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", lock_ring_ID_spacing)

        # endregion <<<<============================[Lock Ring ID Spacing]============================>>>>

        # region <<<<==============================[Lock Ring Diameter]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set LockRing Diameter.                                                  |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : RetClipGrvDiameter                                                           |
        # ===============================================================================================|
        global lock_ring_diameter
        engine_worx_database_cursor.execute(
            'SELECT RetClipGrvDiameter FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'Numeric'                                               |
        # [#] Expected Data output value : Numeric Value > [0.????]                               |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                lock_ring_diameter = data
            else:
                lock_ring_diameter = round(float(data), 4)
        print("[#]lock_ring_diameter FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", lock_ring_diameter)

        # endregion <<<<============================[Lock Ring Diameter]============================>>>>

        # region <<<<==============================[Lock Ring Tool Diameter]==============================>>>>

        # =========================================================================|
        # [#] Define Variable to set the Diameter of Tool used to cut the LockRing.|
        # [#] The Value comes from (HORIZONTAL_SHEETS_FOR_AUTOMATION) Excel File.  |
        # =========================================================================|
        global lock_ring_tool_diameter
        lock_ring_tool_diameter = 0

        # endregion <<<<===========================[Lock Ring Tool Diameter]============================>>>>

        # region <<<<==============================[C/Fren ID Spacing]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set C/Fren ID Spacing.                                                  |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : CFrenGrvInnerDiameterSpace                                                   |
        # ===============================================================================================|
        global cfren_ID_spacing
        engine_worx_database_cursor.execute(
            'SELECT CFrenGrvInnerDiameterSpace FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'Numeric'                                               |
        # [#] Expected Data output value : Numeric Value > [0.????]                               |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                cfren_ID_spacing = data
            else:
                cfren_ID_spacing = round(float(data), 4)
        print("[#]cfren_ID_spacing FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", cfren_ID_spacing)

        # region <<<<==============================[C/Fren Cutter Width]==============================>>>>

        # ===============================================================================================|
        # [#] C/Fren Cutter Width has no Certain Data Field location, and sometimes it's be as a note    |
        #     in 'CFrenNotes', because of that needs to check CFrenNotes in the DataBase.                |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : CFrenNotes                                                                   |
        # ===============================================================================================|
        global cfren_cutter_width
        engine_worx_database_cursor.execute('SELECT CFrenNotes FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ============================================================================================================|
        # [#] Use for loop to iterate through the DateBase and use if_Statement to find out C/Fren Width.             |
        # [#] Two Ways to find out C/Fren Width:                                                                      |
        #    -If CFrenNotes has nothing, and <cfren_ID_spacing> has a value, and Piston has a LockRing as well then:  |
        #        Set <cfren_cutter_width> same as lock_ring_cutter_width                                              |
        #    -If CFrenNotes has a 'text', needs to check the note if it has one of width tool it used for C/Fren then:|
        #        Set <cfren_cutter_width> Equal to tool width                                                         |
        # [#] Use (find() built-in function) to find certain "text" inside the Note, it will return location of       |
        #    the certain "text" when found it, otherwise it will return <-1> to indicate that doesn't found the 'text'|
        # ============================================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if (((data is None) and ((cfren_ID_spacing is not None) and (cfren_ID_spacing != 0)))
                    and ((lock_ring_cutter_width is not None) and (lock_ring_cutter_width != 0))):
                cfren_cutter_width = lock_ring_cutter_width
            elif ((data is not None) and ((cfren_ID_spacing is not None) and (cfren_ID_spacing != 0))):
                if ((data.find('0.039') != -1) or (data.find('.039') != -1)):
                    cfren_cutter_width = 0.039
                elif ((data.find('0.042') != -1) or (data.find('.042') != -1)):
                    cfren_cutter_width = 0.042
                elif ((data.find('0.044') != -1) or (data.find('.044') != -1)):
                    cfren_cutter_width = 0.044
                elif ((data.find('0.047') != -1) or (data.find('.047') != -1)):
                    cfren_cutter_width = 0.047
                elif ((data.find('0.048') != -1) or (data.find('.048') != -1)):
                    cfren_cutter_width = 0.048
                elif ((data.find('0.053') != -1) or (data.find('.053') != -1)):
                    cfren_cutter_width = 0.053
                elif ((data.find('0.059') != -1) or (data.find('.059') != -1)):
                    cfren_cutter_width = 0.059
                elif ((data.find('0.063') != -1) or (data.find('.063') != -1)):
                    cfren_cutter_width = 0.063
                elif ((data.find('0.065') != -1) or (data.find('.065') != -1)):
                    cfren_cutter_width = 0.065
                elif ((data.find('0.067') != -1) or (data.find('.067') != -1)):
                    cfren_cutter_width = 0.067
                elif ((data.find('0.076') != -1) or (data.find('.076') != -1)):
                    cfren_cutter_width = 0.076
                elif ((data.find('0.077') != -1) or (data.find('.077') != -1)):
                    cfren_cutter_width = 0.077
                elif ((data.find('0.088') != -1) or (data.find('.088') != -1)):
                    cfren_cutter_width = 0.088
                elif ((data.find('0.11') != -1) or (data.find('.11') != -1)):
                    cfren_cutter_width = 0.11
                elif ((lock_ring_cutter_width is not None) and (lock_ring_cutter_width != 0)):
                    cfren_cutter_width = lock_ring_cutter_width
                else:
                    cfren_cutter_width = 0
            else:
                cfren_cutter_width = 0
        print("[#]cfren_cutter_width FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", cfren_cutter_width)

        # endregion <<<<============================[C/Fren Cutter Width]============================>>>>

        # endregion <<<<==============================[C/Fren ID Spacing]============================>>>>

        # region <<<<==============================[C/Fren Diameter]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set C/Fren Diameter.                                                    |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : CFrenGrvDiameter                                                             |
        # ===============================================================================================|
        global cfren_diameter
        engine_worx_database_cursor.execute(
            'SELECT CFrenGrvDiameter FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'Numeric'                                               |
        # [#] Expected Data output value : Numeric Value > [0.????]                               |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            # print(i)
            # NEEDS TO MAKE IT FLOAT
            if ((data is None) or (type(data) == str)):
                cfren_diameter = data
            else:
                cfren_diameter = round(float(data), 4)
        print("[#]cfren_diameter FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", cfren_diameter)

        # endregion <<<<============================[C/Fren Diameter]============================>>>>

        # region <<<<==============================[C/Fren Tool Diameter]==============================>>>>

        # =======================================================================|
        # [#] Define Variable to set the Diameter of Tool used to cut the C/Fren.|
        # [#] The Value comes from (HORIZONTAL_SHEETS_FOR_AUTOMATION) Excel File.|
        # =======================================================================|
        global cfren_tool_diameter
        cfren_tool_diameter = 0  # JUST DEFINE THAT TO USE IT ON MATH LATER ,THIS NUMBER COMES FROM TOOLLISTSHEET

        # endregion <<<<===========================[C/Fren Tool Diameter]============================>>>>

        # region <<<<==============================[Semi C/Fren ID Spacing]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Semi_C/Fren ID Spacing.                                             |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : SemiCFrenGrvInnerDiameterSpace                                               |
        # ===============================================================================================|
        global semi_cfren_ID_spacing
        engine_worx_database_cursor.execute(
            'SELECT SemiCFrenGrvInnerDiameterSpace FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'Numeric'                                               |
        # [#] Expected Data output value : Numeric Value > [0.????]                               |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                semi_cfren_ID_spacing = data
            else:
                semi_cfren_ID_spacing = round(float(data), 4)
        print("[#]semi_cfren_ID_spacing FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", semi_cfren_ID_spacing)

        # endregion <<<<============================[Semi C/Fren ID Spacing]============================>>>>

        # region <<<<==============================[Semi C/Fren Width]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Semi_C/Fren Width.                                                  |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : SemiCFrenGrvWidth                                                            |
        # ===============================================================================================|
        global semi_cfren_width
        engine_worx_database_cursor.execute(
            'SELECT SemiCFrenGrvWidth FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'Numeric'                                               |
        # [#] Expected Data output value : Numeric Value > [0.????]                               |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                semi_cfren_width = data
            else:
                semi_cfren_width = round(float(data), 4)
        print("[#]semi_cfren_width FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", semi_cfren_width)

        # endregion <<<<============================[Semi C/Fren Width]============================>>>>

        # region <<<<==============================[Semi C/Fren Depth]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Semi_C/Fren Depth.                                                  |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : SemiCFrenGrvDepth                                                            |
        # ===============================================================================================|
        global semi_cfren_depth
        engine_worx_database_cursor.execute(
            'SELECT SemiCFrenGrvDepth FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'Numeric'                                               |
        # [#] Expected Data output value : Numeric Value > [0.????]                               |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                semi_cfren_depth = data
            else:
                semi_cfren_depth = round(float(data), 4)
        print("[#]semi_cfren_depth FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", semi_cfren_depth)

        # endregion <<<<============================[Semi C/Fren Depth]============================>>>>

        # region <<<<=========================[Ret Clip Notch Angle <First Location>]============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set First Location of Clip Notch.                                       |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : RetClipNotchLocAngle01                                                       |
        # ===============================================================================================|
        global notch_angle_first_location
        engine_worx_database_cursor.execute(
            'SELECT RetClipNotchLocAngle01 FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'Numeric'                                               |
        # [#] Expected Data output value : Numeric Value > [0] or Numeric Value < [0]             |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                notch_angle_first_location = data
            else:
                notch_angle_first_location = round(float(data), 4)
        print("[#]notch_angle_first_location FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", notch_angle_first_location)

        # endregion <<<<=========================[Ret Clip Notch Angle <First Location>]========================>>>>

        # region <<<<=========================[Ret Clip Notch Angle <Second Location>]============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Second Location of Clip Notch.                                      |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : RetClipNotchLocAngle02                                                       |
        # ===============================================================================================|
        global notch_angle_second_location
        engine_worx_database_cursor.execute(
            'SELECT RetClipNotchLocAngle02 FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'Numeric'                                               |
        # [#] Expected Data output value : Numeric Value > [0] or Numeric Value < [0]             |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                notch_angle_second_location = data
            else:
                notch_angle_second_location = round(float(data), 4)
        print("[#]notch_angle_second_location FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", notch_angle_second_location)

        # endregion <<<<=========================[Ret Clip Notch Angle <Second Location>]========================>>>>

        # region <<<<==========================[X_Distance From Origin To Circlip Notch]==========================>>>>

        # ===========================================================================|
        # [#] Define Variable to set the X_distance from Piston Origin to Clip Notch.|
        # [#] The Value will be Calculated by Math Later.                            |
        # ===========================================================================|
        global X_distance_from_origin_to_circlip_notch
        X_distance_from_origin_to_circlip_notch = 0

        # endregion <<<<=======================[X_Distance From Origin To Circlip Notch]=========================>>>>

        # region <<<<==========================[Y_Distance From Origin To Circlip Notch]==========================>>>>

        # ===========================================================================|
        # [#] Define Variable to set the Y_distance from Piston Origin to Clip Notch.|
        # [#] The Value will be Calculated by Math Later.                            |
        # ===========================================================================|
        global Y_distance_from_origin_to_circlip_notch
        Y_distance_from_origin_to_circlip_notch = 0

        # endregion <<<<=======================[Y_Distance From Origin To Circlip Notch]=========================>>>>

        # region <<<<==============================[Double Oil Holes Slots ID Spacing]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set ID_Spacing of Double Oil Hole Slot.                                 |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_GasPortsPistonPinOiling                                            |
        #     Column Name : PressureFedOilHoleInnerDiameterSpace                                         |
        # ===============================================================================================|
        global double_oil_hole_slot_ID_spacing
        engine_worx_database_cursor.execute(
            'SELECT PressureFedOilHoleInnerDiameterSpace FROM SpexPiston_GasPortsPistonPinOiling WHERE PistonID = ?',
            piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'Numeric'                                               |
        # [#] Expected Data output value : Numeric Value > [0.????]                               |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                double_oil_hole_slot_ID_spacing = data
            else:
                double_oil_hole_slot_ID_spacing = round(float(data), 4)
        print("[#]double_oil_hole_slot_ID_spacing FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", double_oil_hole_slot_ID_spacing)

        # endregion <<<<============================[Double Oil Holes Slots ID Spacing]============================>>>>

        # region <<<<==============================[Horizontal Slots OD Spacing]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set OD_Spacing of Horizontal Pin Slot.                                  |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : HorizPistonPinSlotsOuterDiameterSpacing                                      |
        # ===============================================================================================|
        global horizontal_slots_OD_spacing
        engine_worx_database_cursor.execute(
            'SELECT HorizPistonPinSlotsOuterDiameterSpacing FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'Numeric'                                               |
        # [#] Expected Data output value : Numeric Value > [0.????]                               |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                horizontal_slots_OD_spacing = data
            else:
                horizontal_slots_OD_spacing = round(float(data), 4)
        print("[#]horizontal_slots_OD_spacing from DataBase FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", horizontal_slots_OD_spacing)

        # region <<<<==============================[Horizontal Slots Notes]==============================>>>>

        # ===============================================================================================|
        # [#] Some Jobs has the H-SlotsOD_Spacing value in HorizSlotsNotes, or the Notes has Something   |
        #     related on OD_Spacing, because of that needs to Define Variable to set HorizSlotsNotes.    |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : HorizPistonPinSlotsNotes                                                     |
        # ===============================================================================================|
        global horiz_piston_pin_slots_notes
        engine_worx_database_cursor.execute(
            'SELECT HorizPistonPinSlotsNotes FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # =================================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                                |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',         |
        #     otherwise, Round the Numerical Value for 4-Digits.                                           |
        # [#] Expected Data output Type : 'String'                                                         |
        # [#] Expected Data output value : String Value has H-Slots OD_Spacing or Text to decide the Value.|
        # =================================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                horiz_piston_pin_slots_notes = data
            else:
                horiz_piston_pin_slots_notes = round(float(data), 4)
        print("[#]horiz_piston_pin_slots_notes FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", horiz_piston_pin_slots_notes)

        # endregion <<<<============================[Horizontal Slots Notes]============================>>>>

        # ========================================================================================================|
        # [#] Horizontal Slots Design is Changed from old to new jobs and it's different from part to part,       |
        #     because of that,it needs to check some options to decide OD_Spacing value:                          |
        #     [#] If HorizSlotsNotes is not Empty and has a 'Text',                                               |
        #        -Set the OD_Spacing to be equal "", to let user enter the value later.                           |
        #     [#] If H-slots OD_spacing is not Exist or equal <0>,                                                |
        #        -Set the OD_Spacing to be equal "", to keep a chance to the user to enter the value if necessary.|
        #     [#] If H-slots OD_spacing has a value and job has LockRing as well and H-slots OD_spacing is Bigger |
        #         than or equal LockRing_ID_Spacing (that's indicate the new H-Slot Design).                      |
        #        -Set the OD_Spacing to be equal: LockRing_ID_Spacing + LockRing_Cutter_Width.                    |
        #     [#] If H-slots OD_spacing has a value and job has LockRing as well and H-slots OD_spacing is Smaller|
        #         than LockRing_ID_Spacing (that's indicate the old H-Slot Design).                               |
        #        -Set the OD_Spacing to be equal as it's stored in the DataBase without any change by use "pass". |
        #     [#] If H-slots OD_spacing has a value and job has no LockRing:                                      |
        #        -Set the OD_Spacing to be equal as it's stored in the DataBase without any change by use "pass". |
        # ========================================================================================================|
        if (horiz_piston_pin_slots_notes is not None and horiz_piston_pin_slots_notes != ""):
            horizontal_slots_OD_spacing = ""
        elif (horizontal_slots_OD_spacing is None or horizontal_slots_OD_spacing == "" or
              horizontal_slots_OD_spacing == 0):
            horizontal_slots_OD_spacing = ""
        elif ((horizontal_slots_OD_spacing is not None and
               horizontal_slots_OD_spacing != "" and horizontal_slots_OD_spacing != 0) and
              (lock_ring_ID_spacing is not None and lock_ring_ID_spacing != "" and lock_ring_ID_spacing != 0) and
              (lock_ring_cutter_width is not None and lock_ring_cutter_width != "" and lock_ring_cutter_width != 0) and
              (horizontal_slots_OD_spacing >= lock_ring_ID_spacing)):
            horizontal_slots_OD_spacing = round(lock_ring_ID_spacing + lock_ring_cutter_width, 4)
        elif ((horizontal_slots_OD_spacing is not None and horizontal_slots_OD_spacing != "" and
               horizontal_slots_OD_spacing != 0) and
              (lock_ring_ID_spacing is not None and lock_ring_ID_spacing != "" and lock_ring_ID_spacing != 0) and
              (lock_ring_cutter_width is not None and lock_ring_cutter_width != ""
                and lock_ring_cutter_width != 0) and (horizontal_slots_OD_spacing < lock_ring_ID_spacing)):
            pass
        elif ((horizontal_slots_OD_spacing is not None and horizontal_slots_OD_spacing != "" and
               horizontal_slots_OD_spacing != 0)):
            pass
        else:
            horizontal_slots_OD_spacing = ""
        print("[#]horizontal_slots_OD_spacing AFTER logic FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", horizontal_slots_OD_spacing)

        # endregion <<<<============================[Horizontal Slots OD Spacing]============================>>>>

        # region <<<<==============================[Horizontal Slots Arc Diameter]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Arc Diameter of Horizontal Slots.                                   |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : HorizPistonPinSlotsArcDiameter                                               |
        # ===============================================================================================|
        global horizontal_slots_arc_diameter
        engine_worx_database_cursor.execute(
            'SELECT HorizPistonPinSlotsArcDiameter FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'Numeric'                                               |
        # [#] Expected Data output value : Numeric Value > [0.????]                               |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                horizontal_slots_arc_diameter = data
            else:
                horizontal_slots_arc_diameter = round(float(data), 4)
        print("[#]horizontal_slots_arc_diameter FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", horizontal_slots_arc_diameter)

        # endregion <<<<============================[Horizontal Slots Arc Diameter]============================>>>>

        # region <<<<==============================[Horizontal Slots Diameter Depth]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Diameter Depth of Horizontal Slots.                                 |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : HorizPistonPinSlotsDiameterDepth                                             |
        # ===============================================================================================|
        global horizontal_slots_diameter_depth
        engine_worx_database_cursor.execute(
            'SELECT HorizPistonPinSlotsDiameterDepth FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'Numeric'                                               |
        # [#] Expected Data output value : Numeric Value > [0.????]                               |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                horizontal_slots_diameter_depth = data
            else:
                horizontal_slots_diameter_depth = round(float(data), 4)
        print("[#]horizontal_slots_diameter_depth FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", horizontal_slots_diameter_depth)

        # endregion <<<<============================[Horizontal Slots Diameter Depth]============================>>>>

        # region <<<<==========================[Horizontal Slots Through Boss Status]=============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Status of Horizontal Slots Through Boss.                            |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : HorizPistonPinSlotsThruBoss                                                  |
        # ===============================================================================================|
        global horizontal_slots_through_Boss_status
        engine_worx_database_cursor.execute(
            'SELECT HorizPistonPinSlotsThruBoss FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'String'                                                |
        # [#] Expected Data output value : String Value: 'Y' or 'N'                               |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                horizontal_slots_through_Boss_status = data
            else:
                horizontal_slots_through_Boss_status = round(float(data), 4)
        print("[#]horizontal_slots_through_Boss_status FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", horizontal_slots_through_Boss_status)

        # endregion <<<<==========================[Horizontal Slots Through Boss Status]===========================>>>>

        # region <<<<==========================[i_Start Horizontal Slot Number]==========================>>>>

        # =======================================================================|
        # [#] Define Variable to set the i_Start Horizontal Slot Number.         |
        # [#] The Value comes from (HORIZONTAL_SHEETS_FOR_AUTOMATION) Excel File.|
        # =======================================================================|
        global i_start_horizontal_slot
        i_start_horizontal_slot = 0

        # endregion <<<<=======================[i_Start Horizontal Slot Number]=========================>>>>

        # region <<<<==========================[j_Start Horizontal Slot Number]==========================>>>>

        # =======================================================================|
        # [#] Define Variable to set the j_Start Horizontal Slot Number.         |
        # [#] The Value comes from (HORIZONTAL_SHEETS_FOR_AUTOMATION) Excel File.|
        # =======================================================================|
        global j_start_horizontal_slot
        j_start_horizontal_slot = 0

        # endregion <<<<=======================[j_Start Horizontal Slot Number]=========================>>>>

        # region <<<<==========================[Horizontal Slot Radius Number]==========================>>>>

        # =======================================================================|
        # [#] Define Variable to set the horizontal_slot_radius Number.          |
        # [#] The Value comes from (HORIZONTAL_SHEETS_FOR_AUTOMATION) Excel File.|
        # =======================================================================|
        global horizontal_slot_radius
        horizontal_slot_radius = 0

        # endregion <<<<=======================[Horizontal Slot Radius Number]=========================>>>>

        # region <<<<==============================[Ledge Counterbore Diameter]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Ledge Counterbore Diameter.                                         |
        # [#] Ledge Counterbore Diameter has no Certain Data Field location, and sometimes it's be       |
        #     as a note in "PistPinBoreNotes", "RetClipNotes" or in one of "MillingProgramTypes",        |
        #     because of that needs to go through of them and check to find out the Counterbore Diameter.|
        # [#] Reference Jobs Ex: WD-15479, WD-14141, WD-11210, WD-10911, and WD-08723.                   |
        # ===============================================================================================|
        global ledge_counterbore_diameter

        # =============================================================================================|
        # [#] Get the Value of "PistPinBoreNotes" from the DataBase by using 'piston_id' and Access the|
        #     Data location by using:                                                                  |
        #     Table Name : SpexPiston_PinBore                                                          |
        #     Column Name : PistPinBoreNotes                                                           |
        # =============================================================================================|
        engine_worx_database_cursor.execute(
            'SELECT PistPinBoreNotes FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ==================================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                                 |
        # [#] Counterbore Diameter Stored in the "PistPinBoreNotes" in different ways, because of that      |
        #     it needs to check many options to find out the Counterbore Diameter.                          |
        # [#] Set the value to be <0> if "PistPinBoreNotes" is Empty(None).                                 |
        # [#] Set the value to be "" if found some "text" indicated of Counterbore Diameter, to let user    |
        #     enter the value later.                                                                        |
        # [#] Set the value to be <0> if nothing above apply, to keep searching in the next                 |
        #     Notes ("RetClipNotes", "MillingProgramType01"...etc).                                         |
        # [#] Use (find() built-in function) to find certain "text" inside the Note.                        |
        # [#] (Ledge Counterbore Diameter = 0) means nothing found in "PistPinBoreNotes" and keep searching.|
        # [#] (Ledge Counterbore Diameter = "") means Counterbore indicated in "PistPinBoreNotes" and       |
        #     needs user to enter the value later.                                                          |
        # [#] Expected Data output Type : 'String'.                                                         |
        # [#] Expected Data output value : String Value indicate 'text' of Counterbore.                     |
        # ==================================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if (data is None):
                ledge_counterbore_diameter = 0
                # print("ledge_counterbore_diameter >> (if statment for PistPinBoreNotes) ", ledge_counterbore_diameter)
            elif (data is not None and ((data.find('COUNTERBORE') != -1) or (data.find(
                    'LARGER AT LEDGE, CUT AT HORIZONTAL') != -1) or (data.find("C'BORE") != -1) or (data.find(
                    'COUNTER BORE') != -1) or (data.find('COUNTERBORES') != -1) or (data.find('CBORE') != -1) or
                    (data.find('CNTBOR') != -1))):
                ledge_counterbore_diameter = ""
                # print("ledge_counterbore_diameter >> (elif statment for PistPinBoreNotes): ",
                # ledge_counterbore_diameter)
            else:
                ledge_counterbore_diameter = 0
                # print("ledge_counterbore_diameter >> (else statment for PistPinBoreNotes): ",
                #       ledge_counterbore_diameter)

        # ========================================================================================================|
        # [#] If COUNTERBORE didn't detect in the "PistPinBoreNotes" (ledge_counterbore_diameter still equal <0>),|
        #     check "RetClipNotes" Note                                                                           |
        # [#] Get the Value of "RetClipNotes" from the DataBase by using 'piston_id' and Access the               |
        #     Data location by using:                                                                             |
        #     Table Name : SpexPiston_PinBore                                                                     |
        #     Column Name : RetClipNotes                                                                          |
        # ========================================================================================================|
        if (ledge_counterbore_diameter == 0):
            engine_worx_database_cursor.execute(
                'SELECT RetClipNotes FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

            # =================================================================================================|
            # [#] Use for loop to iterate through the DateBase.                                                |
            # [#] Counterbore Diameter Stored in the "RetClipNotes" in different ways, because of that it needs|
            #     to check many options to find out the Counterbore Diameter.                                  |
            # [#] Set the value to be <0> if "RetClipNotes" is Empty(None).                                    |
            # [#] Set the value to be "" if found some "text" indicated of Counterbore Diameter, to let user   |
            #     enter the value later.                                                                       |
            # [#] Set the value to be <0> if nothing above apply, to keep searching in the next                |
            #     Notes ("MillingProgramType01", "MillingProgramType02"...etc).                                |
            # [#] Use (find() built-in function) to find certain "text" inside the Note.                       |
            # [#] (Ledge Counterbore Diameter = 0) means nothing found in "RetClipNotes" and keep searching.   |
            # [#] (Ledge Counterbore Diameter = "") means Counterbore indicated in "RetClipNotes" and          |
            #     needs user to enter the value later.                                                         |
            # [#] Expected Data output Type : 'String'.                                                        |
            # [#] Expected Data output value : String Value indicate 'text' of Counterbore.                    |
            # =================================================================================================|
            for data in engine_worx_database_cursor.fetchone():
                if (data is None):
                    ledge_counterbore_diameter = 0
                    # print("ledge_counterbore_diameter >> (if statment for RetClipNotes): ",
                    # ledge_counterbore_diameter)
                elif (data is not None and (
                        (data.find('COUNTERBORE') != -1) or (data.find('LARGER AT LEDGE, CUT AT HORIZONTAL') != -1) or
                        (data.find("C'BORE") != -1) or (data.find('COUNTER BORE') != -1) or
                        (data.find('COUNTERBORES') != -1) or (data.find('CBORE') != -1) or
                        (data.find('CNTBOR') != -1))):
                    ledge_counterbore_diameter = ""
                    # print("ledge_counterbore_diameter >> (elif statment for RetClipNotes): ",
                    #       ledge_counterbore_diameter)
                else:
                    ledge_counterbore_diameter = 0
                    # print("ledge_counterbore_diameter >> (else statment for RetClipNotes): ",
                    #       ledge_counterbore_diameter)

        # ====================================================================================================|
        # [#] If COUNTERBORE didn't detect in the "RetClipNotes" (ledge_counterbore_diameter still equal <0>),|
        #     check "MillingProgramType01" Note.                                                              |
        # [#] Get the Value of "MillingProgramType01" from the DataBase by using 'piston_id' and Access the   |
        #     Data location by using:                                                                         |
        #     Table Name : SpexPiston_Milling                                                                 |
        #     Column Name : MillingProgramType01                                                              |
        # ====================================================================================================|
        if (ledge_counterbore_diameter == 0):
            engine_worx_database_cursor.execute(
                'SELECT MillingProgramType01 FROM SpexPiston_Milling WHERE PistonID = ?', piston_id)

            # ======================================================================================================|
            # [#] Use for loop to iterate through the DateBase.                                                     |
            # [#] Counterbore Diameter Stored in the "MillingProgramType01" in different ways, because of that      |
            #     it needs to check many options to find out the Counterbore Diameter.                              |
            # [#] Set the value to be <0> if "MillingProgramType01" is Empty(None).                                 |
            # [#] Set the value to be "" if found some "text" indicated of Counterbore Diameter, to let user        |
            #     enter the value later.                                                                            |
            # [#] Set the value to be <0> if nothing above apply, to keep searching in the next                     |
            #     Notes ("MillingProgramType02", "MillingProgramType03"...etc).                                     |
            # [#] Use (find() built-in function) to find certain "text" inside the Note.                            |
            # [#] (Ledge Counterbore Diameter = 0) means nothing found in "MillingProgramType01" and keep searching.|
            # [#] (Ledge Counterbore Diameter = "") means Counterbore indicated in "MillingProgramType01" and       |
            #     needs user to enter the value later.                                                              |
            # [#] Expected Data output Type : 'String'.                                                             |
            # [#] Expected Data output value : String Value indicate 'text' of Counterbore.                         |
            # ======================================================================================================|
            for data in engine_worx_database_cursor.fetchone():
                # print(i)
                if (data is None):
                    ledge_counterbore_diameter = 0
                    # print("ledge_counterbore_diameter >> (if statment for MillingProgramType01): ",
                    #       ledge_counterbore_diameter)
                elif (data is not None and (
                        (data.find('COUNTERBORE') != -1) or (
                        data.find('LARGER AT LEDGE, CUT AT HORIZONTAL') != -1) or
                        (data.find("C'BORE") != -1) or (data.find('COUNTER BORE') != -1) or
                        (data.find('COUNTERBORES') != -1) or (data.find('CBORE') != -1) or (
                                data.find('CNTBOR') != -1))):
                    ledge_counterbore_diameter = ""
                    # print("ledge_counterbore_diameter >> (elif statment for MillingProgramType01): ",
                    #       ledge_counterbore_diameter)
                else:
                    ledge_counterbore_diameter = 0
                    # print("ledge_counterbore_diameter >> (else statment for MillingProgramType01): ",
                    #       ledge_counterbore_diameter)

        # ============================================================================================================|
        # [#] If COUNTERBORE didn't detect in the "MillingProgramType01" (ledge_counterbore_diameter still equal <0>),|
        #     check "MillingProgramType02" Note.                                                                      |
        # [#] Same Description of "MillingProgramType01" with changing the 'Type Number'.                             |
        # ============================================================================================================|
        if (ledge_counterbore_diameter == 0):
            engine_worx_database_cursor.execute(
                'SELECT MillingProgramType02 FROM SpexPiston_Milling WHERE PistonID = ?', piston_id)

            for data in engine_worx_database_cursor.fetchone():
                if (data is None):
                    ledge_counterbore_diameter = 0
                elif (data is not None and (
                        (data.find('COUNTERBORE') != -1) or (
                        data.find('LARGER AT LEDGE, CUT AT HORIZONTAL') != -1) or
                        (data.find("C'BORE") != -1) or (data.find('COUNTER BORE') != -1) or
                        (data.find('COUNTERBORES') != -1) or (data.find('CBORE') != -1) or (
                                data.find('CNTBOR') != -1))):
                    ledge_counterbore_diameter = ""
                else:
                    ledge_counterbore_diameter = 0

        # ============================================================================================================|
        # [#] If COUNTERBORE didn't detect in the "MillingProgramType02" (ledge_counterbore_diameter still equal <0>),|
        #     check "MillingProgramType03" Note.                                                                      |
        # [#] Same Description of "MillingProgramType01" with changing the 'Type Number'                              |
        # ============================================================================================================|
        if (ledge_counterbore_diameter == 0):
            engine_worx_database_cursor.execute(
                'SELECT MillingProgramType03 FROM SpexPiston_Milling WHERE PistonID = ?', piston_id)

            for data in engine_worx_database_cursor.fetchone():
                if (data is None):
                    ledge_counterbore_diameter = 0
                elif (data is not None and (
                        (data.find('COUNTERBORE') != -1) or (
                        data.find('LARGER AT LEDGE, CUT AT HORIZONTAL') != -1) or
                        (data.find("C'BORE") != -1) or (data.find('COUNTER BORE') != -1) or
                        (data.find('COUNTERBORES') != -1) or (data.find('CBORE') != -1) or (
                                data.find('CNTBOR') != -1))):
                    ledge_counterbore_diameter = ""
                else:
                    ledge_counterbore_diameter = 0

        # ============================================================================================================|
        # [#] If COUNTERBORE didn't detect in the "MillingProgramType03" (ledge_counterbore_diameter still equal <0>),|
        #     check "MillingProgramType04" Note.                                                                      |
        # [#] Same Description of "MillingProgramType01" with changing the 'Type Number'                              |
        # ============================================================================================================|
        if (ledge_counterbore_diameter == 0):
            engine_worx_database_cursor.execute(
                'SELECT MillingProgramType04 FROM SpexPiston_Milling WHERE PistonID = ?', piston_id)

            for data in engine_worx_database_cursor.fetchone():
                if (data is None):
                    ledge_counterbore_diameter = 0
                elif (data is not None and (
                        (data.find('COUNTERBORE') != -1) or (
                        data.find('LARGER AT LEDGE, CUT AT HORIZONTAL') != -1) or
                        (data.find("C'BORE") != -1) or (data.find('COUNTER BORE') != -1) or
                        (data.find('COUNTERBORES') != -1) or (data.find('CBORE') != -1) or (
                                data.find('CNTBOR') != -1))):
                    ledge_counterbore_diameter = ""
                else:
                    ledge_counterbore_diameter = 0

        # ============================================================================================================|
        # [#] If COUNTERBORE didn't detect in the "MillingProgramType04" (ledge_counterbore_diameter still equal <0>),|
        #     check "MillingProgramType05" Note.                                                                      |
        # [#] Same Description of "MillingProgramType01" with changing the 'Type Number'                              |
        # ============================================================================================================|
        if (ledge_counterbore_diameter == 0):
            engine_worx_database_cursor.execute(
                'SELECT MillingProgramType05 FROM SpexPiston_Milling WHERE PistonID = ?', piston_id)

            for data in engine_worx_database_cursor.fetchone():
                if (data is None):
                    ledge_counterbore_diameter = 0
                elif (data is not None and (
                        (data.find('COUNTERBORE') != -1) or (
                        data.find('LARGER AT LEDGE, CUT AT HORIZONTAL') != -1) or
                        (data.find("C'BORE") != -1) or (data.find('COUNTER BORE') != -1) or
                        (data.find('COUNTERBORES') != -1) or (data.find('CBORE') != -1) or (
                                data.find('CNTBOR') != -1))):
                    ledge_counterbore_diameter = ""
                else:
                    ledge_counterbore_diameter = 0

        # ============================================================================================================|
        # [#] If COUNTERBORE didn't detect in the "MillingProgramType05" (ledge_counterbore_diameter still equal <0>),|
        #     check "MillingProgramType06" Note.                                                                      |
        # [#] Same Description of "MillingProgramType01" with changing the 'Type Number'                              |
        # ============================================================================================================|
        if (ledge_counterbore_diameter == 0):
            engine_worx_database_cursor.execute(
                'SELECT MillingProgramType06 FROM SpexPiston_Milling WHERE PistonID = ?', piston_id)
            for data in engine_worx_database_cursor.fetchone():
                if (data is None):
                    ledge_counterbore_diameter = 0
                elif (data is not None and (
                        (data.find('COUNTERBORE') != -1) or (
                        data.find('LARGER AT LEDGE, CUT AT HORIZONTAL') != -1) or
                        (data.find("C'BORE") != -1) or (data.find('COUNTER BORE') != -1) or
                        (data.find('COUNTERBORES') != -1) or (data.find('CBORE') != -1) or (
                                data.find('CNTBOR') != -1))):
                    ledge_counterbore_diameter = ""
                else:
                    ledge_counterbore_diameter = 0

        # ============================================================================================================|
        # [#] If COUNTERBORE didn't detect in the "MillingProgramType06" (ledge_counterbore_diameter still equal <0>),|
        #     check "MillingProgramType07" Note.                                                                      |
        # [#] Same Description of "MillingProgramType01" with changing the 'Type Number'                              |
        # ============================================================================================================|
        if (ledge_counterbore_diameter == 0):
            engine_worx_database_cursor.execute(
                'SELECT MillingProgramType07 FROM SpexPiston_Milling WHERE PistonID = ?', piston_id)
            for data in engine_worx_database_cursor.fetchone():
                if (data is None):
                    ledge_counterbore_diameter = 0
                elif (data is not None and (
                        (data.find('COUNTERBORE') != -1) or (
                        data.find('LARGER AT LEDGE, CUT AT HORIZONTAL') != -1) or
                        (data.find("C'BORE") != -1) or (data.find('COUNTER BORE') != -1) or
                        (data.find('COUNTERBORES') != -1) or (data.find('CBORE') != -1) or (
                                data.find('CNTBOR') != -1))):
                    ledge_counterbore_diameter = ""
                else:
                    ledge_counterbore_diameter = 0

        # ============================================================================================================|
        # [#] If COUNTERBORE didn't detect in the "MillingProgramType07" (ledge_counterbore_diameter still equal <0>),|
        #     check "MillingProgramType08" Note.                                                                      |
        # [#] Same Description of "MillingProgramType01" with changing the 'Type Number'                              |
        # ============================================================================================================|
        if (ledge_counterbore_diameter == 0):
            engine_worx_database_cursor.execute(
                'SELECT MillingProgramType08 FROM SpexPiston_Milling WHERE PistonID = ?', piston_id)
            for data in engine_worx_database_cursor.fetchone():
                if (data is None):
                    ledge_counterbore_diameter = 0
                elif (data is not None and (
                        (data.find('COUNTERBORE') != -1) or (
                        data.find('LARGER AT LEDGE, CUT AT HORIZONTAL') != -1) or
                        (data.find("C'BORE") != -1) or (data.find('COUNTER BORE') != -1) or
                        (data.find('COUNTERBORES') != -1) or (data.find('CBORE') != -1) or (
                                data.find('CNTBOR') != -1))):
                    ledge_counterbore_diameter = ""
                else:
                    ledge_counterbore_diameter = 0

        # ============================================================================================================|
        # [#] If COUNTERBORE didn't detect in the "MillingProgramType08" (ledge_counterbore_diameter still equal <0>),|
        #     check "MillingProgramType09" Note.                                                                      |
        # [#] Same Description of "MillingProgramType01" with changing the 'Type Number'                              |
        # ============================================================================================================|
        if (ledge_counterbore_diameter == 0):
            engine_worx_database_cursor.execute(
                'SELECT MillingProgramType09 FROM SpexPiston_Milling WHERE PistonID = ?', piston_id)
            for data in engine_worx_database_cursor.fetchone():
                if (data is None):
                    ledge_counterbore_diameter = 0
                elif (data is not None and (
                        (data.find('COUNTERBORE') != -1) or (
                        data.find('LARGER AT LEDGE, CUT AT HORIZONTAL') != -1) or
                        (data.find("C'BORE") != -1) or (data.find('COUNTER BORE') != -1) or
                        (data.find('COUNTERBORES') != -1) or (data.find('CBORE') != -1) or (
                                data.find('CNTBOR') != -1))):
                    ledge_counterbore_diameter = ""
                else:
                    ledge_counterbore_diameter = 0

        # ============================================================================================================|
        # [#] If COUNTERBORE didn't detect in the "MillingProgramType09" (ledge_counterbore_diameter still equal <0>),|
        #     check "MillingProgramType10" Note.                                                                      |
        # [#] Same Description of "MillingProgramType01" with changing the 'Type Number'                              |
        # ============================================================================================================|
        if (ledge_counterbore_diameter == 0):
            engine_worx_database_cursor.execute(
                'SELECT MillingProgramType10 FROM SpexPiston_Milling WHERE PistonID = ?', piston_id)
            for data in engine_worx_database_cursor.fetchone():
                if (data is None):
                    ledge_counterbore_diameter = 0
                elif (data is not None and (
                        (data.find('COUNTERBORE') != -1) or (
                        data.find('LARGER AT LEDGE, CUT AT HORIZONTAL') != -1) or
                        (data.find("C'BORE") != -1) or (data.find('COUNTER BORE') != -1) or
                        (data.find('COUNTERBORES') != -1) or (data.find('CBORE') != -1) or (
                                data.find('CNTBOR') != -1))):
                    ledge_counterbore_diameter = ""
                else:
                    ledge_counterbore_diameter = 0

        # ============================================================================================================|
        # [#] If COUNTERBORE didn't detect in the "MillingProgramType10" (ledge_counterbore_diameter still equal <0>),|
        #     check "MillingProgramType11" Note.                                                                      |
        # [#] Same Description of "MillingProgramType01" with changing the 'Type Number'                              |
        # ============================================================================================================|
        if (ledge_counterbore_diameter == 0):
            engine_worx_database_cursor.execute(
                'SELECT MillingProgramType11 FROM SpexPiston_Milling WHERE PistonID = ?', piston_id)
            for data in engine_worx_database_cursor.fetchone():
                if (data is None):
                    ledge_counterbore_diameter = 0
                elif (data is not None and (
                        (data.find('COUNTERBORE') != -1) or (
                        data.find('LARGER AT LEDGE, CUT AT HORIZONTAL') != -1) or
                        (data.find("C'BORE") != -1) or (data.find('COUNTER BORE') != -1) or
                        (data.find('COUNTERBORES') != -1) or (data.find('CBORE') != -1) or (
                                data.find('CNTBOR') != -1))):
                    ledge_counterbore_diameter = ""
                else:
                    ledge_counterbore_diameter = 0

        # ============================================================================================================|
        # [#] If COUNTERBORE didn't detect in the "MillingProgramType11" (ledge_counterbore_diameter still equal <0>),|
        #     check "MillingProgramType12" Note.                                                                      |
        # [#] Same Description of "MillingProgramType01" with changing the 'Type Number'                              |
        # ============================================================================================================|
        if (ledge_counterbore_diameter == 0):
            engine_worx_database_cursor.execute(
                'SELECT MillingProgramType12 FROM SpexPiston_Milling WHERE PistonID = ?', piston_id)
            for data in engine_worx_database_cursor.fetchone():
                if (data is None):
                    ledge_counterbore_diameter = 0
                elif (data is not None and (
                        (data.find('COUNTERBORE') != -1) or (
                        data.find('LARGER AT LEDGE, CUT AT HORIZONTAL') != -1) or
                        (data.find("C'BORE") != -1) or (data.find('COUNTER BORE') != -1) or
                        (data.find('COUNTERBORES') != -1) or (data.find('CBORE') != -1) or (
                                data.find('CNTBOR') != -1))):
                    ledge_counterbore_diameter = ""
                else:
                    ledge_counterbore_diameter = 0

        # ============================================================================================================|
        # [#] If COUNTERBORE didn't detect in the "MillingProgramType12" (ledge_counterbore_diameter still equal <0>),|
        #     check "MillingProgramType13" Note.                                                                      |
        # [#] Same Description of "MillingProgramType01" with changing the 'Type Number'                              |
        # ============================================================================================================|
        if (ledge_counterbore_diameter == 0):
            engine_worx_database_cursor.execute(
                'SELECT MillingProgramType13 FROM SpexPiston_Milling WHERE PistonID = ?', piston_id)
            for data in engine_worx_database_cursor.fetchone():
                if (data is None):
                    ledge_counterbore_diameter = 0
                elif (data is not None and (
                        (data.find('COUNTERBORE') != -1) or (
                        data.find('LARGER AT LEDGE, CUT AT HORIZONTAL') != -1) or
                        (data.find("C'BORE") != -1) or (data.find('COUNTER BORE') != -1) or
                        (data.find('COUNTERBORES') != -1) or (data.find('CBORE') != -1) or (
                                data.find('CNTBOR') != -1))):
                    ledge_counterbore_diameter = ""
                else:
                    ledge_counterbore_diameter = 0

        # ============================================================================================================|
        # [#] If COUNTERBORE didn't detect in the "MillingProgramType13" (ledge_counterbore_diameter still equal <0>),|
        #     check "MillingProgramType14" Note.                                                                      |
        # [#] Same Description of "MillingProgramType01" with changing the 'Type Number'                              |
        # ============================================================================================================|
        if (ledge_counterbore_diameter == 0):
            engine_worx_database_cursor.execute(
                'SELECT MillingProgramType14 FROM SpexPiston_Milling WHERE PistonID = ?', piston_id)
            for data in engine_worx_database_cursor.fetchone():
                if (data is None):
                    ledge_counterbore_diameter = 0
                elif (data is not None and (
                        (data.find('COUNTERBORE') != -1) or (
                        data.find('LARGER AT LEDGE, CUT AT HORIZONTAL') != -1) or
                        (data.find("C'BORE") != -1) or (data.find('COUNTER BORE') != -1) or
                        (data.find('COUNTERBORES') != -1) or (data.find('CBORE') != -1) or (
                                data.find('CNTBOR') != -1))):
                    ledge_counterbore_diameter = ""
                else:
                    ledge_counterbore_diameter = 0
        print("[#]ledge_counterbore_diameter FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", ledge_counterbore_diameter)

        # endregion <<<<============================[Ledge Counterbore Diameter]============================>>>>

        # MAYBE NEED VARIABLE OF DISTANCE OF LEDGE_COUNTERBORE TO LOCKRING HERE AND IN THE HORIZ TEMPLATE

        # region <<<<=========[X_Distance Of 375_Slots From Center Of Bore To Center Of Horizontal Slot]==========>>>>

        # ==========================================================================================================|
        # [#] Define Variable to set the X_distance from Center of Bore to Center of H-slot (that has 0.375 Radius).|
        # [#] The Value will be Calculated by Math Later.                                                           |
        # [#]                          **Still Needs Work**                                                         |
        # ==========================================================================================================|
        global X_distance_of_375_slots_from_center_of_bore_to_center_of_horizontal_slot
        X_distance_of_375_slots_from_center_of_bore_to_center_of_horizontal_slot = 0

        # endregion <<<<=======[X_Distance Of 375_Slots From Center Of Bore To Center Of Horizontal Slot]=========>>>>

        # region <<<<=========[Y_Distance Of 375_Slots From Center Of Bore To Center Of Horizontal Slot]==========>>>>

        # ==========================================================================================================|
        # [#] Define Variable to set the Y_distance from Center of Bore to Center of H-slot (that has 0.375 Radius).|
        # [#] The Value will be Calculated by Math Later.                                                           |
        # [#]                          **Still Needs Work**                                                         |
        # ==========================================================================================================|
        global Y_distance_of_375_slots_from_center_of_bore_to_center_of_horizontal_slot
        Y_distance_of_375_slots_from_center_of_bore_to_center_of_horizontal_slot = 0

        # endregion <<<<=======[Y_Distance Of 375_Slots From Center Of Bore To Center Of Horizontal Slot]=========>>>

        # region <<<<==============================[Pressure Fed Oil Hole Type]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Type of Pressure Fed Oil Hole.                                      |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_GasPortsPistonPinOiling                                            |
        #     Column Name : PressureFedOilHoleType                                                       |
        # ===============================================================================================|
        global pressure_fed_oil_hole_type
        engine_worx_database_cursor.execute(
            'SELECT PressureFedOilHoleType FROM SpexPiston_GasPortsPistonPinOiling WHERE PistonID = ?', piston_id)

        # =========================================================================================|
        # [#] Define Variable to set Status of PressureFedHolesAvailability to use it later to warn|
        #     the user if the job has H-Slot to check if needs to do them manually.                |
        # =========================================================================================|
        global pressure_fed_holes_availability_status

        # ======================================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                                     |
        # [#] If no data found or it stored as 'String':                                                        |
        #     -Set the value of "pressure_fed_oil_hole_type" as it is stored in the DataBase, and               |
        #     -Set "pressure_fed_holes_availability_status" equal to <0>.                                       |
        # [#] Otherwise:                                                                                        |
        #     -Set the value of "pressure_fed_oil_hole_type" as it is stored in the DataBase(Numeric Value), and|
        #     -Set "pressure_fed_holes_availability_status" equal to <1>.                                       |
        # [#] Expected Data output Type : 'Numeric'                                                             |
        # [#] Expected Data output value : Numeric Value [0-16] or [99]                                         |
        # ======================================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if (data is None or (type(data) == str)):
                pressure_fed_oil_hole_type = data
                pressure_fed_holes_availability_status = 0
            else:
                pressure_fed_oil_hole_type = data
                pressure_fed_holes_availability_status = 1
        print("[#]pressure_fed_oil_hole_type FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", pressure_fed_oil_hole_type)

        # endregion <<<<============================[Pressure Fed Oil Hole Type]============================>>>>

        # region <<<<==============================[Pressure Fed Oil Hole Notes]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Type of Pressure Fed Oil Hole Notes (just in case it's needed).     |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_GasPortsPistonPinOiling                                            |
        #     Column Name : PressureFedOilHoleNotes                                                      |
        # ===============================================================================================|
        global pressure_fed_oil_hole_notes
        engine_worx_database_cursor.execute(
            'SELECT PressureFedOilHoleNotes FROM SpexPiston_GasPortsPistonPinOiling WHERE PistonID = ?', piston_id)

        # ===============================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                              |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',       |
        #     otherwise, Round the Numerical Value for 4-Digits.                                         |
        # [#] Expected Data output Type : 'String'                                                       |
        # [#] Expected Data output value : String Value contains 'text' related in Pressure Fed Oil Hole.|
        # ===============================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                pressure_fed_oil_hole_notes = data
            else:
                pressure_fed_oil_hole_notes = round(float(data), 4)
        print("[#]pressure_fed_oil_hole_notes FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", pressure_fed_oil_hole_notes)

        # endregion <<<<============================[Pressure Fed Oil Hole Notes]============================>>>>

        # region <<<<==============================[Piston Overall Length]==============================>>>>

        # ====================================================================================================|
        # [#] Define Variable to set Overall Length of Piston.                                                |
        # [#] Piston Overall Length stored in the DataBase in different fields (Location), because of that and|
        #     to make sure there is always value of Overall Length, need to check all of them in the DataBase.|
        # [#] Needs to make sure to use same variable name when adding lathe machines.                        |
        # ====================================================================================================|
        global piston_overall_length

        # =========================================================================================================|
        # [#] Get the Value of Overall Length by looking to the first location in the DataBase by using 'piston_id'|
        #     and Access the Data location by using:                                                               |
        #     -Table Name : SpexPiston_Milling                                                                     |
        #     -Column Name : SkirtMillExhaustOverallLengthFromTE                                                   |
        # =========================================================================================================|
        engine_worx_database_cursor.execute(
            'SELECT SkirtMillExhaustOverallLengthFromTE FROM SpexPiston_Milling WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'Numeric'                                               |
        # [#] Expected Data output value : Numeric Value > [0.????]                               |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                piston_overall_length = data
            else:
                piston_overall_length = round(float(data), 4)

        # =====================================================================================================|
        # [#] If Overall Length didn't detect in the location above (Overall Length still equal <0> or <None>),|
        #     Check 'SkirtMillIntakeOverallLengthFromTE' Location.                                             |
        # [#] Get the Value of Overall Length by looking to the location in the DataBase by using 'piston_id'  |
        #     and Access the Data location by using:                                                           |
        #     -Table Name : SpexPiston_Milling                                                                 |
        #     -Column Name : SkirtMillIntakeOverallLengthFromTE                                                |
        # =====================================================================================================|
        if (piston_overall_length is None or piston_overall_length == 0):
            engine_worx_database_cursor.execute(
                'SELECT SkirtMillIntakeOverallLengthFromTE FROM SpexPiston_Milling WHERE PistonID = ?', piston_id)

            # ========================================================================================|
            # [#] Use for loop to iterate through the DateBase.                                       |
            # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
            #     otherwise, Round the Numerical Value for 4-Digits.                                  |
            # [#] Expected Data output Type : 'Numeric'                                               |
            # [#] Expected Data output value : Numeric Value > [0.????]                               |
            # ========================================================================================|
            for data in engine_worx_database_cursor.fetchone():
                if ((data is None) or (type(data) == str)):
                    piston_overall_length = data
                else:
                    piston_overall_length = round(float(data), 4)

        # =====================================================================================================|
        # [#] If Overall Length didn't detect in the location above (Overall Length still equal <0> or <None>),|
        #     Check 'OverallLengthExh' Location.                                                               |
        # [#] Get the Value of Overall Length by looking to the location in the DataBase by using 'piston_id'  |
        #     and Access the Data location by using:                                                           |
        #     -Table Name : SpexPiston_SemiFinishTurn                                                          |
        #     -Column Name : OverallLengthExh                                                                  |
        # =====================================================================================================|
        if (piston_overall_length is None or piston_overall_length == 0):
            engine_worx_database_cursor.execute(
                'SELECT OverallLengthExh FROM SpexPiston_SemiFinishTurn WHERE PistonID = ?', piston_id)

            # ========================================================================================|
            # [#] Use for loop to iterate through the DateBase.                                       |
            # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
            #     otherwise, Round the Numerical Value for 4-Digits.                                  |
            # [#] Expected Data output Type : 'Numeric'                                               |
            # [#] Expected Data output value : Numeric Value > [0.????]                               |
            # ========================================================================================|
            for data in engine_worx_database_cursor.fetchone():
                if ((data is None) or (type(data) == str)):
                    piston_overall_length = data
                else:
                    piston_overall_length = round(float(data), 4)

        # =====================================================================================================|
        # [#] If Overall Length didn't detect in the location above (Overall Length still equal <0> or <None>),|
        #     Check 'OverallLengthInt' Location.                                                               |
        # [#] Get the Value of Overall Length by looking to the location in the DataBase by using 'piston_id'  |
        #     and Access the Data location by using:                                                           |
        #     -Table Name : SpexPiston_SemiFinishTurn                                                          |
        #     -Column Name : OverallLengthInt                                                                  |
        # =====================================================================================================|
        if (piston_overall_length is None or piston_overall_length == 0):
            engine_worx_database_cursor.execute(
                'SELECT OverallLengthInt FROM SpexPiston_SemiFinishTurn WHERE PistonID = ?', piston_id)

            # ========================================================================================|
            # [#] Use for loop to iterate through the DateBase.                                       |
            # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
            #     otherwise, Round the Numerical Value for 4-Digits.                                  |
            # [#] Expected Data output Type : 'Numeric'                                               |
            # [#] Expected Data output value : Numeric Value > [0.????]                               |
            # ========================================================================================|
            for data in engine_worx_database_cursor.fetchone():
                if ((data is None) or (type(data) == str)):
                    piston_overall_length = data
                else:
                    piston_overall_length = round(float(data), 4)
        print("[#]piston_overall_length FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", piston_overall_length)

        # endregion <<<<============================[Piston Overall Length]============================>>>>

        # region <<<<==============================[Engine Stroke Type]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Type of Engine Stroke.                                              |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston                                                                    |
        #     Column Name : EngineStrokeType                                                             |
        # ===============================================================================================|
        global engine_stroke_type
        engine_worx_database_cursor.execute(
            'SELECT EngineStrokeType FROM SpexPiston WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'String'                                                |
        # [#] Expected Data output value : String Value: '4 Cycle' or '2 Cycle'                   |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                engine_stroke_type = data
            else:
                engine_stroke_type = round(float(data), 4)
        print("[#]engine_stroke_type FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", engine_stroke_type)

        # endregion <<<<============================[Engine Stroke Type]============================>>>>

        # region <<<<==============================[Legacy PinBore Comments]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Old Comments of PinBore.                                            |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore                                                            |
        #     Column Name : LegacyPinBoreComments                                                        |
        # ===============================================================================================|
        global legacy_pin_bore_comments
        engine_worx_database_cursor.execute(
            'SELECT LegacyPinBoreComments FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'String'.                                               |
        # [#] Expected Data output value : String Value contains PinBore Info                     |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                legacy_pin_bore_comments = data
            else:
                legacy_pin_bore_comments = round(float(data), 4)
        print("[#]legacy_pin_bore_comments FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", legacy_pin_bore_comments)

        # endregion <<<<============================[Legacy PinBore Comments]============================>>>>

        # region <<<<==============================[Pilot Bore Notes]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Comments of PilotBore (just in case it's needed).                   |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore.                                                           |
        #     Column Name : PilotBoreNotes.                                                              |
        # ===============================================================================================|
        global pilot_bore_notes
        engine_worx_database_cursor.execute(
            'SELECT PilotBoreNotes FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'String'.                                               |
        # [#] Expected Data output value : String Value contains PilotBore Info.                  |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                pilot_bore_notes = data
            else:
                pilot_bore_notes = round(float(data), 4)
        print("[#]pilot_bore_notes FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", pilot_bore_notes)

        # endregion <<<<============================[Pilot Bore Notes]============================>>>>

        # region <<<<==============================[Piston Pin Bore Notes]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Comments of Piston Pin Bore (just in case it's needed).             |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore.                                                           |
        #     Column Name : PistPinBoreNotes.                                                            |
        # ===============================================================================================|
        global piston_pin_bore_notes
        engine_worx_database_cursor.execute(
            'SELECT PistPinBoreNotes FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'String'.                                               |
        # [#] Expected Data output value : String Value contains PistonPinBore Info.              |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            # print(i)
            if ((data is None) or (type(data) == str)):
                piston_pin_bore_notes = data
            else:
                piston_pin_bore_notes = round(float(data), 4)
        print("[#]piston_pin_bore_notes FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", piston_pin_bore_notes)

        # endregion <<<<============================[Piston Pin Bore Notes]============================>>>>

        # region <<<<==============================[Lock Ring Notes]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Comments of Lock Ring (just in case it's needed).                   |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore.                                                           |
        #     Column Name : RetClipNotes.                                                                |
        # ===============================================================================================|
        global ret_clip_grv_notes
        engine_worx_database_cursor.execute(
            'SELECT RetClipNotes FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'String'.                                               |
        # [#] Expected Data output value : String Value contains LockRing Info.                   |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                ret_clip_grv_notes = data
            else:
                ret_clip_grv_notes = round(float(data), 4)
        print("[#]ret_clip_grv_notes FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", ret_clip_grv_notes)

        # endregion <<<<============================[Lock Ring Notes]============================>>>>

        # region <<<<==============================[C/Fren Grv Notes]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Comments of C/Fren (just in case it's needed).                      |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore.                                                           |
        #     Column Name : CFrenNotes.                                                                  |
        # ===============================================================================================|
        global cfren_grv_notes
        engine_worx_database_cursor.execute(
            'SELECT CFrenNotes FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'String'.                                               |
        # [#] Expected Data output value : String Value contains C/Fren Info.                     |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                cfren_grv_notes = data
            else:
                cfren_grv_notes = round(float(data), 4)
        print("[#]cfren_grv_notes FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", cfren_grv_notes)

        # endregion <<<<============================[C/Fren Grv Notes]============================>>>>

        # region <<<<==============================[Semi C/Fren Grv Notes]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Comments of Semi C/Fren (just in case it's needed).                 |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore.                                                           |
        #     Column Name : SemiCFrenGrvNotes.                                                           |
        # ===============================================================================================|
        global semi_cfren_grv_notes
        engine_worx_database_cursor.execute(
            'SELECT SemiCFrenGrvNotes FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'String'.                                               |
        # [#] Expected Data output value : String Value contains Semi C/Fren Info.                |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                semi_cfren_grv_notes = data
            else:
                semi_cfren_grv_notes = round(float(data), 4)
        print("[#]semi_cfren_grv_notes FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", semi_cfren_grv_notes)

        # endregion <<<<============================[Semi C/Fren Grv Notes]============================>>>>

        # region <<<<==============================[Ret Clip Notch Notes]==============================>>>>

        # ===============================================================================================|
        # [#] Define Variable to set Comments of RetClipNotch (just in case it's needed).                |
        # [#] Get the Value from the DataBase by using 'piston_id' and Access the Data location by using:|
        #     Table Name : SpexPiston_PinBore.                                                           |
        #     Column Name : RetClipNotchNotes.                                                           |
        # ===============================================================================================|
        global ret_clip_notch_notes
        engine_worx_database_cursor.execute(
            'SELECT RetClipNotchNotes FROM SpexPiston_PinBore WHERE PistonID = ?', piston_id)

        # ========================================================================================|
        # [#] Use for loop to iterate through the DateBase.                                       |
        # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
        #     otherwise, Round the Numerical Value for 4-Digits.                                  |
        # [#] Expected Data output Type : 'String'.                                               |
        # [#] Expected Data output value : String Value contains RetClipNotch Info.               |
        # ========================================================================================|
        for data in engine_worx_database_cursor.fetchone():
            if ((data is None) or (type(data) == str)):
                ret_clip_notch_notes = data
            else:
                ret_clip_notch_notes = round(float(data), 4)
        print("[#]ret_clip_notch_notes FOR " + new_program_number_for_old_horizontal_machine + ":")
        print("     ", ret_clip_notch_notes)

        # endregion <<<<============================[Ret Clip Notch Notes]============================>>>>

        # region <<<<==============================[Horiz Piston Pin Slots Notes]==============================>>>>
        # ....................It's Done Above At <<Horizontal Slots OD Spacing Section>>...........................
        # endregion <<<<============================[Horiz Piston Pin Slots Notes]=============================>>>>

    # ===============================================================================================|
    # [#] Use (except) block to Handle any Error may occur when accessing the Pin Bore Information in|
    #     the DataBase and avoid APP crash.                                                          |
    # [#] Use (return) to stop the function (APP as a result) from executing the rest of the code.   |
    # [#] Call (failed_to_create_old_horizontal_machine_program) Function to warn the user.          |
    # ===============================================================================================|
    except Exception as error:
        print("exception BLOCK in (four_cycle_pin_bore_variables) Function <Spec Info>")
        fail_messages_of_creating_old_horizontal_machine_program.append(
            "Failed to Find or Detect" + '[b][u][color=ffffff] Pin Bore Information. [/color][/u][/b]' +
            "\n" + "An Error has occurred : " + "\n" + '[color=ff1a1a]' + str(error) + '[/color]' + "\n" +
            "Double Check if there is a Data for this Job in Engine Worx." +
            "\n" + "Work with Engineering to fix the Issue.")
        email_messages_of_creating_old_horizontal_machine_program.append(
            "Failed to Find or Detect PinBore Information." + "\n" + "An Error has occurred : " + "\n" + str(error) +
            "\n" + "Double Check if there is a Data for this Job in Engine Worx." +
            "\n" + "Work with Engineering to fix the Issue." + "\n")
        failed_to_create_old_horizontal_machine_program(self)
        return

    # endregion <<<<============================[Piston Information]=============================>>>>

    # ===================================================================================================|
    # [#] Define List to contain the confirmation details that was needed to create the program.         |
    # [#] Set the list here to avoid reset the list when call (create_program_for_old_horizontal_machine)|
    #     function after finalize the confirmation details.                                              |
    # ===================================================================================================|
    global old_horizontal_program_confirmation_email_message_list
    old_horizontal_program_confirmation_email_message_list = []

    # =================================================================================================|
    # [#] Define Variable to set Status of needing confirmation to create the program.                 |
    # [#] Set the value to be "False" by default assuming no confirmation needed to create programs.   |
    # [#] Change the value to be "True" when find confirmation details that's need to finalize by User.|
    # =================================================================================================|
    global called_need_confirmation_to_create_old_horizontal_machine_program
    called_need_confirmation_to_create_old_horizontal_machine_program = False
    print("called_need_confirmation_to_create_old_horizontal_machine_program in Four Cycle: ",
          called_need_confirmation_to_create_old_horizontal_machine_program)

    # region <<<<==============================[Forging Information]==============================>>>>

    # =================================================================================|
    # [#] Set all Variable that related on Forging Info and get them from the DataBase.|
    # [#] Define all Forging Variables above to be able to use them in different Places|
    #     in the Code even if don't get their Data.                                    |
    # =================================================================================|
    global forging_number
    global forge_spec_id
    global forge_ref_length
    global forging_diameter
    global forging_diameter_OD_at_rougher
    global forging_outside_boss_spacing
    global forging_inside_boss_spacing
    global forging_out_side_ring_belt_height
    global forging_hollow_dome_rise
    global forging_type

    # region <<<<==============================[Forging Number]==============================>>>>

    # =================================================================================================================|
    # [#] Steps to Set the Forging Number:                                                                             |
    #     -First, needs to get 'ForgeSpecID' of Piston from 'SpexPiston' table.                                        |
    #     -Then, use 'ForgeSpecID' to get 'ForgeItemID' from 'SpexForge' table.                                        |
    # [#] 'ForgeSpecID': is the unique ID for each Forging stored in the DataBase.                                     |
    # [#] 'ForgeItemID': is the Forging Number of the job (Ex: F6601X, F6064Z, F555Z. FJE426B-HEX...Etc).              |
    # [#] Check 'piston_id' if equal <1> (that has set earlier in 'create_program_for_old_horizontal_machine' Function)|
    #     to indicate any job (that user entered) has no 'piston_id' (when user enter nothing, 'Test', or              |
    #     any job number is not exist in the DataBase).                                                                |
    # [#] Use (try/except) Blocks to Handle any Error may occur when connecting with the DataBase,                     |
    #     Error could happen if job is missing the Forging Number (many old jobs have this issue)                      |
    # =================================================================================================================|
    if (piston_id != 1):
        try:
            print("try in forging number")
            # ==============================================================================================|
            # [#] First, needs to get Forging ID by using 'piston_id' and Access the Data location by using:|
            #     Table Name : SpexPiston.                                                                  |
            #     Column Name : ForgeSpecID.                                                                |
            # ==============================================================================================|
            engine_worx_database_cursor.execute(
                'SELECT ForgeSpecID FROM SpexPiston WHERE PistonID = ?', piston_id)
            for data in engine_worx_database_cursor.fetchone():
                if ((data is None) or (type(data) == str)):
                    forge_spec_id = data
                else:
                    forge_spec_id = round(float(data), 4)
            print("[#]forge_spec_id FOR " + new_program_number_for_old_horizontal_machine + ":")
            print("     ", forge_spec_id)

            # ===================================================================================================|
            # [#] Then, needs to get Forging Number by using 'ForgeSpecID' and Access the Data location by using:|
            #     Table Name : SpexForge.                                                                        |
            #     Column Name : ForgeItemID.                                                                     |
            # ===================================================================================================|
            engine_worx_database_cursor.execute(
                'SELECT ForgeItemID FROM SpexForge WHERE ForgeSpecID = ?', forge_spec_id)
            for data in engine_worx_database_cursor.fetchone():
                if ((data is None) or (type(data) == str)):
                    forging_number = data
                else:
                    forging_number = round(float(data), 4)
            print("[#]forging_number FOR " + new_program_number_for_old_horizontal_machine + ":")
            print("     ", forging_number)

        # =====================================================================================================|
        # [#] Use (except) block to Handle any Error may occur when accessing the DataBase and avoid APP crash.|
        # [#] Use (return) to stop the function (APP as a result) from executing the rest of the code.         |
        # =====================================================================================================|
        except Exception as error:
            print("exception in forging number")
            fail_messages_of_creating_old_horizontal_machine_program.append(
                "Failed to Find or Detect" + '[b][u][color=ffffff] Forging Number [/color][/u][/b]' +
                "\n" + "An Error has occurred : " + "\n" + '[color=ff1a1a]' + str(error) + '[/color]' + "\n" +
                "Double Check if the Job has NO Forging Info, Nor Forging NOT listed" + "\n" + "on the DataBase. " +
                "\n" + "Work with Engineering to fix the Issue.")
            email_messages_of_creating_old_horizontal_machine_program.append(
                "Failed to Find or Detect Forging Number" + "\n" + "An Error has occurred : " + "\n" + str(error) +
                "\n" + "Double Check if the Job has NO Forging Info, " + "Nor Forging NOT listed on the DataBase. " +
                "\n" + "Work with Engineering to fix the Issue." + "\n")

            # ========================================================================================|
            # [#] Set All Forging Variable to be 'None' In case Error has occurred to avoid APP crash.|
            # [#] Call (failed_to_create_old_horizontal_machine_program) Function to warn the user.   |
            # ========================================================================================|
            forging_number = None
            forge_ref_length = None
            forging_diameter = None
            forging_diameter_OD_at_rougher = None
            forging_outside_boss_spacing = None
            forging_inside_boss_spacing = None
            forging_out_side_ring_belt_height = None
            forging_hollow_dome_rise = None
            forging_type = None
            failed_to_create_old_horizontal_machine_program(self)
            return

    # =============================================================================================|
    # [#] If 'piston_id' is equal <1> (that indicate the job (that user entered) has no 'piston_id'|
    #     (when user enter nothing, 'Test', or any job number is not exist in the DataBase).       |
    # [#] Set All Forging Variable to be 'None' to avoid APP crash.                                |
    # [#] Use (return) to stop the function (APP as a result) from executing the rest of the code. |
    # =============================================================================================|
    else:
        print("else Statment in forging number")
        forging_number = None
        forge_ref_length = None
        forging_diameter = None
        forging_diameter_OD_at_rougher = None
        forging_outside_boss_spacing = None
        forging_inside_boss_spacing = None
        forging_out_side_ring_belt_height = None
        forging_hollow_dome_rise = None
        forging_type = None
        return

    # endregion <<<<============================[Forging Number]============================>>>>

    # region <<<<==============================[Forge Ref Length]==============================>>>>

    # =================================================================================================|
    # [#] Get the Value from the DataBase by using 'ForgeItemID' and Access the Data location by using:|
    #     Table Name : SpexForge.                                                                      |
    #     Column Name : ForgeRefLength.                                                                |
    # =================================================================================================|
    engine_worx_database_cursor.execute(
        'SELECT ForgeRefLength FROM SpexForge WHERE ForgeItemID = ?', forging_number)

    # ========================================================================================|
    # [#] Use for loop to iterate through the DateBase.                                       |
    # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
    #     otherwise, Round the Numerical Value for 4-Digits.                                  |
    # [#] Expected Data output Type : 'Numeric'                                               |
    # [#] Expected Data output value : Numeric Value > [0.????]                               |
    # ========================================================================================|
    for data in engine_worx_database_cursor.fetchone():
        if ((data is None) or (type(data) == str)):
            forge_ref_length = data
        else:
            forge_ref_length = round(float(data), 4)
    print("[#]forge_ref_length FOR " + new_program_number_for_old_horizontal_machine + ":")
    print("     ", forge_ref_length)

    # endregion <<<<============================[Forge Ref Length]============================>>>>

    # region <<<<==============================[Forging Diameter]==============================>>>>

    # =================================================================================================|
    # [#] Get the Value from the DataBase by using 'ForgeItemID' and Access the Data location by using:|
    #     Table Name : SpexForge.                                                                      |
    #     Column Name : ForgeOD.                                                                       |
    # =================================================================================================|
    engine_worx_database_cursor.execute(
        'SELECT ForgeOD FROM SpexForge WHERE ForgeItemID = ?', forging_number)

    # ========================================================================================|
    # [#] Use for loop to iterate through the DateBase.                                       |
    # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
    #     otherwise, Round the Numerical Value for 4-Digits.                                  |
    # [#] Expected Data output Type : 'Numeric'                                               |
    # [#] Expected Data output value : Numeric Value > [0.????]                               |
    # ========================================================================================|
    for data in engine_worx_database_cursor.fetchone():
        if ((data is None) or (type(data) == str)):
            forging_diameter = data
        else:
            forging_diameter = round(float(data), 4)
    print("[#]forging_diameter FOR " + new_program_number_for_old_horizontal_machine + ":")
    print("     ", forging_diameter)

    # endregion <<<<============================[Forging Diameter]============================>>>>

    # region <<<<==============================[Forging Diameter At Rougher]==============================>>>>

    # =================================================================================================|
    # [#] Get the Value from the DataBase by using 'ForgeItemID' and Access the Data location by using:|
    #     Table Name : SpexForge.                                                                      |
    #     Column Name : ODAtRougher.                                                                   |
    # =================================================================================================|
    engine_worx_database_cursor.execute(
        'SELECT ODAtRougher FROM SpexForge WHERE ForgeItemID = ?', forging_number)

    # ========================================================================================|
    # [#] Use for loop to iterate through the DateBase.                                       |
    # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
    #     otherwise, Round the Numerical Value for 4-Digits.                                  |
    # [#] Expected Data output Type : 'Numeric'                                               |
    # [#] Expected Data output value : Numeric Value > [0.????]                               |
    # ========================================================================================|
    for data in engine_worx_database_cursor.fetchone():
        if ((data is None) or (type(data) == str)):
            forging_diameter_OD_at_rougher = data
        else:
            forging_diameter_OD_at_rougher = round(float(data), 4)
    print("[#]forging_diameter_OD_at_rougher FOR " + new_program_number_for_old_horizontal_machine + ":")
    print("     ", forging_diameter_OD_at_rougher)

    # endregion <<<<============================[Forging Diameter At Rougher]============================>>>>

    # region <<<<==============================[Forging Boss Outside Spacing]==============================>>>>

    # =================================================================================================|
    # [#] Get the Value from the DataBase by using 'ForgeItemID' and Access the Data location by using:|
    #     Table Name : SpexForge.                                                                      |
    #     Column Name : BossOutsdSpace.                                                                |
    # =================================================================================================|
    engine_worx_database_cursor.execute(
        'SELECT BossOutsdSpace FROM SpexForge WHERE ForgeItemID = ?', forging_number)

    # ========================================================================================|
    # [#] Use for loop to iterate through the DateBase.                                       |
    # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
    #     otherwise, Round the Numerical Value for 4-Digits.                                  |
    # [#] Expected Data output Type : 'Numeric'                                               |
    # [#] Expected Data output value : Numeric Value > [0.????]                               |
    # ========================================================================================|
    for data in engine_worx_database_cursor.fetchone():
        if ((data is None) or (type(data) == str)):
            forging_outside_boss_spacing = data
        else:
            forging_outside_boss_spacing = round(float(data), 4)
    print("[#]forging_outside_boss_spacing FOR " + new_program_number_for_old_horizontal_machine + ":")
    print("     ", forging_outside_boss_spacing)

    # endregion <<<<============================[Forging Boss Outside Spacing]============================>>>>

    # region <<<<==============================[Forging Boss Inside Spacing]==============================>>>>

    # =================================================================================================|
    # [#] Get the Value from the DataBase by using 'ForgeItemID' and Access the Data location by using:|
    #     Table Name : SpexForge.                                                                      |
    #     Column Name : BossInsdSpace.                                                                 |
    # =================================================================================================|
    engine_worx_database_cursor.execute(
        'SELECT BossInsdSpace FROM SpexForge WHERE ForgeItemID = ?', forging_number)

    # ========================================================================================|
    # [#] Use for loop to iterate through the DateBase.                                       |
    # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
    #     otherwise, Round the Numerical Value for 4-Digits.                                  |
    # [#] Expected Data output Type : 'Numeric'                                               |
    # [#] Expected Data output value : Numeric Value > [0.????]                               |
    # ========================================================================================|
    for data in engine_worx_database_cursor.fetchone():
        if ((data is None) or (type(data) == str)):
            forging_inside_boss_spacing = data
        else:
            forging_inside_boss_spacing = round(float(data), 4)
    print("[#]forging_inside_boss_spacing FOR " + new_program_number_for_old_horizontal_machine + ":")
    print("     ", forging_inside_boss_spacing)

    # endregion <<<<============================[Forging Boss Inside Spacing]============================>>>>

    # region <<<<==============================[Forging Out Side Ring_Belt Height]==============================>>>>

    # =================================================================================================|
    # [#] Get the Value from the DataBase by using 'ForgeItemID' and Access the Data location by using:|
    #     Table Name : SpexForge.                                                                      |
    #     Column Name : RingBeltHtOutsd.                                                               |
    # [#] is the distance from origin to end point before the boss.                                    |
    # =================================================================================================|
    engine_worx_database_cursor.execute(
        'SELECT RingBeltHtOutsd FROM SpexForge WHERE ForgeItemID = ?', forging_number)

    # ========================================================================================|
    # [#] Use for loop to iterate through the DateBase.                                       |
    # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
    #     otherwise, Round the Numerical Value for 4-Digits.                                  |
    # [#] Expected Data output Type : 'Numeric'                                               |
    # [#] Expected Data output value : Numeric Value > [0.????]                               |
    # ========================================================================================|
    for data in engine_worx_database_cursor.fetchone():
        if ((data is None) or (type(data) == str)):
            forging_out_side_ring_belt_height = data
        else:
            forging_out_side_ring_belt_height = round(float(data), 4)
    print("[#]forging_out_side_ring_belt_height FOR " + new_program_number_for_old_horizontal_machine + ":")
    print("     ", forging_out_side_ring_belt_height)

    # endregion <<<<============================[Forging Out Side Ring_Belt Height]============================>>>>

    # region <<<<==============================[Forging Hollow Dome Rise]==============================>>>>

    # =================================================================================================|
    # [#] Get the Value from the DataBase by using 'ForgeItemID' and Access the Data location by using:|
    #     Table Name : SpexForge.                                                                      |
    #     Column Name : HDRise.                                                                        |
    # =================================================================================================|
    engine_worx_database_cursor.execute(
        'SELECT HDRise FROM SpexForge WHERE ForgeItemID = ?', forging_number)

    # ========================================================================================|
    # [#] Use for loop to iterate through the DateBase.                                       |
    # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
    #     otherwise, Round the Numerical Value for 4-Digits.                                  |
    # [#] Expected Data output Type : 'Numeric'                                               |
    # [#] Expected Data output value : Numeric Value > [0.????]                               |
    # ========================================================================================|
    for data in engine_worx_database_cursor.fetchone():
        if ((data is None) or (type(data) == str)):
            forging_hollow_dome_rise = data
        else:
            forging_hollow_dome_rise = round(float(data), 4)
    print("[#]forgeing_hollow_dome_rise FOR " + new_program_number_for_old_horizontal_machine + ":")
    print("     ", forging_hollow_dome_rise)

    # endregion <<<<============================[Forging Hollow Dome Rise]============================>>>>

    # region <<<<==============================[Forging Type <<4 or 2 Cycle>>]==============================>>>>

    # =================================================================================================|
    # [#] Get the Value from the DataBase by using 'ForgeItemID' and Access the Data location by using:|
    #     Table Name : SpexForge.                                                                      |
    #     Column Name : EngineType.                                                                    |
    # =================================================================================================|
    engine_worx_database_cursor.execute(
        'SELECT EngineType FROM SpexForge WHERE ForgeItemID = ?', forging_number)

    # ========================================================================================|
    # [#] Use for loop to iterate through the DateBase.                                       |
    # [#] Set the value as it is stored in the DataBase if no data found or stored as'String',|
    #     otherwise, Round the Numerical Value for 4-Digits.                                  |
    # [#] Expected Data output Type : 'String'                                                |
    # [#] Expected Data output value : String Value: '4CYCLE' or '2CYCLE'                     |
    # ========================================================================================|
    for data in engine_worx_database_cursor.fetchone():
        if ((data is None) or (type(data) == str)):
            forging_type = data
        else:
            forging_type = round(float(data), 4)
    print("[#]forging_type FOR " + new_program_number_for_old_horizontal_machine + ":")
    print("     ", forging_type)

    # endregion <<<<============================[Forging Type <<4 or 2 Cycle>>]============================>>>>

    # endregion <<<<==============================[Forging Information]==============================>>>>

# endregion <<<<==================================[Four Cycle Pin Bore Function]==================================>>>>


# region <<<<=====================================[Pin Bore Machines Screen]=====================================>>>>

class PinBoreScreen(Screen):
    # =====================================================================|
    #  Create Function to NOT open the New Horizontal Machine(127) Screen .|
    # =====================================================================|
    def still_work_on_it(self, obj):
        close_button = MDRaisedButton(text='Close', on_release=self.close_pin_bore_screen_window, font_size=16)
        self.pin_bore_screen_message_window = MDDialog(title='[color=cccc00]Warning Message[/color]', text=(
            '[color=ffffff]Still Work On It, Thanks for your Patience. [/color]'), size_hint=(0.7, 1.0),
                                                       buttons=[close_button], auto_dismiss=False)
        self.pin_bore_screen_message_window.open()

    def close_pin_bore_screen_window(self, obj):
        print("Close_PinBore_Dialog" + " is called")
        self.pin_bore_screen_message_window.dismiss()

# endregion <<<<====================================[Pin Bore Machines Screen]====================================>>>>


# region <<<<============================[Old Horizontal Machines(28,29,32) Items List]===========================>>>>

# ========================================================================================================|
# [#] Create the Class to be able to create List's Items that will be used to let user choose from options|
#     that are needed for some confirmation details that need to finalize by the user.                    |
# ========================================================================================================|
class OldHorizontalMachineItem(OneLineAvatarListItem):
    divider = None

# endregion <<<<==========================[Old Horizontal Machines(28,29,32) Items List]==========================>>>>


# region <<<<===========================[Old Horizontal Machines(28,29,32) Functions]============================>>>>

# ================================================================================|
# [#] Create Function to create Horizontal PinBore program in the Original folder.|
# ================================================================================|
def create_old_horizontal_machine_program_in_original_folder(self):
    print("(create_old_horizontal_machine_program_in_original_folder) Function >> called")

    # =================================================================================================|
    # [#] Note<1,1>                                                                                    |
    # [#] Define Variable to set Status of needing confirmation to create the program.                 |
    # [#] Set the value to be "False" by default assuming no confirmation needed to create programs.   |
    # [#] Change the value to be "True" when find confirmation details that's need to finalize by User.|
    # =================================================================================================|
    global called_need_confirmation_to_create_old_horizontal_machine_program
    called_need_confirmation_to_create_old_horizontal_machine_program = False
    print("called_need_confirmation_to_create_old_horizontal_machine_program on create Program in Original Folder: ",
          called_need_confirmation_to_create_old_horizontal_machine_program)

    # ================================================================================================================|
    # [#] Note<1,2>                                                                                                   |
    # [#] Define Variable to set Path of the Original Folder.                                                         |
    # [#] To set the Path on the 'Test Mode':                                                                         |
    #     Access MDTextField of (id: OriginalFolderPathOfOldHorizontalMachineTestMode) in (OldHorizontalSettingScreen)|
    #     from Screens_Builder to get the File Path.                                                                  |
    # [#] To set the Path on the 'Live Mode':                                                                         |
    #     Access MDTextField of (id: OriginalFolderPathOfOldHorizontalMachine) in (OldHorizontalSettingScreen) from   |
    #     Screens_Builder to get the File Path.                                                                       |
    # ================================================================================================================|
    global original_folder_path_of_old_horizontal_machine
    print("activate_test_mode on Original Folder: ", activate_test_mode)
    if (activate_test_mode == True):
        original_folder_path_of_old_horizontal_machine = self.manager.get_screen('OldHorizontalSettingScreen').ids[
            "OriginalFolderPathOfOldHorizontalMachineTestMode"].text
    elif (activate_test_mode == False):
        original_folder_path_of_old_horizontal_machine = self.manager.get_screen('OldHorizontalSettingScreen').ids[
            "OriginalFolderPathOfOldHorizontalMachine"].text
    print("original_folder_path_of_old_horizontal_machine: ", original_folder_path_of_old_horizontal_machine)

    # ==========================================================================================================|
    # [#] Note<1,3>                                                                                             |
    # [#] Define Variable to set the Program that will saved in the Original Folder.                            |
    # [#] Variable contains: Original Folder Path + "\\" + "P" + New Job Number + ".MIN" (which is the extension|
    #     that make machine able to read the File).                                                             |
    # [#] This Variable used to create Program for Job that have one Specific Offset (Zero,positive or Negative)|
    # ==========================================================================================================|
    global new_horizontal_program_for_old_machine_in_original_folder

    # ================================================================================================================|
    # [#] Note<1,4>                                                                                                   |
    # [#] Define Variable to set the Program <To 0 Direction> that will saved in the Original Folder.                 |
    # [#] Variable contains: Original Folder Path + "\\" + "P" + New Job Number + "TO0" +".MIN" (which is the         |
    #     extension that make machine able to read the File).                                                         |
    # [#] Some Job designed to have Offset Each Way, This Variable used to create Program <To 0 Direction> (Positive).|
    # ================================================================================================================|
    global new_horizontal_program_To0_direction_for_old_machine_in_original_folder

    # =================================================================================================================|
    # [#] Note<1,5>                                                                                                    |
    # [#] Define Variable to set the Program <To 180 Direction> that will saved in the Original Folder.                |
    # [#] Variable contains: Original Folder Path + "\\" + "P" + New Job Number + "TO180" +".MIN" (which is the        |
    #     extension that make machine able to read the File).                                                          |
    # [#] Some Job designed to have Offset Each Way, This Variable used to create Program <To180 Direction> (Negative).|
    # =================================================================================================================|
    global new_horizontal_program_To180_direction_for_old_machine_in_original_folder

    # ================================================================================================================|
    # [#] Note<1,6>                                                                                                   |
    # [#] Define List to contain all Lines Of Horizontal PinBore program <To 0 Direction> that will saved in          |
    #      the Original Folder.                                                                                       |
    # [#] Use 'Copy' Method to copy The Main List (pin_bore_program_lines_of_old_horizontal_machine) that contains the|
    #     Main program lines to the (horizontal_program_lines_To0_direction_for_old_machine_in_original_folder)       |
    #     list to Modify it for Changes that Program <To 0 Direction> are needed.                                     |
    # ================================================================================================================|
    global horizontal_program_lines_To0_direction_for_old_machine_in_original_folder
    horizontal_program_lines_To0_direction_for_old_machine_in_original_folder = \
        pin_bore_program_lines_of_old_horizontal_machine.copy()

    # ================================================================================================================|
    # [#] Note<1,7>                                                                                                   |
    # [#] Define List to contain all Lines Of Horizontal PinBore program <To 180 Direction> that will saved in        |
    #      the Original Folder.                                                                                       |
    # [#] Use 'Copy' Method to copy The Main List (pin_bore_program_lines_of_old_horizontal_machine) that contains the|
    #     Main program lines to the (horizontal_program_lines_To180_direction_for_old_machine_in_original_folder)     |
    #     list to Modify it for Changes that Program <To 180 Direction> are needed.                                   |
    # ================================================================================================================|
    global horizontal_program_lines_To180_direction_for_old_machine_in_original_folder
    horizontal_program_lines_To180_direction_for_old_machine_in_original_folder = \
        pin_bore_program_lines_of_old_horizontal_machine.copy()

    # =============================================================================================================|
    # [#] Note<1,8>                                                                                                |
    # [#] Use (try) Block to try to create Program with certain 'Job Number' in the Original Folder.               |
    # [#] Use (except) Block with specific Exception of (FileExistsError) to indicate the program with this certain|
    #    'Job Number' is Exist in the Original Folder and check with user If needs to save over the Existing file. |
    # =============================================================================================================|
    try:
        print("try old_horizontal in Original Folder is called")
        # ==================================================================================================|
        # [#] Note<1,9>                                                                                     |
        # [#] Check if Offset Direction is 'Each Way', to Create Programs for <To 0> and <To 180> Directions|
        # ==================================================================================================|
        if (offset_direction == "OFFSET EACH WAY" and offset_amount != 0):

            # =======================================================================|
            # [#] Note<1,10>                                                         |
            # [#] Set the Program <To 0 Direction> to contains:                      |
            #     Original Folder Path + "\\" + "P" + New Job Number + "TO0" +".MIN".|
            # =======================================================================|
            new_horizontal_program_To0_direction_for_old_machine_in_original_folder = (
                    original_folder_path_of_old_horizontal_machine + "\\" + "P" +
                    new_program_number_for_old_horizontal_machine + "TO0.MIN")

            # ======================================================================================================|
            # [#] Note<1,11>                                                                                        |
            # [#] Access the First Element of List of (Program <To 0 Direction>) by use ([0] index) to Set the First|
            #     line of the Program with all changes that needed (Ex: (PART WD-13000 TO 0 -- 05/14/2022 <SYS>))   |
            # [#] <SYS>: Used to Indicate this Program is Created By 'WisecoProgramsMaker' APP.                     |
            # [#] (notes_for_old_horizontal_machine_program): To Add any note that needed to the Top of Program.    |
            # ======================================================================================================|
            horizontal_program_lines_To0_direction_for_old_machine_in_original_folder[0] = (
                    '(PART ' + new_program_number_for_old_horizontal_machine + " TO 0" + ' -- ' + today_date +
                    ' <SYS>' + ')' + "\n".join(notes_for_old_horizontal_machine_program))

            # =======================================================================================================|
            # [#] Note<1,12>                                                                                         |
            # [#] Set Variable to create (Program <To 0 Direction>) in the Original Folder.                          |
            # [#] Steps to Create File:                                                                              |
            #    [#] Use <open()> Method To create empty File inside Folder with Parameter "x" to create the file    |
            #        if it's NOT exist on the Folder, if it is Exist, it will returns an error of (FileExistsError)  |
            #    [#] Use [.write()] to Add Content of the List of (Program <To 0 Direction>) to the File that Created|
            #        with Use ['\n'.join()] to Add List Elements of (Program <To 0 Direction>) Line by Line..        |
            #        ..(ie: [Element],new line,[Element]).                                                           |
            #    [#] Use [.close()] Method to close the File (with contents) has been created in Original Folder.    |
            # [#] Use (try/except) Blocks to Handle any Error may occur when Accessing or Finding Files and Folders. |
            # =======================================================================================================|
            try:
                create_new_horizontal_program_To0_direction_for_old_machine_in_original_folder = open(
                    new_horizontal_program_To0_direction_for_old_machine_in_original_folder, "x")
                create_new_horizontal_program_To0_direction_for_old_machine_in_original_folder.write(
                    '\n'.join(horizontal_program_lines_To0_direction_for_old_machine_in_original_folder))
                create_new_horizontal_program_To0_direction_for_old_machine_in_original_folder.close()

            # ===============================================================================================|
            # [#] Note<1,13>                                                                                 |
            # [#] Use (except) Block with specific Exceptions of (PermissionError) and (FileNotFoundError) to|
            #     Handle any Error may occur when Accessing or Finding Files and Folders.                    |
            # [#] Call (failed_to_create_old_horizontal_machine_program) Function to warn the user.          |
            # [#] Use (return) to stop the function (APP as a result) from executing the rest of the code.   |
            # ===============================================================================================|
            except PermissionError or FileNotFoundError as error:
                fail_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find" + '[b][color=ffff00] ORIGINAL [/color][/b]' +
                    "Folder location to Save the Program." + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]'
                    + str(error) + '[/color]' + "\n" + "Double Check Network, and Folder Location." + "\n")
                email_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find Original Folder location to Save the Program." + "\n" +
                    "Double Check Network, and File Location." + "\n")
                failed_to_create_old_horizontal_machine_program(self)
                return

            # =========================================================================|
            # [#] Note<1,14>                                                           |
            # [#] Set the Program <To 180 Direction> to contains:                      |
            #     Original Folder Path + "\\" + "P" + New Job Number + "TO180" +".MIN".|
            # =========================================================================|
            new_horizontal_program_To180_direction_for_old_machine_in_original_folder = (
                    original_folder_path_of_old_horizontal_machine + "\\" + "P" +
                    new_program_number_for_old_horizontal_machine + "TO180.MIN")

            # ========================================================================================================|
            # [#] Note<1,15>                                                                                          |
            # [#] For (Program <To 180 Direction>),needs to make more changes to the main program lines(that's copied)|
            #     [#] Access the First Element of List of (Program <To 180 Direction>) by use ([0] index) to Set      |
            #         the First line of the Program with all changes that needed.                                     |
            #         (Ex: (PART WD-13000 TO 180 -- 05/14/2022 <SYS>)).                                               |
            #        - <SYS>: Used to Indicate this Program is Created By 'WisecoProgramsMaker' APP.                  |
            #        - (notes_for_old_horizontal_machine_program): To Add any note that needed to the Top of Program. |
            # ========================================================================================================|
            horizontal_program_lines_To180_direction_for_old_machine_in_original_folder[0] = (
                    '(PART ' + new_program_number_for_old_horizontal_machine + " TO 180" + ' -- ' + today_date +
                    ' <SYS>' + ')' + "\n".join(notes_for_old_horizontal_machine_program))

            # ========================================================================================================|
            # [#] Note<1,16>                                                                                          |
            #     [#] Access the Element that start with ('VC155=') of List of (Program <To 180 Direction>) by use    |
            #       ([VC155_variable_index] index that's set earlier) to change the Offset value to be Negative.      |
            # ========================================================================================================|
            horizontal_program_lines_To180_direction_for_old_machine_in_original_folder[VC155_variable_index] = (
                    'VC155=' + format((-1) * offset_amount) + '  (Offset)')

            # ======================================================================================================|
            # [#] Note<1,17>                                                                                        |
            #    [#] Access the Element that start with ('VC173=') of List of (Program <To 180 Direction>) by use   |
            #       ([VC173_variable_index] index that's set earlier) to change the <Y_distance of clip notch> value|
            #       to be (Y_distance of circlip notch - <2> _multiply_ <offset_amount>).                           |
            # ======================================================================================================|
            if (notch_angle_first_location != 0 and notch_angle_first_location is not None):
                horizontal_program_lines_To180_direction_for_old_machine_in_original_folder[VC173_variable_index] = (
                    'VC173=' + format(float(Y_distance_from_origin_to_circlip_notch) - (2 * offset_amount), '.4f') +
                    '  (yCirclipNotch)')

            # =======================================================================================================|
            # [#] Note<1,18>                                                                                         |
            # [#] Set Variable to create (Program <To 180 Direction>) in the Original Folder.                        |
            # [#] Steps to Create File:                                                                              |
            #    [#] Use <open()> Method To create empty File inside Folder with Parameter "x" to create the file    |
            #        if it's NOT exist on the Folder, if it is Exist, it will returns an error of (FileExistsError)  |
            #    [#] Use [.write()] to Add Content of the List of (Program <To 180 Direction>) to the File that      |
            #        Created with Use ['\n'.join()] to Add List Elements of (Program <To 0 Direction>) Line by Line..|
            #        ..(ie: [Element],new line,[Element]).                                                           |
            #    [#] Use [.close()] Method to close the File (with contents) has been created in Original Folder.    |
            # [#] Use (try/except) Blocks to Handle any Error may occur when Accessing or Finding Files and Folders. |
            # =======================================================================================================|
            try:
                create_new_horizontal_program_To180_direction_for_old_machine_in_original_folder = open(
                    new_horizontal_program_To180_direction_for_old_machine_in_original_folder, "x")
                create_new_horizontal_program_To180_direction_for_old_machine_in_original_folder.write(
                    '\n'.join(horizontal_program_lines_To180_direction_for_old_machine_in_original_folder))
                create_new_horizontal_program_To180_direction_for_old_machine_in_original_folder.close()

            # ===============================================================================================|
            # [#] Note<1,19>                                                                                 |
            # [#] Use (except) Block with specific Exceptions of (PermissionError) and (FileNotFoundError) to|
            #     Handle any Error may occur when Accessing or Finding Files and Folders.                    |
            # [#] Call (failed_to_create_old_horizontal_machine_program) Function to warn the user.          |
            # [#] Use (return) to stop the function (APP as a result) from executing the rest of the code.   |
            # ===============================================================================================|
            except PermissionError or FileNotFoundError as error:
                fail_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find" + '[b][color=ffff00] ORIGINAL [/color][/b]' +
                    "Folder location to Save the Program." + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]'
                    + str(error) + '[/color]' + "\n" + "Double Check Network, and Folder Location." + "\n")
                email_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find Original Folder location to Save the Program." + "\n" +
                    "Double Check Network, and File Location." + "\n")
                failed_to_create_old_horizontal_machine_program(self)
                return

        # =============================================================================================================|
        # [#] Note<1,20>                                                                                               |
        # [#] If Offset Direction is not 'Each Way', ie: NO need to Create Programs for <To 0> and <To 180> Directions.|
        # =============================================================================================================|
        else:
            # ===============================================================|
            # [#] Note<1,21>                                                 |
            # [#] Set the Program to contains:                               |
            #     Original Folder Path + "\\" + "P" + New Job Number +".MIN".|
            # ===============================================================|
            new_horizontal_program_for_old_machine_in_original_folder = (
                    original_folder_path_of_old_horizontal_machine + "\\" + "P" +
                    new_program_number_for_old_horizontal_machine + ".MIN")

            # ==================================================================================================|
            # [#] Access the First Element of List of (Main Program lines) by use ([0] index) to Set the First  |
            #     line of the Program with all changes that needed (Ex: (PART WD-13000 -- 05/14/2022 <SYS>)).   |
            # [#] <SYS>: Used to Indicate this Program is Created By 'WisecoProgramsMaker' APP.                 |
            # [#] (notes_for_old_horizontal_machine_program): To Add any note that needed to the Top of Program.|
            # ==================================================================================================|

            pin_bore_program_lines_of_old_horizontal_machine[0] = (
                    '(PART ' + new_program_number_for_old_horizontal_machine + ' -- ' + today_date + ' <SYS>' + ')' +
                    "\n".join(notes_for_old_horizontal_machine_program))

            # ======================================================================================================|
            # [#] Note<1,22>                                                                                        |
            # [#] Set Variable to create the Program in the Original Folder.                                        |
            # [#] Steps to Create File:                                                                             |
            #    [#] Use <open()> Method To create empty File inside Folder with Parameter "x" to create the file   |
            #        if it's NOT exist on the Folder, if it is Exist, it will returns an error of (FileExistsError) |
            #    [#] Use [.write()] to Add Content of the List of Program (Main Program lines) to the File that     |
            #        Created with Use ['\n'.join()] to Add List Elements of (Main Program lines) Line by Line...    |
            #        ...(ie: [Element],new line,[Element]).                                                         |
            #    [#] Use [.close()] Method to close the File (with contents) has been created in Original Folder.   |
            # [#] Use (try/except) Blocks to Handle any Error may occur when Accessing or Finding Files and Folders.|
            # ======================================================================================================|
            try:
                create_new_horizontal_program_for_old_machine_in_original_folder = open(
                    new_horizontal_program_for_old_machine_in_original_folder, "x")
                create_new_horizontal_program_for_old_machine_in_original_folder.write(
                    '\n'.join(pin_bore_program_lines_of_old_horizontal_machine))
                create_new_horizontal_program_for_old_machine_in_original_folder.close()

            # ===============================================================================================|
            # [#] Note<1,23>                                                                                 |
            # [#] Use (except) Block with specific Exceptions of (PermissionError) and (FileNotFoundError) to|
            #     Handle any Error may occur when Accessing or Finding Files and Folders.                    |
            # [#] Call (failed_to_create_old_horizontal_machine_program) Function to warn the user.          |
            # [#] Use (return) to stop the function (APP as a result) from executing the rest of the code.   |
            # ===============================================================================================|
            except PermissionError or FileNotFoundError as error:
                fail_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find" + '[b][color=ffff00] ORIGINAL [/color][/b]' +
                    "Folder location to Save the Program." + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]'
                    + str(error) + '[/color]' + "\n" + "Double Check Network, and Folder Location." + "\n")
                email_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find Original Folder location to Save the Program." + "\n" +
                    "Double Check Network, and File Location." + "\n")
                failed_to_create_old_horizontal_machine_program(self)
                return

        # ===================================================================================================|
        # [#] Note<1,24>                                                                                     |
        # [#] Create Message_Window to Inform the User :                                                     |
        #     -Add the Confirmation details that was needed to create the program to the Email messages list.|
        #     -Add the Success message of creating program to the Email messages list.                       |
        # ===================================================================================================|
        if (old_horizontal_program_confirmation_email_message_list != []):
            email_messages_of_creating_old_horizontal_machine_program.append(
                "\n".join(old_horizontal_program_confirmation_email_message_list) + "\n")

        success_messages_of_creating_old_horizontal_machine_program = \
            ["\n" + "Program has been **CREATED** successfully in **ORIGINAL** Folder." + "\n"]
        email_messages_of_creating_old_horizontal_machine_program.append("\n".join(
            success_messages_of_creating_old_horizontal_machine_program))

        # ===============================================================================================|
        # [#] Note<1,25>                                                                                 |
        # [#] Create CloseButton that will call (close_old_horizontal_window_of_original_folder) Function|
        #     when User click on the Button to close the Window and do some actions.                     |
        # [#] (close_old_horizontal_window_of_original_folder) Function Should locate on Class of:       |
        #     [OldHorizontalScreen(Screen)].                                                             |
        # ===============================================================================================|
        close_button = MDRaisedButton(
            text='Close', on_release=self.close_old_horizontal_window_of_original_folder, font_size=16)

        # ==========================================================================|
        # [#] Note<1,26>                                                            |
        # [#] Create Message_Window of the Old Horizontal Screen to Inform the User.|
        # ==========================================================================|
        self.old_horizontal_screen_message_window = MDDialog(title='[color=248f24]Success Message[/color]', text=(
                '[b][i][u][color=0099ff]' + self.ids["JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                '[color=ffffff] Program Has been Created Successfully in [/color]' + '[color=ffff00]ORIGINAL[/color]' +
                '[color=ffffff] Folder.[/color]'), size_hint=(0.7, 1.0), buttons=[close_button], auto_dismiss=False)
        self.old_horizontal_screen_message_window.open()
        print()

    # =============================================================================================================|
    # [#] Note<1,27>                                                                                               |
    # [#] When Program is Exist in the Original Folder.                                                            |
    # [#] Use (except) Block with specific Exception of (FileExistsError) to indicate the program with this certain|
    #    'Job Number' is Exist in the Original Folder and check with user If needs to save over the Existing file. |
    # =============================================================================================================|
    except(FileExistsError):

        # ============================================================================================================|
        # [#] Note<1,28>                                                                                              |
        # [#] Create Yes_Button that will call (replace_existing_old_horizontal_machine_program_in_original_folder)   |
        #     Function when User click on the Button to SaveOver and Replace the Exist program.                       |
        # [#] (replace_existing_old_horizontal_machine_program_in_original_folder) Function Should locate on Class of:|
        #     [OldHorizontalScreen(Screen)].                                                                          |
        # ============================================================================================================|
        yes_button = MDRaisedButton(text='Yes',
                                    on_release=self.replace_existing_old_horizontal_machine_program_in_original_folder,
                                    font_size=16)

        # ======================================================================================================|
        # [#] Note<1,29>                                                                                        |
        # [#] Create No_Button that will call Functions of:                                                     |
        #     -(close_old_horizontal_window_of_original_folder) to close the Message_Window.                    |
        #     -(skip_create_old_horizontal_machine_program_in_original_folder) to not Create program in Original|
        #       Folder when User click No_Button.                                                               |
        # [#] Functions above Should locate on Class of: [OldHorizontalScreen(Screen)].                         |
        # ======================================================================================================|
        no_button = MDRaisedButton(text='No',
                                   on_release=self.close_old_horizontal_window_of_original_folder,
                                   on_press=self.skip_create_old_horizontal_machine_program_in_original_folder,
                                   font_size=16)

        # ==========================================================================|
        # [#] Note<1,30>                                                            |
        # [#] Create Message_Window of the Old Horizontal Screen to Inform the User.|
        # ==========================================================================|
        self.old_horizontal_screen_message_window = MDDialog(title='[color=0066ff]Confirmation Message[/color]',
                                                             text=('[b][i][u][color=0099ff]' + self.ids[
                                                                 "JobNumberForOldHorizontalMachine"].text +
                                                                 '[/color][/u][/i][/b]' +
                                                                 '[color=ffffff] Program already Exists in [/color]' +
                                                                 '[color=ffff00]ORIGINAL[/color]' +
                                                                 '[color=ffffff] Folder.[/color]' + '\n' +
                                                                 '[color=ffffff]Do you want to replace it ?[/color]'),
                                                             size_hint=(0.7, 1.0), buttons=[yes_button, no_button],
                                                             auto_dismiss=False)
        self.old_horizontal_screen_message_window.open()


        # <<Keep them until we make sure we don't need them>>
        # # MAYBE WE DO NOT NEED IT
        # # TO CLEAR THE LIST TO BE EMPTY IF ALREADY THE PROGRAM EXIST
        # success_messages_of_creating_old_horizontal_machine_program = []
        # # maybe need logic to send this message to trello in case there is warnning make this program
        # email_messages_of_creating_old_horizontal_machine_program.append("\n".join(
        #     success_messages_of_creating_old_horizontal_machine_program))


# ===============================================================================|
# [#] Create Function to create Horizontal PinBore program in the Running folder.|
# ===============================================================================|
def create_old_horizontal_machine_program_in_running_folder(self):
    print("(create_old_horizontal_machine_program_in_running_folder) Function >> called")

    # ===========================================================================================|
    # [#] Same Description of the above Function (Note<1,1>), Except it is on the Running Folder.|
    # ===========================================================================================|
    global called_need_confirmation_to_create_old_horizontal_machine_program
    called_need_confirmation_to_create_old_horizontal_machine_program = False
    print("called_need_confirmation_to_create_old_horizontal_machine_program on create Program in Runnung Folder: ",
          called_need_confirmation_to_create_old_horizontal_machine_program)

    # ===========================================================================================|
    # [#] Same Description of the above Function (Note<1,2>), Except it is on the Running Folder.|
    # ===========================================================================================|
    global running_folder_path_of_old_horizontal_machine
    print("activate_test_mode on Running Folder: ", activate_test_mode)
    if (activate_test_mode == True):
        running_folder_path_of_old_horizontal_machine = self.manager.get_screen('OldHorizontalSettingScreen').ids[
            "RunningFolderPathOfOldHorizontalMachineTestMode"].text
    elif (activate_test_mode == False):
        running_folder_path_of_old_horizontal_machine = self.manager.get_screen('OldHorizontalSettingScreen').ids[
            "RunningFolderPathOfOldHorizontalMachine"].text
    print("running_folder_path_of_old_horizontal_machine: ", running_folder_path_of_old_horizontal_machine)

    # ===========================================================================================|
    # [#] Same Description of the above Function (Note<1,3>), Except it is on the Running Folder.|
    # ===========================================================================================|
    global new_horizontal_program_for_old_machine_in_running_folder

    # ===========================================================================================|
    # [#] Same Description of the above Function (Note<1,4>), Except it is on the Running Folder.|
    # ===========================================================================================|
    global new_horizontal_program_To0_direction_for_old_machine_in_running_folder

    # ===========================================================================================|
    # [#] Same Description of the above Function (Note<1,5>), Except it is on the Running Folder.|
    # ===========================================================================================|
    global new_horizontal_program_To180_direction_for_old_machine_in_running_folder

    # ===========================================================================================|
    # [#] Same Description of the above Function (Note<1,6>), Except it is on the Running Folder.|
    # ===========================================================================================|
    global horizontal_program_lines_To0_direction_for_old_machine_in_running_folder
    horizontal_program_lines_To0_direction_for_old_machine_in_running_folder = \
        pin_bore_program_lines_of_old_horizontal_machine.copy()

    # ===========================================================================================|
    # [#] Same Description of the above Function (Note<1,7>), Except it is on the Running Folder.|
    # ===========================================================================================|
    global horizontal_program_lines_To180_direction_for_old_machine_in_running_folder
    horizontal_program_lines_To180_direction_for_old_machine_in_running_folder = \
        pin_bore_program_lines_of_old_horizontal_machine.copy()

    # ===========================================================================================|
    # [#] Same Description of the above Function (Note<1,8>), Except it is on the Running Folder.|
    # ===========================================================================================|
    try:
        print("try old_horizontal in Running Folder is called")
        # ===========================================================================================|
        # [#] Same Description of the above Function (Note<1,9>), Except it is on the Running Folder.|
        # ===========================================================================================|
        if (offset_direction == "OFFSET EACH WAY" and offset_amount != 0):

            # ============================================================================================|
            # [#] Same Description of the above Function (Note<1,10>), Except it is on the Running Folder.|
            # ============================================================================================|
            new_horizontal_program_To0_direction_for_old_machine_in_running_folder = (
                    running_folder_path_of_old_horizontal_machine + "\\" + "P" +
                    new_program_number_for_old_horizontal_machine + "TO0.MIN")

            # ============================================================================================|
            # [#] Same Description of the above Function (Note<1,11>), Except it is on the Running Folder.|
            # ============================================================================================|
            horizontal_program_lines_To0_direction_for_old_machine_in_running_folder[0] = (
                    '(PART ' + new_program_number_for_old_horizontal_machine + " TO 0" + ' -- ' + today_date
                    + ' <SYS>' + ')')

            # ============================================================================================|
            # [#] Same Description of the above Function (Note<1,12>), Except it is on the Running Folder.|
            # ============================================================================================|
            try:
                create_new_horizontal_program_To0_direction_for_old_machine_in_running_folder = open(
                    new_horizontal_program_To0_direction_for_old_machine_in_running_folder, "x")
                create_new_horizontal_program_To0_direction_for_old_machine_in_running_folder.write(
                    '\n'.join(horizontal_program_lines_To0_direction_for_old_machine_in_running_folder))
                create_new_horizontal_program_To0_direction_for_old_machine_in_running_folder.close()

            # ============================================================================================|
            # [#] Same Description of the above Function (Note<1,13>), Except it is on the Running Folder.|
            # ============================================================================================|
            except PermissionError or FileNotFoundError as error:
                # print("except LOAD_HORIZONTAL_SHEETS_FOR_AUTOMATION")
                fail_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find" + '[b][color=33cc33] RUNNING [/color][/b]' + "Folder location to Save the Program."
                    + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]' + str(
                        error) + '[/color]' + "\n" + "Double Check Network, and Folder Location." + "\n")
                email_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find Running Folder location to Save the Program." + "\n" +
                    "Double Check Network, and File Location." + "\n")
                failed_to_create_old_horizontal_machine_program(self)
                return

            # ============================================================================================|
            # [#] Same Description of the above Function (Note<1,14>), Except it is on the Running Folder.|
            # ============================================================================================|
            new_horizontal_program_To180_direction_for_old_machine_in_running_folder = (
                    running_folder_path_of_old_horizontal_machine + "\\" + "P" +
                    new_program_number_for_old_horizontal_machine + "TO180.MIN")

            # ============================================================================================|
            # [#] Same Description of the above Function (Note<1,15>), Except it is on the Running Folder.|
            # ============================================================================================|
            horizontal_program_lines_To180_direction_for_old_machine_in_running_folder[0] = (
                    '(PART ' + new_program_number_for_old_horizontal_machine + " TO 180" + ' -- ' + today_date
                    + ' <SYS>' + ')')

            # ============================================================================================|
            # [#] Same Description of the above Function (Note<1,16>), Except it is on the Running Folder.|
            # ============================================================================================|
            horizontal_program_lines_To180_direction_for_old_machine_in_running_folder[VC155_variable_index] = (
                    'VC155=' + format((-1) * offset_amount) + '  (Offset)')

            # ============================================================================================|
            # [#] Same Description of the above Function (Note<1,17>), Except it is on the Running Folder.|
            # ============================================================================================|
            if (notch_angle_first_location != 0 and notch_angle_first_location is not None):
                horizontal_program_lines_To180_direction_for_old_machine_in_running_folder[VC173_variable_index] = (
                        'VC173=' + format(float(Y_distance_from_origin_to_circlip_notch) - (2 * offset_amount), '.4f') +
                        '  (yCirclipNotch)')

            # ============================================================================================|
            # [#] Same Description of the above Function (Note<1,18>), Except it is on the Running Folder.|
            # ============================================================================================|
            try:
                create_new_horizontal_program_To180_direction_for_old_machine_in_running_folder = open(
                    new_horizontal_program_To180_direction_for_old_machine_in_running_folder, "x")
                create_new_horizontal_program_To180_direction_for_old_machine_in_running_folder.write(
                    '\n'.join(horizontal_program_lines_To180_direction_for_old_machine_in_running_folder))
                create_new_horizontal_program_To180_direction_for_old_machine_in_running_folder.close()

            # ============================================================================================|
            # [#] Same Description of the above Function (Note<1,19>), Except it is on the Running Folder.|
            # ============================================================================================|
            except PermissionError or FileNotFoundError as error:
                fail_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find" + '[b][color=33cc33] RUNNING [/color][/b]' + "Folder location to Save the Program."
                    + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]' + str(
                        error) + '[/color]' + "\n" + "Double Check Network, and Folder Location." + "\n")
                email_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find Running Folder location to Save the Program." + "\n" +
                    "Double Check Network, and File Location." + "\n")
                failed_to_create_old_horizontal_machine_program(self)
                return

        # ============================================================================================|
        # [#] Same Description of the above Function (Note<1,20>), Except it is on the Running Folder.|
        # ============================================================================================|
        else:
            # ============================================================================================|
            # [#] Same Description of the above Function (Note<1,21>), Except it is on the Running Folder.|
            # ============================================================================================|
            new_horizontal_program_for_old_machine_in_running_folder = (
                    running_folder_path_of_old_horizontal_machine + "\\" + "P" +
                    new_program_number_for_old_horizontal_machine + ".MIN")

            # ============================================================================================|
            # [#] Same Description of the above Function (Note<1,22>), Except it is on the Running Folder.|
            # ============================================================================================|
            try:
                create_new_horizontal_program_for_old_machine_in_running_folder = open(
                    new_horizontal_program_for_old_machine_in_running_folder, "x")
                create_new_horizontal_program_for_old_machine_in_running_folder.write('\n'.join(
                    pin_bore_program_lines_of_old_horizontal_machine))
                create_new_horizontal_program_for_old_machine_in_running_folder.close()

            # ============================================================================================|
            # [#] Same Description of the above Function (Note<1,23>), Except it is on the Running Folder.|
            # ============================================================================================|
            except PermissionError or FileNotFoundError as error:
                fail_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find" + '[b][color=33cc33] RUNNING [/color][/b]' + "Folder location to Save the Program."
                    + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]' + str(
                        error) + '[/color]' + "\n" + "Double Check Network, and Folder Location." + "\n")
                email_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find Running Folder location to Save the Program." + "\n" +
                    "Double Check Network, and File Location." + "\n")
                failed_to_create_old_horizontal_machine_program(self)
                return

        # ============================================================================================|
        # [#] Same Description of the above Function (Note<1,24>), Except it is on the Running Folder.|
        # ============================================================================================|
        success_messages_of_creating_old_horizontal_machine_program = \
            ["\n" + "Program has been **CREATED** successfully in **RUNNING** Folder."]
        email_messages_of_creating_old_horizontal_machine_program.append("\n".join(
            success_messages_of_creating_old_horizontal_machine_program))

        # ============================================================================================|
        # [#] Same Description of the above Function (Note<1,25>), Except it is on the Running Folder.|
        # ============================================================================================|
        close_button = MDRaisedButton(text='Close', on_release=self.close_old_horizontal_window_of_running_folder,
                                      font_size=16)

        # ============================================================================================|
        # [#] Same Description of the above Function (Note<1,26>), Except it is on the Running Folder.|
        # ============================================================================================|
        self.old_horizontal_screen_message_window = MDDialog(title='[color=248f24]Success Message[/color]', text=(
                '[b][i][u][color=0099ff]' + self.ids["JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                '[color=ffffff] Program Has been Created Successfully in [/color]' +
                '[color=33cc33]RUNNING[/color]' + '[color=ffffff] Folder.[/color]' + '\n' +
                '[color=ffffff]After closing this window, the program will open on CIMCO Editor.[/color]' + '\n' +
                '[color=ffffff]Please Double Check and Report any Problem on Trello.[/color]'),
                                                             size_hint=(0.7, 1.0), buttons=[close_button],
                                                             auto_dismiss=False)
        self.old_horizontal_screen_message_window.open()
        print()

    # ============================================================================================|
    # [#] Same Description of the above Function (Note<1,27>), Except it is on the Running Folder.|
    # ============================================================================================|
    except(FileExistsError):

        # ============================================================================================|
        # [#] Same Description of the above Function (Note<1,28>), Except it is on the Running Folder.|
        # ============================================================================================|
        yes_button = MDRaisedButton(text='Yes',
                                    on_release=self.replace_existing_old_horizontal_machine_program_in_running_folder,
                                    font_size=16)

        # ============================================================================================|
        # [#] Same Description of the above Function (Note<1,29>), Except it is on the Running Folder.|
        # ============================================================================================|
        no_button = MDRaisedButton(text='No', on_release=self.close_old_horizontal_screen_window, font_size=16)

        # ============================================================================================|
        # [#] Same Description of the above Function (Note<1,30>), Except it is on the Running Folder.|
        # ============================================================================================|
        self.old_horizontal_screen_message_window = MDDialog(title='[color=0066ff]Confirmation Message[/color]',
                                                             text=('[b][i][u][color=0099ff]' + self.ids[
                                                                 "JobNumberForOldHorizontalMachine"].text +
                                                                 '[/color][/u][/i][/b]' +
                                                                 '[color=ffffff] Program already Exists in [/color]' +
                                                                 '[color=33cc33]Running[/color]' +
                                                                 '[color=ffffff] Folder.[/color]' + '\n' +
                                                                 '[color=ffffff]Do you want to replace it ?[/color]'),
                                                             size_hint=(0.7, 1.0), buttons=[yes_button, no_button],
                                                             auto_dismiss=False)
        self.old_horizontal_screen_message_window.open()


        # <<Keep them until we make sure we don't need them>>
        # # MAYBE WE DO NOT NEED IT
        # # TO CLEAR THE LIST TO BE EMPTY IF ALREADY THE PROGRAM IS EXIST
        # success_messages_of_creating_old_horizontal_machine_program = []
        # email_messages_of_creating_old_horizontal_machine_program.append("\n".join(
        #     success_messages_of_creating_old_horizontal_machine_program))


# ===============================================================================================================|
# [#] Create Function to Set the confirmation details that's need to finalize by User to create the program.     |
# [#] (title, sub_function, dialog_type, content): are all parameters will set according to confirmation details.|
# ===============================================================================================================|
def need_confirmation_to_create_old_horizontal_machine_program(self, title, sub_function, dialog_type, content):
    print("(need_confirmation_to_create_old_horizontal_machine_program) Function >> called")
    print('\n'.join(verification_messages_of_creating_old_horizontal_machine_program))

    # =================================================================================================|
    # [#] Define Variable to set Status of needing confirmation to create the program.                 |
    # [#] Set the value to be "False" by default assuming no confirmation needed to create programs.   |
    # [#] Change the value to be "True" when find confirmation details that's need to finalize by User.|
    # =================================================================================================|
    global called_need_confirmation_to_create_old_horizontal_machine_program
    called_need_confirmation_to_create_old_horizontal_machine_program = True
    print("called_need_confirmation_to_create_old_horizontal_machine_program on <Need Confirmation> Function: ",
          called_need_confirmation_to_create_old_horizontal_machine_program)

    # ============================================================================|
    # [#] Create Enter_Button that will call Functions of:                        |
    #     -(sub_function) will be set according to confirmation details.          |
    #     -(create_program_for_old_horizontal_machine) to Create the program.     |
    # [#] Function above Should locate on Class of: [OldHorizontalScreen(Screen)].|
    # ============================================================================|
    enter_button = MDRaisedButton(text='Enter', on_press=sub_function,
                                  on_release=self.create_program_for_old_horizontal_machine,
                                  font_size=16)

    # =============================================================================================|
    # [#] Create CloseButton that will call (close_old_horizontal_screen_window) Function when User|
    #     click on the Button to close the Window and do some actions.                             |
    # [#] Function above Should locate on Class of: [OldHorizontalScreen(Screen)].                 |
    # =============================================================================================|
    close_button = MDRaisedButton(text='Close', on_release=self.close_old_horizontal_screen_window,
                                  font_size=16)

    # =============================================================================================|
    # [#] Types of Confirmation details that's need to create the program :                        |
    #     -(custom): it needs the user to enter value to finalize the confirmation details.        |
    #     -(confirmation): it needs the user to choose Option to finalize the confirmation details.|
    # =============================================================================================|

    # ==============================================================================================|
    # [#] custom: it needs the user to enter value to finalize the confirmation details.            |
    # [#] Create Dialog with All Parameters come from the confirmation details that needed to inform|
    #     the user to finalize the confirmation details.                                            |
    # [#] (content_cls): it will contain the content of the confirmation details.                   |
    # ==============================================================================================|
    if (dialog_type == "custom"):
        self.old_horizontal_screen_message_window = MDDialog(
            title=title, type=dialog_type, content_cls=content, size_hint=(0.7, 1.0),
            buttons=[enter_button, close_button], auto_dismiss=False)

    # ==============================================================================================|
    # [#] confirmation: it needs the user to choose Option to finalize the confirmation details.    |
    # [#] Create Dialog with All Parameters come from the confirmation details that needed to inform|
    #     the user to finalize the confirmation details.                                            |
    # [#] (items): it will contain the Options of the confirmation details.                         |
    # ==============================================================================================|
    elif (dialog_type == "confirmation"):
        self.old_horizontal_screen_message_window = MDDialog(
            title=title, type=dialog_type, items=content, size_hint=(0.7, 1.0),
            buttons=[enter_button, close_button], auto_dismiss=False)

    # =========================================|
    # [#] To Open the Message Window of Screen.|
    # =========================================|
    self.old_horizontal_screen_message_window.open()


# =========================================================================|
# [#] Create Function to Set the Warning details that's User needs to know.|
# =========================================================================|
def old_horizontal_machine_program_needs_attention(self):
    print("(old_horizontal_machine_program_needs_attention) Function >> called")

    # =============================================================================================|
    # [#] Create CloseButton that will call (close_old_horizontal_screen_window) Function when User|
    #     click on the Button to close the Window and do some actions.                             |
    # [#] Function above Should locate on Class of: [OldHorizontalScreen(Screen)].                 |
    # =============================================================================================|
    close_button = MDRaisedButton(text='Close', on_release=self.close_old_horizontal_screen_window, font_size=16)

    # ==========================================================================|
    # [#] Create Message_Window of the Old Horizontal Screen to Inform the User.|
    # ==========================================================================|
    self.old_horizontal_screen_message_window = MDDialog(title='[color=cccc00]Warning Message[/color]', text=(
            '[color=ffffff]' + '\n'.join(warning_messages_of_creating_old_horizontal_machine_program) + '[/color]'),
                                                         size_hint=(0.7, 1.0),
                                                         buttons=[close_button], auto_dismiss=False)
    self.old_horizontal_screen_message_window.open()


# =============================================================================================|
# [#] Create Function to Set the failure details of Creating Program that's User needs to know.|
# =============================================================================================|
def failed_to_create_old_horizontal_machine_program(self):
    print("(failed_to_create_old_horizontal_machine_program) Function >> called")

    # =============================================================================================|
    # [#] Create CloseButton that will call (close_old_horizontal_screen_window) Function when User|
    #     click on the Button to close the Window and do some actions.                             |
    # [#] Function above Should locate on Class of: [OldHorizontalScreen(Screen)].                 |
    # =============================================================================================|
    close_button = MDRaisedButton(text='Close', on_release=self.close_old_horizontal_screen_window, font_size=16)

    # ==========================================================================|
    # [#] Create Message_Window of the Old Horizontal Screen to Inform the User.|
    # ==========================================================================|
    self.old_horizontal_screen_message_window = MDDialog(title='[color=990000]Warning Message[/color]', text=(
            '[color=ffffff]' + '\n'.join(fail_messages_of_creating_old_horizontal_machine_program) + '[/color]'),
                                                         size_hint=(0.7, 1.0),
                                                         buttons=[close_button], auto_dismiss=False)
    self.old_horizontal_screen_message_window.open()


# =================================================================================|
# [#] Create Function to Set All details of Creating Program and Send them to User.|
# =================================================================================|
def send_email_about_create_old_horizontal_machine_program(self):
    print("(send_email_about_create_old_horizontal_machine_program) Function >> called")

    # ==============================================================================================================|
    # [#] Use (try/except) Blocks to Handle any Error may occur when connecting the Server or trying Send the Email.|
    # ==============================================================================================================|
    try:
        # ===================================================================|
        # [#] Use (smtplib) Library to Send and Manege Emails.               |
        # [#] Information needed to Set connection using a Gmail account:    |
        #     -Application Name Used to Send Emails: "outlook.office365.com".|
        #     -Server Name: "smtp.office365.com".                            |
        #     -Port Number (that needs to be able to connect the server): 587|
        # ===================================================================|
        service_app = "outlook.office365.com"
        smtp_server = "smtp.office365.com"
        port = 587

        # ======================================================================================================|
        # [#] Still Needs work after we figure out Email thing with JJ from IT department                       |
        # [#] Use Email of 'CurrentUser' to send email to the Trello Board, or Use ONE Email to send All Emails.|
        # ======================================================================================================|

        # ======================================================================================================|
        # [#] In case we need to use 'CurrentUser' option:                                                      |
        # ======================================================================================================|
        # ======================================================================================================|
        # [#] Use Email of <Current User> to send email to the Trello Board.                                    |
        # ======================================================================================================|
        sender_email = user_email_address

        # =============================================================================================================|
        # We MAKE THIS OPTION TO ALLOW USER TO ENTER PASSWORD IF WE HAVE AUTHENTICATION PROBLEM TO MAKE SURE>>>        |
        # >>>THE ISSUE IS NOT THE PASSWORD                                                                             |
        # MAKE 'IF STATEMENT' TO SET EMAIL PASSWORD TO BE WHAT USER ENTER IN <USER SETTING SCREEN>, OTHERWISE>>>>>     |
        # >>USING KEYRING PACKAGE TO GET PASSWORD FROM Windows Credential Manager WHILE THEY ARE SAVING IN USER COMPUTER
        # =============================================================================================================|
        sender_password_input = self.manager.get_screen('UserSettingScreen').ids["EmailPassword"].text
        if(sender_password_input != ""):
            sender_password = sender_password_input
        else:
            sender_password = keyring.get_password(service_app, sender_email)
        print("sender_email: ", sender_email)

        # ===============================================================================================|
        # [#] Set Trello Board Address as a Receiver Email                                               |
        # [#] To set the Trello Board Address:                                                           |
        #     Access MDTextField of (id: TrelloEmailAddress) in (AppSettingScreen) from  Screens_Builder.|
        # ===============================================================================================|
        trello_board_email = self.manager.get_screen('AppSettingScreen').ids["TrelloEmailAddress"].text

        # ================================================|
        # Creates SMTP session to connect with the server.|
        # ================================================|
        email_server = smtplib.SMTP(smtp_server, port)

        # ==============================================================================|
        # Use <starttls()> method call to connect using the TLS encryption for security.|
        # ==============================================================================|
        email_server.starttls()

        # ============================================================================================================|
        # [#] Login to Sender Email.                                                                                  |
        # [#] comment it out to not send any thing for now until we figure out Email thing with JJ from IT department.|
        # ============================================================================================================|
        # email_server.login(sender_email, sender_password)

        # ==================================================|
        # Set Email Contents that will send to Trello Board.|
        # ==================================================|
        if (fail_messages_of_creating_old_horizontal_machine_program != []):
            card_title = "Failed to Create Program for " + new_program_number_for_old_horizontal_machine
            card_label = "#FAILED"
            card_member = "@moemenalatweh1 " + "@" + trello_user_name
        elif (warning_messages_of_creating_old_horizontal_machine_program != []):
            card_title = new_program_number_for_old_horizontal_machine + " Program Needs Attention."
            card_label = "#WARNING"
            card_member = "@moemenalatweh1 " + "@" + trello_user_name
        else:
            card_title = new_program_number_for_old_horizontal_machine
            card_label = "#SUCCESS"
            card_member = "@moemenalatweh1 " + "@" + trello_user_name

        # ====================================================|
        # Create Email Message that will send to Trello Board.|
        # ====================================================|
        email_message = f"""From: Alatweh Moemen <malatweh@rwbteam.com>
Subject: {card_title}  {card_label}  {card_member}


{"".join(email_messages_of_creating_old_horizontal_machine_program)}"""

        # ============================================================================================================|
        # [#] Send the Email to Trello Board.                                                                         |
        # [#] comment it out to not send any thing for now until we figure out Email thing with JJ from IT department.|
        # ============================================================================================================|
        # email_server.sendmail(sender_email, trello_board_email, email_message)

        # ===================================================|
        # End the SMTP session to disconnect with the server.|
        # ===================================================|
        email_server.quit()

    # =========================================================================================================|
    # [#] Use (except) Block to Handle any Error may occur when connecting the Server or trying Send the Email.|
    # [#] Error could happen like: Wrong Login Info, Enable Server Connection...Etc.                           |
    # [#] Use (return) to stop the function (APP as a result) from executing the rest of the code.             |
    # [#] Call (failed_to_create_old_horizontal_machine_program) Function to warn the user.                    |
    # =========================================================================================================|
    except Exception as error:
        fail_messages_of_creating_old_horizontal_machine_program.append(
            "Failed to send Email to Trello board." + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]' + str(
                error) + '[/color]' + "\n" + "Double Check Network and Login Authentication." + "\n")
        failed_to_create_old_horizontal_machine_program(self)
        return


# endregion <<<<=========================[Old Horizontal Machines(28,29,32) Functions]===========================>>>>


# region <<<<=============================[Old Horizontal Machines(28,29,32) Screen]=============================>>>>

# ==========================================|
# [#] Create Class of (OldHorizontalScreen).|
# ==========================================|
class OldHorizontalScreen(Screen):

    # ======================================================================================================|
    # [#] Create Function that will be activated once enter the 'OldHorizontalScreen' to Set some Variables.|
    # [#] Needs to add it to 'OldHorizontalScreen' on Screens_Builder as well.                              |
    # ======================================================================================================|
    def on_pre_enter(self):

        # =================================================================================================|
        # [#] Define Variable to set Status of needing confirmation to create the program.                 |
        # [#] Set the value to be "False" by default assuming no confirmation needed to create programs.   |
        # [#] Change the value to be "True" when find confirmation details that's need to finalize by User.|
        # =================================================================================================|
        global called_need_confirmation_to_create_old_horizontal_machine_program
        called_need_confirmation_to_create_old_horizontal_machine_program = False
        print("called_need_confirmation_to_create_old_horizontal_machine_program <on_pre_enter>: ",
              called_need_confirmation_to_create_old_horizontal_machine_program)

        # ========================================================================|
        # [#] By Default Deactivate Test Mode by putting ('False).                |
        # [#] Needs to add it to 'OldHorizontalScreen' on Screens_Builder as well.|
        # ========================================================================|
        global activate_test_mode
        activate_test_mode = False
        print("activate_test_mode on_pre_enter: ", activate_test_mode)

    # ===================================================================================|
    # [#] Create Function to Set the Live Mode.                                          |
    # [#] Live Mode: is the Mode that save the Created Programs on the actual Folder that|
    #     connected with the Machines.                                                   |
    # ===================================================================================|
    def set_live_mode(self, obj):
        print("(set_live_mode) Function >> called")

        # ===============================================================================================|
        # Set COLOR of 'text' and 'borders' of LiveModeButton to be GREEN to indicate is the Active Mode.|
        # ===============================================================================================|
        self.ids["LiveModeButton"].text_color = 0, 1, 0, 1
        self.ids["LiveModeButton"].line_color = 0, 1, 0, 1

        # ===================================================================================================|
        # Set COLOR of 'text' and 'borders' of TestModeButton to be WHITE to indicate is the Deactivate Mode.|
        # ===================================================================================================|
        self.ids["TestModeButton"].text_color = 1, 1, 1, 1
        self.ids["TestModeButton"].line_color = 1, 1, 1, 1

        # =================================================|
        # [#] Deactivate the Test Mode by putting ('False).|
        # =================================================|
        global activate_test_mode
        activate_test_mode = False
        print("activate_test_mode on_live_mode: ", activate_test_mode)

    # ==================================================================================|
    # [#] Create Function to Set the Test Mode.                                         |
    # [#] Test Mode: is the Mode that save the Created Programs on different Folder that|
    #     NOT connected to the Machines.                                                |
    # ==================================================================================|
    def set_test_mode(self, obj):
        print("(set_test_mode) Function >> called")

        # ===============================================================================================|
        # Set COLOR of 'text' and 'borders' of TestModeButton to be GREEN to indicate is the Active Mode.|
        # ===============================================================================================|
        self.ids["TestModeButton"].text_color = 0, 1, 0, 1
        self.ids["TestModeButton"].line_color = 0, 1, 0, 1

        # ===================================================================================================|
        # Set COLOR of 'text' and 'borders' of LiveModeButton to be WHITE to indicate is the Deactivate Mode.|
        # ===================================================================================================|
        self.ids["LiveModeButton"].text_color = 1, 1, 1, 1
        self.ids["LiveModeButton"].line_color = 1, 1, 1, 1

        # ==============================================|
        # [#] Activate the Test Mode by putting ('True).|
        # ==============================================|
        global activate_test_mode
        activate_test_mode = True
        print("activate_test_mode on_Test_mode: ", activate_test_mode)

    # ===========================================================================|
    # [#] Create Function to Create PinBore Programs for Old Horizontal Machines.|
    # ===========================================================================|
    def create_program_for_old_horizontal_machine(self, obj):
        print("(create_program_for_old_horizontal_machine) Function >> called")

        # =====================================================================================|
        # [#] APP allow guest users to see the APP without any action, because of that needs to|
        #     use (if Statement) to not allow any guest user to make programs.                 |
        # =====================================================================================|
        if (self.manager.get_screen('LoginScreen').ids["Email"].text in guests_email_address_list):
            close_button = MDRaisedButton(
                text='Close', on_release=self.close_old_horizontal_screen_window, font_size=16)

            self.old_horizontal_screen_message_window = MDDialog(title='[color=990000]Warning Message[/color]', text=(
             '[color=ffffff]Sorry, You are NOT authorized to make CNC Programs,' + "\n" +
             'THANKS for your understanding.' + '[/color]'),
                                size_hint=(0.7, 1.0), buttons=[close_button], auto_dismiss=False)
            self.old_horizontal_screen_message_window.open()
            return

        # ========================================================================================|
        # [#] Define List to contain ALL Success Messages of creating programs to inform the user.|
        # ========================================================================================|
        global success_messages_of_creating_old_horizontal_machine_program
        success_messages_of_creating_old_horizontal_machine_program = []

        # =============================================================================================|
        # [#] Define List to contain ALL Verification Messages of creating programs to inform the user.|
        # =============================================================================================|
        global verification_messages_of_creating_old_horizontal_machine_program
        verification_messages_of_creating_old_horizontal_machine_program = []

        # =====================================================================================|
        # [#] Define List to contain ALL Fail Messages of creating programs to inform the user.|
        # =====================================================================================|
        global fail_messages_of_creating_old_horizontal_machine_program
        fail_messages_of_creating_old_horizontal_machine_program = []

        # ========================================================================================|
        # [#] Define List to contain ALL Warning Messages of creating programs to inform the user.|
        # ========================================================================================|
        global warning_messages_of_creating_old_horizontal_machine_program
        warning_messages_of_creating_old_horizontal_machine_program = []

        # =========================================================================================================|
        # [#] Define List to contain ALL Notes that needed to be added to the top of the created Program.          |
        # [#] Set the list to have Empty text to avoid break the code if list not used (index[0] can't be nothing).|
        # =========================================================================================================|
        global notes_for_old_horizontal_machine_program
        notes_for_old_horizontal_machine_program = ['']

        # ======================================================================================================|
        # [#] Define Variable to set index of (notes_for_old_horizontal_machine_program) list to be <0> to start|
        #     add them (if any) to the top of the created Program.                                              |
        # ======================================================================================================|
        global notes_for_old_horizontal_machine_program_index
        notes_for_old_horizontal_machine_program_index = 0

        # ===================================================================================================|
        # [#] Define Variable to set Status of Creating program on the Original Folder when User click on    |
        #     No_Button on (Replace Program Question).                                                       |
        # [#] Set Status to be 'False' by default to try always create program on original folder unless user|
        #     click on No_Button on (Replace Program Question).                                              |
        # ===================================================================================================|
        global dont_create_old_horizontal_machine_program_in_original_folder
        dont_create_old_horizontal_machine_program_in_original_folder = False

        # ======================================================================================|
        # [#] Define List to contain ALL Email Messages of creating programs to inform the user.|
        # [#] Set the Default message on the top of the email to contain:                       |
        #     Job Number + "Program on Old Horizontal Machine" + "Created by : " + User Name.   |
        # ======================================================================================|
        global email_messages_of_creating_old_horizontal_machine_program
        email_messages_of_creating_old_horizontal_machine_program = []
        email_messages_of_creating_old_horizontal_machine_program.append(
            self.ids["JobNumberForOldHorizontalMachine"].text + " Program on " + "Old Horizontal Machine" + "\n" +
            "Created by : " + connected_user_name + "\n" + "\n")

        # ==============================================================================================|
        # [#] Define Variable to set Job Number of the Program.                                         |
        # [#] Access MDTextField of (id: JobNumberForOldHorizontalMachine) in (OldHorizontalScreen) from|
        #     Screens_Builder to get Job Number of the Program.                                         |
        # ==============================================================================================|
        global new_program_number_for_old_horizontal_machine
        new_program_number_for_old_horizontal_machine = self.ids["JobNumberForOldHorizontalMachine"].text
        print("Job Number: ", new_program_number_for_old_horizontal_machine)

        # =============================================================================|
        # [#] Define Variable to set ID of Job (Piston).                               |
        # [#] piston_id: is the unique ID for each Job (Piston) stored in the DataBase.|
        # =============================================================================|
        global piston_id

        # ==========================================================================|
        # [#] Define Variable to set ID of Forging.                                 |
        # [#] ForgeSpecID: is the unique ID for each Forging stored in the DataBase.|
        # ==========================================================================|
        global forge_spec_id

        # ======================================================================================================|
        # [#] Use (if Statement) to Set 'piston_id' by using the Job Number.                                    |
        #    [#] Check if the user enter NOTHING in the TextInput of Job Number to set 'piston_id' to           |
        #        equal <1> to avoid Error of Accessing the DaraBase without PISTON ID.                          |
        #    [#] Check if the user enter any thing like "TEST" in the TextInput of Job Number to set 'piston_id'|
        #        to equal <1> to avoid Error of creating program for <Test> Jobs.                               |
        #        -Needs to do that because user may try to enter 'Test' to try the APP and Some Jobs stored     |
        #         as 'TEST' in the DataBase.                                                                    |
        #    [#] In Cases above Set 'piston_id' to equal <1> and Set 'forge_spec_id' to equal 'None' to         |
        #        avoid an error when use them with the DataBase.                                                |
        #    [#] Use (return) to stop the function (APP as a result) from executing the rest of the code.       |
        #    [#] Call (failed_to_create_old_horizontal_machine_program) Function to warn the user.              |
        # ======================================================================================================|
        if (self.ids["JobNumberForOldHorizontalMachine"].text == "" or
                self.ids["JobNumberForOldHorizontalMachine"].text == " "):
            fail_messages_of_creating_old_horizontal_machine_program.append("Please Enter The Job Number." + "\n")
            piston_id = 1
            forge_spec_id = None
            failed_to_create_old_horizontal_machine_program(self)
            return
        elif ((self.ids["JobNumberForOldHorizontalMachine"].text == "TEST") or
              (self.ids["JobNumberForOldHorizontalMachine"].text == "TEST1") or
              (self.ids["JobNumberForOldHorizontalMachine"].text == "TEST2") or
              (self.ids["JobNumberForOldHorizontalMachine"].text == "TEST3")):
            fail_messages_of_creating_old_horizontal_machine_program.append(
                "Job Number can't be like that, Please Enter Valid Number." + "\n")
            piston_id = 1
            forge_spec_id = None
            failed_to_create_old_horizontal_machine_program(self)
            return

        # =============================================================================================|
        #    [#] If the User Enter Job Number and was not something like 'Test', Access the DataBase to|
        #        try to Set the 'piston_id'.                                                           |
        # =============================================================================================|
        else:
            # ======================================================================================================|
            # [#] Get the 'piston_id' from the DataBase by using 'Job Number' and Access the Data location by using:|
            #     Table Name : SpexPiston                                                                           |
            #     Column Name : PistonID                                                                            |
            # [#] Use (try/except) Blocks to Handle any Error of not finding the 'piston_id' in the DataBase.       |
            # ======================================================================================================|
            try:
                engine_worx_database_cursor.execute(
                    "SELECT PistonID FROM SpexPiston WHERE Piston = ?", new_program_number_for_old_horizontal_machine)

                # ================================================================================================|
                # [#] Start Setting 'piston_id' to equal <1> and 'forge_spec_id' to equal 'None' in case the user |
                #     enter NOT valid job number or it was not released or exist in the DataBase to avoid an error|
                #     when use them with the DataBase.                                                            |
                # ================================================================================================|
                piston_id = 1
                forge_spec_id = None

                # ================================================================|
                # [#] Use for loop to iterate through the DateBase.               |
                # [#] Set the value as it is stored in the DataBase.              |
                # [#] Expected Data output Type : 'Numeric'                       |
                # [#] Expected Data output value : Numeric Value  <1> to <+55000>.|
                # ================================================================|
                for data in engine_worx_database_cursor.fetchone():
                    piston_id = data
                print("[#]PistonID FOR " + new_program_number_for_old_horizontal_machine + ":")
                print("     ", piston_id)

            # ============================================================================================|
            # [#] Use (except) Blocks to Handle any Error of not finding the 'piston_id' in the DataBase. |
            # [#] Call (failed_to_create_old_horizontal_machine_program) Function to warn the user.       |
            # [#] Use (return) to stop the function (APP as a result) from executing the rest of the code.|
            # ============================================================================================|
            except Exception as error:
                fail_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Use Job Number " + '[b][u][color=0099ff]' + self.ids[
                     "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/b]' +
                    " to Create the Program." + "\n" + "An Error has occurred : " + "\n" + '[color=ff1a1a]' + str(error)
                    + '[/color]' + "\n" + "Make Sure the Job Number you entered is correct and it has been "
                    "created in EW." + "\n" + "Otherwise, Check Network, and Connection." + "\n")
                failed_to_create_old_horizontal_machine_program(self)
                return

        # =============================================================================================================|
        # [#] Call (four_cycle_pin_bore_variables) Function to Set Job Variables (information) from the DataBase.      |
        # [#] When Certain Job has confirmation details that needed to create the program, the APP will Stop           |
        #    Executing the Code until the user finalize the confirmation details (by Entering value or choosing option)|
        #    and Click Button to try to create the program again which will call this function                         |
        #    (create_program_for_old_horizontal_machine), at this moment the code is using the data from DataBase plus |
        #    the data that user just finalized and changed (the confirmation details), therefor it needs to check      |
        #    the confirmation details STATUS to avoid set all the data again and keep repeating over and over again.   |
        #    - When Status is 'False', it means the job has NO confirmation details and call the function.             |
        #    - When Status is changed to 'True', it means the user finalize the confirmation details and no need       |
        #      to call the Function again.                                                                             |
        # =============================================================================================================|
        if (called_need_confirmation_to_create_old_horizontal_machine_program == False):
            four_cycle_pin_bore_variables(self)

        # ==========================================================================================|
        # [#] Stop the code here when there is something cause failed messages to avoid Code Errors.|
        # ==========================================================================================|
        if (fail_messages_of_creating_old_horizontal_machine_program != []):
            print(">>Stop The Code for FAILED MESSAGES in (create_program_for_old_horizontal_machine) Function")
            return

        # =================================================================================================|
        # [#] Because the Old Horizontal Machines not setup to run 2-Stroke Jobs, needs to Check Forging   |
        #     Type (from Forging DataBase) if it's "2CYCLE" to stop Executing the Code and inform the user.|
        # [#] Call (failed_to_create_old_horizontal_machine_program) Function to warn the user.            |
        # [#] Use (return) to stop the function (APP as a result) from executing the rest of the code.     |
        # =================================================================================================|
        if (forging_number != "" and forging_number is not None):
            if(forging_type == "2CYCLE"):
                fail_messages_of_creating_old_horizontal_machine_program.append(
                    "Machine doesn't setup to run" + '[b][u][color=ffffff] 2 Stroke job [/color][/u][/b]' +
                    " (2Cycle Forging: " + forging_number + ")," + '\n' +
                    "Try to make the program on different machines." + "\n")
                email_messages_of_creating_old_horizontal_machine_program.append(
                    "Machine doesn't setup to run 2 Stroke job (2Cycle Forging: " + forging_number + ")," +
                    '\n' + "Try to make the program on different machines." + "\n")
                failed_to_create_old_horizontal_machine_program(self)
                return

        # ================================================================================================|
        # [#] Because the Old Horizontal Machines not setup to run 2-Stroke Jobs, needs to Check Engine   |
        #     Type (from Piston DataBase) if it's "2CYCLE" to stop Executing the Code and inform the user.|
        # [#] Call (failed_to_create_old_horizontal_machine_program) Function to warn the user.           |
        # [#] Use (return) to stop the function (APP as a result) from executing the rest of the code.    |
        # ================================================================================================|
        if (forge_spec_id is not None):
            if (engine_stroke_type == "2 Cycle"):
                fail_messages_of_creating_old_horizontal_machine_program.append(
                    "Machine doesn't setup to run" + '[b][u][color=ffffff] 2 Stroke job [/color][/u][/b]' +
                    " (Engine Stroke Type: " + engine_stroke_type + ")," +
                    '\n' + "Try to make the program on different machines." + "\n")
                email_messages_of_creating_old_horizontal_machine_program.append(
                    "Machine doesn't setup to run 2 Stroke job (Engine Stroke Type: " + engine_stroke_type + ")," +
                    '\n' + "Try to make the program on different machines." + "\n")
                failed_to_create_old_horizontal_machine_program(self)
                return

        # ======================================================================================================|
        # [#] Call (load_horizontal_machine_tool_list_sheets) Function to Set and Load all Tool information     |
        #     from the Tool List from the Excel Sheet.                                                          |
        # [#] Use (try/except) Blocks to Handle any Error may occur when Accessing or Finding Files and Folders.|
        # [#] Call (failed_to_create_old_horizontal_machine_program) Function to warn the user.                 |
        # [#] Use (return) to stop the function (APP as a result) from executing the rest of the code.          |
        # ======================================================================================================|
        try:
            load_horizontal_machine_tool_list_sheets(self)
        except Exception as error:
            fail_messages_of_creating_old_horizontal_machine_program.append(
                "Failed to Find, Load, or Access" + '[b][u][color=ffffff] Horizontal Tool List [/color][/u][/b]' +
                "File to Create the Program." + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]' + str(
                    error) + '[/color]' + "\n" + "Double Check Network, and File Location." + "\n")
            email_messages_of_creating_old_horizontal_machine_program.append(
                "Failed to Find, Load, or Access Horizontal Tool List File to Create the Program." + "\n" +
                "Double Check Network, and File Location." + "\n")
            failed_to_create_old_horizontal_machine_program(self)
            return

        # region <<<<=====================================[PROBE PROGRAMS]=====================================>>>>

        # ==================================================================================|
        # [#] Define Variable to set Probe Programs Folder path.                            |
        # [#] Access MDTextField of (id: ProbePrograms) in (OldHorizontalSettingScreen) from|
        #     Screens_Builder to get Probe Programs Folder Path.                            |
        # ==================================================================================|
        probe_programs_folder_path_of_old_horizontal_machine = self.manager.get_screen(
            'OldHorizontalSettingScreen').ids["ProbePrograms"].text

        # ============================================================================================================|
        # [#] Forging Numbers in the DataBase are stored with the Rev extension like                                  |
        #    (F6566XA0,F6064MZA1,F4027TDXA2,FJE160-HEX...Etc),On other hand the ProbPrograms saved in variant ways,   |
        #    Sometimes it saved with the whole thing(with one of the REV) like 'F6228XA4', sometimes without the 'Rev'|
        #    like 'F6444X', because of that needs to 'Filter' the Forging Number to not have the Rev extension to make|
        #    ProbePrograms searching process more efficient.                                                          |
        # [#] Define Variable of 'Filtered' Forging Number that will use to Search the ProbePrograms and              |
        #     Set it to be "F" (because all forging numbers start with that).                                         |
        # [#] Define Variable of 'digit' and set it to be <0> to use it to iterate through the whole Forging Number.  |
        # [#] Use (while and for) loops with (if statement) to Set the 'Filtered' Forging Number by adding            |
        #     each digit of Forging number until find 'X' or 'Z' Letter to add it as well and End the loop.           |
        # ============================================================================================================|
        if forging_number != "" and forging_number is not None:
            forging_number_for_probe_program = "F"
            digit = 0
            # ==============================================================================================|
            # [#] Set the Condition of the (while) loop to End the loop when find 'X' or 'Z' Letter or reach|
            #     to the last digit of Forging Number (by putting 'forging_number[-1]').                    |
            # ==============================================================================================|
            while (digit != "X" and digit != "x" and digit != "Z" and digit != "z" and digit != forging_number[-1]):

                # ==============================================================================================|
                # [#] Start iterate through the Forging Number From index (position) <1> not <0> (by putting    |
                #     forging_number[1:]) because no need to add the 'F' letter again (it set on                |
                #     'forging_number_for_probe_program' above).                                                |
                # [#] Add each digit to the 'Filtered' Forging Number until find 'X' or 'Z' Letter to add       |
                #     it as well then End (break) the loop.                                                     |
                # [#] As a Result, 'Filtered' Forging Number will be forging number without Rev extension like: |
                #     F6566X, F6064MZ, F4027TDX, FJE160-HEX...Etc.                                              |
                # [#] Some ProbPrograms stored for "X" not "Z" or vice versa, for these forging the User needs  |
                #     to store the ProbProgram in both letters ('X' and 'Z') to be always able to find the      |
                #     program while it's created.                                                               |
                # ==============================================================================================|
                for digit in forging_number[1:]:
                    forging_number_for_probe_program = forging_number_for_probe_program + digit
                    if (digit == "X" or digit == "x" or digit == "Z" or digit == "z"):
                        break

                    # <<Keep them until we make sure we don't need them>>
                    # forging_number_for_probe_program = forging_number_for_probe_program + digit
                        # just for now  ,
                    # ALSO F4016, F4027, F4043, F4076, F4366, F4592, F4606, F4693, F4735, F4752, F4838, F4843, F4977,
                    # F5516, F6028, F6035, F6048, F6064, F6068, F6149, F6160, F6301, F6370, F6419, FJE426B.
                    # BESIDE WHAT WE HAVE BELOW: F6052, F6056, F4678, F6167
                    # if (forging_number_for_probe_program == "F6052" or forging_number_for_probe_program == "F6056" or
                    #     forging_number_for_probe_program == "F4678" or forging_number_for_probe_program == "F6167"):
                    #     forging_number_for_probe_program = forging_number_for_probe_program + "X"
            print(forging_number_for_probe_program)

            # ========================================================================================================|
            # [#] Define list to add all ProbPrograms that found from search of the 'Filtered' Forging Number.        |
            # [#] Use (for) loop with (glob.glob) method to search in (probe_programs_folder) path.                   |
            # [#] Use (*) after Forging Number to search the file even with part of name to find all possible results.|
            #     (Ex: name of F4035X it gives result of F4035XA0, F8835X it gives result of F8835XA1...Etc).         |
            # [#] Use (try/except) Blocks to Handle any Error may occur when Accessing or Finding Files and Folders.  |
            # ========================================================================================================|
            try:
                result_of_probe_program_search_for_old_horizontal_machine = []
                for file in glob.glob(probe_programs_folder_path_of_old_horizontal_machine + '*\*' +
                                      forging_number_for_probe_program + '*'):
                    result_of_probe_program_search_for_old_horizontal_machine.append(file)
                print(len(result_of_probe_program_search_for_old_horizontal_machine))
                print(result_of_probe_program_search_for_old_horizontal_machine)

                # =======================================================================================|
                # [#] Define list to add all ProbPrograms lines one by one After Finding the ProbProgram.|
                # =======================================================================================|
                probe_programs_lines_of_old_horizontal_machine = []

                # ====================================================================================================|
                # [#] Use (if statement) to check ProbPrograms Search results.                                        |
                # [#] If number of ProbPrograms that founded is <1>, open the program and take the first line and     |
                #     save it in another variable to use it later in horizontal template.                             |
                # [#] First line of ProbProgram will be something like: 'O6446 (F6446XA0)', 'OJ160 (FJE160-HEX)'...Etc|
                #     which it is what horizontal template needs to call the ProbProgram.                             |
                # ====================================================================================================|
                if ((len(result_of_probe_program_search_for_old_horizontal_machine) == 1) and
                        (forging_number != "" and forging_number is not None)):
                    print("RESULT_OF_PROBE_PROGRAM_SEARCH: ", result_of_probe_program_search_for_old_horizontal_machine)

                    # =================================================================================================|
                    # [#] Use <with open()> method to open the program (by use index[0] of the results list) as        |
                    #     current file to iterate through its lines and add them to the ProbPrograms lines list.       |
                    # [#] Use 'rt' to: read a file as text.                                                            |
                    # [#] Use [line.rstrip('\n')] to strip newline and add it to list (ie:[Element],new line,[Element])|
                    # =================================================================================================|
                    with open(result_of_probe_program_search_for_old_horizontal_machine[0], 'rt') as current_program:
                        for line in current_program:
                            probe_programs_lines_of_old_horizontal_machine.append(line.rstrip('\n'))

                        # ==========================================================================================|
                        # [#] Take the first line and save it in another variable to use it later in horiz template.|
                        # ==========================================================================================|
                        probe_program_of_old_horizontal_machine = probe_programs_lines_of_old_horizontal_machine[0]

                # <<Keep them until we make sure we don't need them>>
                # maybe we don't need it
                # elif (probe_programs_folder_path_of_old_horizontal_machine == ""):
                #     fail_messages_of_creating_old_horizontal_machine_program.append(
                #         "Failed to Find, Load, or Access" + '[b][u][color=ffffff] Probe Programs [/color][/u][/b]' +
                #         "File to Create the Program." + "\n" + "Double Check Network, and File Location." + "\n")
                #     email_messages_of_creating_old_horizontal_machine_program.append(
                #         "Failed to Find, Load, or Access Probe Programs File to Create the Program." + "\n" +
                #         "Double Check Network, and File Location." + "\n")
                #     failed_to_create_old_horizontal_machine_program(self)
                #     return

                # ======================================|
                # [#] In case Forging Number is missing.|
                # maybe it's not needed                 |
                # ======================================|
                elif (forging_number == "" or forging_number is None):
                    fail_messages_of_creating_old_horizontal_machine_program.append(
                        "Forging Number does NOT found." + '\n' +
                        "Double Check Job Spec with Engineering and Try Again." +
                        "\n")
                    email_messages_of_creating_old_horizontal_machine_program.append("\n".join(
                        fail_messages_of_creating_old_horizontal_machine_program))
                    failed_to_create_old_horizontal_machine_program(self)
                    return
                # IF NO PROGRAM FOUND, WE NEED TO CHECK MANUALLY IF IT IS THERE ,
                # OTHERWISE CREATE NEW PROBE PROGRAM AND TRY AGAIN      <and forging_number != "F02567X">

                # ================================================================================================|
                # [#] If number of ProbPrograms that founded is <0>, that's indicate there is NO ProbProgram found|
                #     for the Forging Number, therefor User needs to Double Check Probe Programs Folder in case   |
                #     it saved in another way, if it is NOT there, User needs to Create new Probe Program.        |
                # [#] Call (failed_to_create_old_horizontal_machine_program) Function to warn the user.           |
                # [#] Use (return) to stop the function (APP as a result) from executing the rest of the code.    |
                # ================================================================================================|
                elif ((len(result_of_probe_program_search_for_old_horizontal_machine) == 0) and
                      (forging_number != "" and forging_number is not None)):
                    fail_messages_of_creating_old_horizontal_machine_program.append(
                        "Probe Program does NOT found for Forging number of " + forging_number + '\n' +
                        "Double Check Probe Programs Folder (maybe it saved for 'X' not 'Z' or vice versa)." + '\n' +
                        "If it is NOT there, Create new Probe Program and Try Again."
                        + '\n' + '\n' + "Otherwise Double Check Network, and File Location." + "\n")
                    email_messages_of_creating_old_horizontal_machine_program.append("\n".join(
                        fail_messages_of_creating_old_horizontal_machine_program))
                    failed_to_create_old_horizontal_machine_program(self)
                    return

                # =====================================================================================================|
                # [#] If number of ProbPrograms that founded is More than <1>, that's indicate many ProbPrograms found |
                #   for the Forging Number, therefor User needs to Double Check Probe Programs Folder and delete the   |
                #   Unnecessary programs (sometimes that happens if program saved as test or specific edition Versions)|
                # [#] Call (failed_to_create_old_horizontal_machine_program) Function to warn the user.                |
                # [#] Use (return) to stop the function (APP as a result) from executing the rest of the code.         |
                # =====================================================================================================|
                elif ((len(result_of_probe_program_search_for_old_horizontal_machine) > 1) and
                      (forging_number != "" and forging_number is not None)):

                    # ===============================================================================================|
                    # [#] Inform the User of the result of ProbPrograms that found to help to Fix the Confusion.     |
                    # [#] If number of ProbPrograms that founded is less than <20>, show them to the User.           |
                    # [#] If number of ProbPrograms that founded is more than <20> (it's rear), show the User general|
                    #     message without the ProbPrograms because there is no enough room for more than 20 programs.|
                    # ===============================================================================================|
                    if (len(result_of_probe_program_search_for_old_horizontal_machine) <= 20):
                        fail_messages_of_creating_old_horizontal_machine_program.append(
                            "Many Probe Programs found for this Forging " + forging_number + '\n' + '\n' +
                            ('\n'.join(result_of_probe_program_search_for_old_horizontal_machine)) + '\n' + '\n' +
                            "Fix the Confusion and Try Again." + "\n")
                    else:
                        fail_messages_of_creating_old_horizontal_machine_program.append(
                            "Many Probe Programs found for this Forging " + forging_number + '\n' +
                            "Fix the Confusion and Try Again." + "\n")
                    email_messages_of_creating_old_horizontal_machine_program.append("\n".join(
                        fail_messages_of_creating_old_horizontal_machine_program))
                    print(len(result_of_probe_program_search_for_old_horizontal_machine))
                    print(result_of_probe_program_search_for_old_horizontal_machine)
                    failed_to_create_old_horizontal_machine_program(self)
                    return

            # =================================================================================================|
            # [#] Use (except) Block to Handle any Error may occur when Accessing or Finding Files and Folders.|
            # =================================================================================================|
            except Exception as error:
                fail_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find, Load, or Access" + '[b][u][color=ffffff] Probe Programs [/color][/u][/b]' +
                    "Folder to Create the Program." + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]' + str(
                        error) + '[/color]' + "\n" + "Double Check Network, and File Location." + "\n")
                email_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find, Load, or Access Probe Programs File to Create the Program." + "\n" +
                    "Double Check Network, and File Location." + "\n")
                failed_to_create_old_horizontal_machine_program(self)
                return
        # endregion <<<<====================================[PROBE PROGRAMS]====================================>>>>

        # region  <<<<============================[Old Horizontal For-Loop Template]============================>>>>

        # =======================================================================================|
        # [#] Define Variable to set the Path of the Old Horizontal Template.                    |
        # [#] Access MDTextField of (id: HorizontalTemplate) in (OldHorizontalSettingScreen) from|
        #     Screens_Builder to get Old Horizontal Template Path.                               |
        # =======================================================================================|
        old_horizontal_template_file_path = self.manager.get_screen('OldHorizontalSettingScreen').ids[
            "HorizontalTemplate"].text
        print(old_horizontal_template_file_path)

        # ==============================================================================|
        # [#] Define list to add all OldHorizontalTemplate lines one by one to the list.|
        # ==============================================================================|
        global pin_bore_program_lines_of_old_horizontal_machine
        pin_bore_program_lines_of_old_horizontal_machine = []

        # TO OPEN HORIZONTAL_TEMPLATE AND ADD EACH SINGLE LINE TO list WE CREATE
        # ABOVE(pin_bore_program_lines_of_old_horizontal_machine)
        # ======================================================================================================|
        # [#] Use <with open()> method to open the template as current file to iterate through its lines and    |
        #     add(append) them to the OldHorizontalTemplateLines List.                                          |
        # [#] Use 'rt' to: read a file as text.                                                                 |
        # [#] Use [line.rstrip('\n')] to strip newline and add it to list (ie:[Element],new line,[Element]).    |
        # [#] Use (try/except) Blocks to Handle any Error may occur when Accessing or Finding Files and Folders.|
        # ======================================================================================================|
        try:
            with open(old_horizontal_template_file_path, 'rt') as current_program:
                for line in current_program:
                    pin_bore_program_lines_of_old_horizontal_machine.append(line.rstrip('\n'))

                # =======================================================================================|
                # [#] Access the First Element of List of (OldHorizontalTemplateLines) by use ([0] index)|
                #     to Set the First line of the Program with all changes that needed.                 |
                #     (Ex: (PART WD-13000 -- 05/14/2022 <SYS>)).                                         |
                # [#] <SYS>: Used to Indicate this Program is Created By 'WisecoProgramsMaker' APP.      |
                # =======================================================================================|
                pin_bore_program_lines_of_old_horizontal_machine[0] = (
                        '(PART ' + new_program_number_for_old_horizontal_machine + ' -- ' + today_date +
                        ' <SYS>' + ')')

                # region  <<<<============================[Tool List]============================>>>>

                # ==============================================================================================|
                # [#] Use (for and while) loops with (if statement) for each possible tool to add them to the   |
                #     Tool List of the Program.                                                                 |
                # [#] The Tool List will be different from Job to Job (related to features that job has),       |
                #     in other words, there is no specific line in the template can access to change or modify, |
                #     therefor it needs to add the desire tool to the template as a new line, to be able to do  |
                #     that, it needs to find index (position) of the Tool List in the template to start add     |
                #     the tools one by one.                                                                     |
                # [#] Use the line of '(**********TOOL LIST**********)' from the template to set the index of   |
                #     the Tool List to start add the applicable tools.                                          |
                # [#] Define Variable to use in the (while) loop condition to indicate the end of the tool list,|
                #     Set it to be <0> in the beginning, and change it to <1> after finish adding the last tool |
                #     to the Tool List to avoid enter the (for) loop over and over again.                       |
                # ==============================================================================================|
                end_of_tool_list_of_old_horizontal_program = 0
                for line in pin_bore_program_lines_of_old_horizontal_machine:
                    while end_of_tool_list_of_old_horizontal_program == 0:

                        # ==============================================================================|
                        # [#] Use the line of '(**********TOOL LIST**********)' from the template to set|
                        #     the index (location) of the Tool List.                                    |
                        # [#] Use (try/except) Blocks to Handle any Error of not finding the line of    |
                        #     '(**********TOOL LIST**********)' in the template.                        |
                        # ==============================================================================|
                        try:
                            tool_list_index = pin_bore_program_lines_of_old_horizontal_machine.index(
                                    '(**********TOOL LIST**********)')
                            # =======================================================================================|
                            # [#] Add <1> to the tool_list_index to start adding the applicable tools after the line.|
                            # =======================================================================================|
                            tool_list_index += 1
                        except Exception as error:
                            fail_messages_of_creating_old_horizontal_machine_program.append(
                                "Failed to Find the Index of Tool List in " +
                                '[b][u][color=ffffff] Horizontal template. [/color][/u][/b]' +
                                "to add the tools." + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]' +
                                str(error) + '[/color]' + "\n" +
                                "Make sure Horizontal template contain this line EXACTLY as it is:"
                                + "\n" + "(**********TOOL LIST**********)" + "\n")
                            email_messages_of_creating_old_horizontal_machine_program.append(
                                "Failed to Find the Index of Tool List in Horizontal template." + "\n" + "\n" +
                                "Make sure Horizontal template contain this line EXACTLY as it is:" + "\n" + "\n" +
                                "`(**********TOOL LIST**********)`" + "\n")
                            failed_to_create_old_horizontal_machine_program(self)
                            return

                        # region  <<<<============================[Pilot Bore Tool]============================>>>>

                        # ==========================================================================================|
                        # [#] PilotBoreDepth is the value that need to use in the program, but the DataBase has both|
                        #     values of PilotBoreDepthToDome and PilotBoreDepthToDeck, and one of them will use as  |
                        #     PilotBoreDepth that needed (it depends on the Piston and Forging info), therefor it   |
                        #     needs a logic to decide which one can use.                                            |
                        # ==========================================================================================|
                        global pilot_bore_depth

                        # ==========================================================================================|
                        # [#] If PilotBoreDepthToDeck (from Piston DataBase) has a value and                        |
                        #     ForgingHollowDomeRise (from Forging DataBase) has a value then Set:                   |
                        #     PilotBoreDepth to be equal PilotBoreDepthToDeck.                                      |
                        # [#] Use ForgingHollowDomeRise because it's actually what it use to calculate the          |
                        #     PilotBoreDepthToDeck value, in other words, PilotBoreDepth it will be always (To Dome)|
                        #     unless the Forging has a value in ForgingHollowDomeRise, and it calculated by:        |
                        #     PilotBoreDepthToDeck = PilotBoreDepthToDome - ForgingHollowDomeRise.                  |
                        # [#] Also use ForgingHollowDomeRise value as safety check of using PilotBoreDepthToDeck,   |
                        #     because many old jobs has PilotBoreDepthToDeck value (when it should NOT), but it     |
                        #     happened when transfer from Qantel spec to the new Spec Format.                       |
                        # [#] For 'pilot_bore_depth_verifying_status' explanation check:                            |
                        #     pilot_bore_depth_verifying_status <Note> in [Four Cycle Pin Bore Function] Section.   |
                        # ==========================================================================================|
                        if ((pilot_bore_depth_to_deck is not None and pilot_bore_depth_to_deck != 0 and
                             pilot_bore_depth_to_deck != "")
                                and (forging_hollow_dome_rise is not None and forging_hollow_dome_rise != 0 and
                                     forging_hollow_dome_rise != "") and pilot_bore_depth_verifying_status == False):
                            pilot_bore_depth = pilot_bore_depth_to_deck

                        # ============================================================================================|
                        # [#] While condition above is not apply, Check If PilotBoreDepthToDome (from Piston DataBase)|
                        #     has a value then Set: PilotBoreDepth to be equal PilotBoreDepthToDome.                  |
                        # [#] For 'pilot_bore_depth_verifying_status' explanation check:                              |
                        #     pilot_bore_depth_verifying_status <Note> in [Four Cycle Pin Bore Function] Section.     |
                        # ============================================================================================|
                        elif ((pilot_bore_depth_to_dome is not None and pilot_bore_depth_to_dome != 0 and
                               pilot_bore_depth_to_dome != "") and pilot_bore_depth_verifying_status == False):
                            pilot_bore_depth = pilot_bore_depth_to_dome

                        # <<Keep them until we make sure we don't need them>>
                        # ===============================================================================|
                        # [#] If none of the above conditions apply set the value to be equal <0> to warn|
                        #     the user of the issue to fix it.                                           |
                        # ===============================================================================|
                        # else:
                        #     pilot_bore_depth = 0
                        print("[#]pilot_bore_depth FOR " + new_program_number_for_old_horizontal_machine + ":")
                        print("     ", pilot_bore_depth)

                        # =========================================================================================|
                        # [#] PilotBoreDiameter could be cut by tool of <2.25> or <1.70> diameter size, therefor it|
                        #     needs a logic to decide which one can use.                                           |
                        # [#] For 'pilot_availability_status' explanation check:                                   |
                        #     pilot_availability_status <Note> in [Four Cycle Pin Bore Function] Section.          |
                        # =========================================================================================|
                        global pilot_availability_status
                        if (pilot_diameter is not None and pilot_diameter != ""):

                            # ========================================================================================|
                            # [#] If PilotBoreDiameter value (from Piston DataBase) is equal <2.25>:                  |
                            #    -Add the tool description of (2.25 Diameter) from tool list from the Excel Sheet.    |
                            #    -Set pilot_availability_status to equal <1>.                                         |
                            #    -NO needs to change the tool number in 'CONSTANTS' section in the template because   |
                            #     it use tool number <01> by default which is the tool number of (2.25 Diameter) tool.|
                            #    -Add <1> to the tool_list_index to add the next applicable tool after this tool.     |
                            # ========================================================================================|
                            if (pilot_diameter >= 2.2495 and pilot_diameter <= 2.2505):
                                pilot_availability_status = 1

                                # =====================================================================================|
                                # [#] Steps to add the tool description from the Excel Sheet to ToolList in template:  |
                                #  [#] Needs to find the index of the tool in the 'miscellaneous_tool_list' (that set  |
                                #      earlier in [Load Horizontal Sheets Function] section) by use the                |
                                #      ToolUsage of the tool which is ('2.250 PILOT BORE FOR 2.25 DIA').               |
                                #  [#] Store the INDEX founded above in variable of 'miscellaneous_tool_list_index'.   |
                                #  [#] Add the tool description to the ToolList of the template by:                    |
                                #   -Use <insert> method to OldHorizontalTemplateLines list in current tool_list_index.|
                                #   -Use <PandasLibrary> method to locate the tool description by Access the Excel File|
                                #    'horizontal_tool_list_file' and Sheet of name ['MISCELLANEOUS_TOOL_LIST'] and     |
                                #    column of ['DESCRIPTION(FOR_MACHINES_27/28/32)'] in location of                   |
                                #    'miscellaneous_tool_list_index' that founded above.                               |
                                # =====================================================================================|
                                miscellaneous_tool_list_index = miscellaneous_tool_list.index(
                                    '2.250 PILOT BORE FOR 2.25 DIA')
                                pin_bore_program_lines_of_old_horizontal_machine.insert(
                                    tool_list_index, horizontal_tool_list_file['MISCELLANEOUS_TOOL_LIST'].loc[
                                        miscellaneous_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                                tool_list_index += 1

                            # ====================================================================================|
                            # [#] If PilotBoreDiameter value (from Piston DataBase) is equal <1.70>:              |
                            #    -Add the tool description of (1.70 Diameter) from tool list from the Excel Sheet.|
                            #    -Set pilot_availability_status to equal <1>.                                     |
                            #    -Needs to change the tool number in 'CONSTANTS' section in the template to use   |
                            #     the tool number of (1.70 Diameter) tool because the default tool number is '01' |
                            #     which is the tool number of (2.25 Diameter) tool.                               |
                            #    -Add <1> to the tool_list_index to add the next applicable tool after this tool. |
                            # ====================================================================================|
                            elif (pilot_diameter >= 1.6995 and pilot_diameter <= 1.7005):
                                pilot_availability_status = 1

                                # =====================================================================================|
                                # [#] Steps to add the tool description from the Excel Sheet to ToolList in template:  |
                                #  [#] Needs to find the index of the tool in the 'miscellaneous_tool_list' (that set  |
                                #      earlier in [Load Horizontal Sheets Function] section) by use the                |
                                #      ToolUsage of the tool which is ('1.70 PILOT BORE FOR 1.70 DIA').                |
                                #  [#] Store the INDEX founded above in variable of 'miscellaneous_tool_list_index'.   |
                                #  [#] Add the tool description to the ToolList of the template by:                    |
                                #   -Use <insert> method to OldHorizontalTemplateLines list in current tool_list_index.|
                                #   -Use <PandasLibrary> method to locate the tool description by Access the Excel File|
                                #    'horizontal_tool_list_file' and Sheet of name ['MISCELLANEOUS_TOOL_LIST'] and     |
                                #    column of ['DESCRIPTION(FOR_MACHINES_27/28/32)'] in location of                   |
                                #    'miscellaneous_tool_list_index' that founded above.                               |
                                # =====================================================================================|
                                miscellaneous_tool_list_index = miscellaneous_tool_list.index(
                                    '1.70 PILOT BORE FOR 1.70 DIA')
                                pin_bore_program_lines_of_old_horizontal_machine.insert(
                                    tool_list_index, horizontal_tool_list_file['MISCELLANEOUS_TOOL_LIST'].loc[
                                        miscellaneous_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                                tool_list_index += 1

                                # ====================================================================================|
                                # [#] Steps to change the tool description in (**CONSTANTS**) section in the template:|
                                #   [#] Needs to find the index of the default tool description which is:             |
                                #       '(T01 IS THE STD 2.250 DIA. PILOT BORE TOOL)'.                                |
                                #   [#] Store the INDEX founded above in variable of 'pilot_tool_note_index'.         |
                                #   [#] Change the tool description by:                                               |
                                #      -Use 'pilot_tool_note_index' that founded above.                               |
                                #      -Use <PandasLibrary> method to locate the tool description by Access the Excel |
                                #       File 'horizontal_tool_list_file' and Sheet of name ['MISCELLANEOUS_TOOL_LIST']|
                                #       and column of ['DESCRIPTION(FOR_MACHINES_27/28/32)'] in location of           |
                                #       'miscellaneous_tool_list_index' that founded above.                           |
                                # ====================================================================================|
                                pilot_tool_note_index = pin_bore_program_lines_of_old_horizontal_machine.index(
                                    '(T01 IS THE STD 2.250 DIA. PILOT BORE TOOL)')
                                pin_bore_program_lines_of_old_horizontal_machine[pilot_tool_note_index] = \
                                    horizontal_tool_list_file['MISCELLANEOUS_TOOL_LIST'].loc[
                                        miscellaneous_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)']

                                # ===================================================================================|
                                # [#] Steps to change the tool number in (**CONSTANTS**) section in the template:    |
                                #   [#] Needs to find the index of the default tool number which is:                 |
                                #       'VC101=01'.                                                                  |
                                #   [#] Store the INDEX founded above in variable of 'VC101_variable_index'.         |
                                #   [#] Change the tool number by:                                                   |
                                #      -Use 'VC101_variable_index' that founded above.                               |
                                #      -Use <PandasLibrary> method to locate the tool number by Access the Excel File|
                                #       'horizontal_tool_list_file' and Sheet of name ['MISCELLANEOUS_TOOL_LIST'] and|
                                #       column of ['TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)'] in location of        |
                                #       'miscellaneous_tool_list_index' that founded above.                          |
                                # ===================================================================================|
                                VC101_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(
                                    'VC101=01')
                                pin_bore_program_lines_of_old_horizontal_machine[VC101_variable_index] = \
                                    ('VC101=' + str(horizontal_tool_list_file['MISCELLANEOUS_TOOL_LIST'].loc[
                                        miscellaneous_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']))

                        # ============================================================================================|
                        # [#] If Pilot Diameter didn't Detected above and status doesn't change to <1>,               |
                        #     (ie: pilot_availability_status still equal ""), it's need to ask the user to choose     |
                        #     the Pilot Diameter (Big or Small) by call function of (need_confirmation).               |
                        # [#] Many old jobs are missing the Pilot Diameter as value from the DataBase, instead, it's  |
                        #      mentioned in the 'Legacy Comments', because of that needs to check and ask user to fix.|
                        # ============================================================================================|
                        if (pilot_availability_status == ""):

                            # =========================================================|
                            # [#] Set verification_message to be the current Situation.|
                            # =========================================================|
                            verification_messages_of_creating_old_horizontal_machine_program = [
                                "[u]Pilot Bore Diameter[/u] does NOT found, Please Choose the Diameter"]

                            # ==============================================================|
                            # [#] Add verification_message to the Confirmation Details list.|
                            # ==============================================================|
                            old_horizontal_program_confirmation_email_message_list.append(
                                "# It was need confirmation of :" + "\n" + (
                                    "**Pilot Bore Diameter** does NOT found, Please Choose the Diameter"))

                            # ===============================================================================|
                            # [#] Set the confirmation details that's need to finalize by User,which they are|
                            #     the Function parameters (title, sub_function, dialog_type, content).       |
                            # ===============================================================================|
                            # ==================================================================|
                            # [#] title: set it to have the header and the verification_message.|
                            # ==================================================================|
                            self.title = '[color=0066ff]Confirmation Message[/color] ' + '\n' + '\n' + \
                                         '[b][i][color=ffffff]' + \
                                         verification_messages_of_creating_old_horizontal_machine_program[0] + \
                                         '[/color][/i][/b]'

                            # ======================================================================|
                            # [#] Set option of the Big Pilot Diameter.                             |
                            # [#] OldHorizontalMachineItem:                                         |
                            #     The Class that created above to be able to create List's Items.   |
                            # [#] set_big_pilot_bore_diameter:                                      |
                            #     The Function that created below to set the Big Pilot diameter.    |
                            # [#] big_pilot_bore_diameter_Option_Status:                            |
                            #     Set Status to be 'False' by default to indicate is not picked yet,|
                            #     and it will change to be 'True' when user choose this option.     |
                            # ======================================================================|
                            self.big_pilot_bore_diameter = OldHorizontalMachineItem(
                                text="2.25 Pilot Bore Diameter", on_release=self.set_big_pilot_bore_diameter)
                            self.big_pilot_bore_diameter_Option_Status = False

                            # ======================================================================|
                            # [#] Set option of the Small Pilot Diameter.                           |
                            # [#] OldHorizontalMachineItem:                                         |
                            #     The Class that created above to be able to create List's Items.   |
                            # [#] set_small_pilot_bore_diameter:                                    |
                            #     The Function that created below to set the Small Pilot diameter.  |
                            # [#] small_pilot_bore_diameter_Option_Status:                          |
                            #     Set Status to be 'False' by default to indicate is not picked yet,|
                            #     and it will change to be 'True' when user choose this option.     |
                            # ======================================================================|
                            self.small_pilot_bore_diameter = OldHorizontalMachineItem(
                                text="1.70 Pilot Bore Diameter", on_release=self.set_small_pilot_bore_diameter)
                            self.small_pilot_bore_diameter_Option_Status = False

                            # ===========================================|
                            # [#] Create items list to have both options.|
                            # ===========================================|
                            self.items = [self.big_pilot_bore_diameter, self.small_pilot_bore_diameter]

                            # =============================================================================|
                            # [#] Call Function of (need_confirmation) by pass the parameters that created:|
                            #   [#] title: self.title                                                      |
                            #   [#] sub_function: self.choose_pilot_bore_diameter (it's Function           |
                            #                     created below to set the Pilot diameter).                |
                            #   [#] dialog_type: "confirmation" (it used to indicate that user             |
                            #                     needs to choose Option).                                 |
                            #   [#] content: self.items (the ItemsList that have the options).             |
                            # =============================================================================|
                            need_confirmation_to_create_old_horizontal_machine_program(
                                self, self.title, self.choose_pilot_bore_diameter, "confirmation", self.items)
                            print("pilot_availability_status in create function: ", pilot_availability_status)
                            return
                        # endregion  <<<<============================[Pilot Bore Tool]============================>>>>

                        # region  <<<<=========================[Legacy Comments Check]===========================>>>>

                        # ====================================================================================|
                        # [#] Many old Jobs have 'Legacy Comments' that contains Some PinBore Info, because of|
                        #     that Needs to check the 'Legacy Comments' to get or verify these Information.   |
                        # [#] Put this section here (before everything) to make sure to have the accurate info|
                        #     of (pilot_bore_depth, pilot_to_pin, and pin_hole_diameter).                     |
                        # ====================================================================================|

                        # region  <<<<======================[Pilot Bore Depth Value Check]=========================>>>>

                        # =======================================================================================|
                        # [#] Check if the Job has 'Legacy Comments' and the status not verified yet.            |
                        # [#] For 'pilot_bore_depth_verifying_status' explanation check:                         |
                        #     pilot_bore_depth_verifying_status <Note> in [Four Cycle Pin Bore Function] Section.|
                        # =======================================================================================|
                        if (legacy_pin_bore_comments is not None and legacy_pin_bore_comments != "" and
                                pilot_bore_depth_verifying_status == False):

                            # =========================================================|
                            # [#] Set verification_message to be the current Situation.|
                            # =========================================================|
                            verification_messages_of_creating_old_horizontal_machine_program = [
                             "Look to Legacy Comments and Double Check" + ' [u]Pilot Bore Depth[/u], ' +
                             "Correct The Value If Necessary :"]

                            # ==============================================================|
                            # [#] Add verification_message to the Confirmation Details list.|
                            # ==============================================================|
                            old_horizontal_program_confirmation_email_message_list.append(
                                "# It was need confirmation of :" + "\n" + (
                                    "Look to Legacy Comments and Double Check **Pilot Bore Depth**, "
                                    "Correct The Value If Necessary"))

                            # ==================================================================|
                            # [#] title: set it to have the header and the verification_message.|
                            # ==================================================================|
                            self.title = '[color=0066ff]Confirmation Message[/color]' + '\n' + \
                                         '\n' + '[b][i][color=ffffff]' + \
                                         verification_messages_of_creating_old_horizontal_machine_program[0] + \
                                         '[/color][/i][/b]'

                            # ====================================================================|
                            # [#] Create the window_message that contain the Confirmation Details.|
                            # ====================================================================|
                            self.Dialog_BoxLayout = BoxLayout(height=30)

                            # =============================================================================|
                            # [#] Create the TextInput field to display the value from 'Legacy Comments' to|
                            #     let user verify or correct the value.                                    |
                            # [#] (input_filter="float"): to accept Numeric input Only.                    |
                            # =============================================================================|
                            self.pilot_bore_depth_confirmation_text_field = TextInput(
                                text=str(pilot_bore_depth), multiline=False, input_filter="float",
                                background_color=[60 / 255.0, 60 / 255.0, 60 / 255.0, 1],
                                foreground_color=[1, 1, 1, 1], size_hint_y=None, height=30)

                            # ===========================================================================|
                            # [#] Add the TextInput field created above to be part of the window_message.|
                            # ===========================================================================|
                            self.Dialog_BoxLayout.add_widget(self.pilot_bore_depth_confirmation_text_field)

                            # =================================================================================|
                            # [#] Call Function of (need_confirmation) by pass the parameters that created:    |
                            #   [#] title: self.title                                                          |
                            #   [#] sub_function: self.enter_pilot_bore_depth_value (it's the Function created |
                            #                     below to set the Pilot Bore Depth).                          |
                            #   [#] dialog_type: "custom" (it used to indicate that user needs to enter value).|
                            #   [#] content: self.Dialog_BoxLayout (the window_message that have the Details). |
                            # =================================================================================|
                            need_confirmation_to_create_old_horizontal_machine_program(
                                self, self.title, self.enter_pilot_bore_depth_value, "custom", self.Dialog_BoxLayout)
                            print("Pilot Bore Depth BEFORE User Input: ", pilot_bore_depth)
                            return
                        # endregion  <<<<=====================[Pilot Bore Depth Value Check]=======================>>>>

                        # region  <<<<====================[Pilot To Pin Value Check]=====================>>>>

                        # =======================================================================================|
                        # [#] Check if the Job has 'Legacy Comments' and the status not verified yet.            |
                        # [#] For 'pilot_to_pin_verifying_status' explanation check:                             |
                        #     pilot_bore_depth_verifying_status <Note> in [Four Cycle Pin Bore Function] Section.|
                        # =======================================================================================|
                        if (legacy_pin_bore_comments is not None and legacy_pin_bore_comments != "" and
                                pilot_to_pin_verifying_status == False):

                            # =========================================================|
                            # [#] Set verification_message to be the current Situation.|
                            # =========================================================|
                            verification_messages_of_creating_old_horizontal_machine_program = [
                                "Look to Legacy Comments and Double Check" + ' [u]Pilot To Pin[/u], ' + "\n" +
                                "Correct The Value If Necessary :"]

                            # ==============================================================|
                            # [#] Add verification_message to the Confirmation Details list.|
                            # ==============================================================|
                            old_horizontal_program_confirmation_email_message_list.append(
                                "# It was need confirmation of :" + "\n" + (
                                    "Look to Legacy Comments and Double Check **Pilot To Pin**, "
                                    "Correct The Value If Necessary"))

                            # ==================================================================|
                            # [#] title: set it to have the header and the verification_message.|
                            # ==================================================================|
                            self.title = '[color=0066ff]Confirmation Message[/color]' + '\n' + \
                                         '\n' + '[b][i][color=ffffff]' + \
                                         verification_messages_of_creating_old_horizontal_machine_program[0] + \
                                         '[/color][/i][/b]'

                            # ====================================================================|
                            # [#] Create the window_message that contain the Confirmation Details.|
                            # ====================================================================|
                            self.Dialog_BoxLayout = BoxLayout(height=30)

                            # =============================================================================|
                            # [#] Create the TextInput field to display the value from 'Legacy Comments' to|
                            #     let user verify or correct the value.                                    |
                            # [#] (input_filter="float"): to accept Numeric input Only.                    |
                            # =============================================================================|
                            self.pilot_to_pin_confirmation_text_field = TextInput(
                                text=str(pilot_to_pin), multiline=False, input_filter="float",
                                background_color=[60 / 255.0, 60 / 255.0, 60 / 255.0, 1],
                                foreground_color=[1, 1, 1, 1], size_hint_y=None, height=30)

                            # ===========================================================================|
                            # [#] Add the TextInput field created above to be part of the window_message.|
                            # ===========================================================================|
                            self.Dialog_BoxLayout.add_widget(self.pilot_to_pin_confirmation_text_field)

                            # =================================================================================|
                            # [#] Call Function of (need_confirmation) by pass the parameters that created:    |
                            #   [#] title: self.title                                                          |
                            #   [#] sub_function: self.enter_pilot_to_pin_value (it's the Function created     |
                            #                     below to set the Pilot To Pin).                              |
                            #   [#] dialog_type: "custom" (it used to indicate that user needs to enter value).|
                            #   [#] content: self.Dialog_BoxLayout (the window_message that have the Details). |
                            # =================================================================================|
                            need_confirmation_to_create_old_horizontal_machine_program(
                                self, self.title, self.enter_pilot_to_pin_value, "custom", self.Dialog_BoxLayout)
                            print("Pilot To Pin BEFORE User Input: ", pilot_to_pin)
                            return
                        # endregion  <<<<===================[Pilot To Pin Value Check]====================>>>>

                        # region  <<<<====================[Pin Hole Diameter Value Check]=====================>>>>

                        # ========================================================================================|
                        # [#] Check if the Job has 'Legacy Comments' and the status not verified yet.             |
                        # [#] For 'pin_hole_diameter_verifying_status' explanation check:                         |
                        #     pin_hole_diameter_verifying_status <Note> in [Four Cycle Pin Bore Function] Section.|
                        # ========================================================================================|
                        if (legacy_pin_bore_comments is not None and legacy_pin_bore_comments != "" and
                                pin_hole_diameter_verifying_status == False):

                            # =========================================================|
                            # [#] Set verification_message to be the current Situation.|
                            # =========================================================|
                            verification_messages_of_creating_old_horizontal_machine_program = [
                             "Look to Legacy Comments and Double Check" + ' [u]Pin Hole Diameter[/u], ' +
                             "Correct The Value If Necessary :"]

                            # ==============================================================|
                            # [#] Add verification_message to the Confirmation Details list.|
                            # ==============================================================|
                            old_horizontal_program_confirmation_email_message_list.append(
                                "# It was need confirmation of :" + "\n" + (
                                    "Look to Legacy Comments and Double Check **Pin Hole Diameter**, "
                                    "Correct The Value If Necessary"))

                            # ==================================================================|
                            # [#] title: set it to have the header and the verification_message.|
                            # ==================================================================|
                            self.title = '[color=0066ff]Confirmation Message[/color]' + '\n' + \
                                         '\n' + '[b][i][color=ffffff]' + \
                                         verification_messages_of_creating_old_horizontal_machine_program[0] + \
                                         '[/color][/i][/b]'

                            # ====================================================================|
                            # [#] Create the window_message that contain the Confirmation Details.|
                            # ====================================================================|
                            self.Dialog_BoxLayout = BoxLayout(height=30)

                            # =============================================================================|
                            # [#] Create the TextInput field to display the value from 'Legacy Comments' to|
                            #     let user verify or correct the value.                                    |
                            # [#] (input_filter="float"): to accept Numeric input Only.                    |
                            # =============================================================================|
                            self.pin_hole_diameter_confirmation_text_field = TextInput(
                                text=str(pin_hole_diameter), multiline=False, input_filter="float",
                                background_color=[60 / 255.0, 60 / 255.0, 60 / 255.0, 1],
                                foreground_color=[1, 1, 1, 1], size_hint_y=None, height=30)

                            # ===========================================================================|
                            # [#] Add the TextInput field created above to be part of the window_message.|
                            # ===========================================================================|
                            self.Dialog_BoxLayout.add_widget(self.pin_hole_diameter_confirmation_text_field)

                            # =================================================================================|
                            # [#] Call Function of (need_confirmation) by pass the parameters that created:    |
                            #   [#] title: self.title                                                          |
                            #   [#] sub_function: self.enter_pin_hole_diameter_value (it's the Function created|
                            #                     below to set the Pin Hole Diameter).                         |
                            #   [#] dialog_type: "custom" (it used to indicate that user needs to enter value).|
                            #   [#] content: self.Dialog_BoxLayout (the window_message that have the Details). |
                            # =================================================================================|
                            need_confirmation_to_create_old_horizontal_machine_program(
                                self, self.title, self.enter_pin_hole_diameter_value, "custom", self.Dialog_BoxLayout)
                            print("Pin Hole Diameter BEFORE User Input: ", pin_hole_diameter)
                            return
                        # endregion  <<<<===================[Pin Hole Diameter Value Check]====================>>>>

                        # endregion  <<<<======================[Legacy Comments Check]============================>>>>

                        # region  <<<<============================[Ledge Cut Tool]============================>>>>

                        # ========================================================================================|
                        # [#] Steps to decide Ledge Cut Status:                                                   |
                        #   [#] Calculate the distance from origin of the piston to the highest point of pin hole.|
                        #   [#] Use ForgingOutsideRingBeltHeight from (Forging DataBase).                         |
                        #   [#] Check Ledge Cut Status:                                                           |
                        #     -if (distance_from_origin_to_highest_point_of_pin_bore - 0.05) <GreaterThan>        |
                        #      forging_out_side_ring_belt_height:                                                 |
                        #         it's NO Need to Use Ledge cut tool.                                             |
                        #     -if distance_from_origin_to_highest_point_of_pin_bore - 0.05) <SmallerThanOrEqual>  |
                        #      forging_out_side_ring_belt_height)):                                               |
                        #         it's Need to Use Ledge cut tool.                                                |
                        #     -The (0.05) above it use to consider the distance of the radius between Top of the  |
                        #      Forging boss and top of the pin hole.                                              |
                        # [#] For Explanation of LedgeCutStatus calculation check WordFile of (Ladge Cut Status)  |
                        #     that locate at ().                                                                  |
                        # ========================================================================================|
                        if ((pilot_bore_depth is not None and pilot_bore_depth != "" and pilot_bore_depth != 0) and
                                (pilot_to_pin is not None and pilot_to_pin != "" and pilot_to_pin != 0) and
                                (pin_hole_diameter is not None and pin_hole_diameter != "" and pin_hole_diameter != 0)):

                            # =======================================================================================|
                            # The first part of (if statement) doesn't affect anything for now, condition need to
                            # change if need to use later (if ever needs to use ledge_counterbore in LadgeCutStatus)
                            # (and ledge_counterbore_diameter != ""): we put it now to make the code work unless>>
                            # >> we figure out way to get Ledge Counterbore Diameter
                            # MOST LIKELY WE DON'T NEED TO WORRY ABOUT THAT IF WE HAVE Counterbore but
                            # <<Keep them until we make sure we don't need them>>
                            # =======================================================================================|
                            if (ledge_counterbore_diameter != 0 and ledge_counterbore_diameter is not None and
                                    ledge_counterbore_diameter != ""):
                                distance_from_origin_to_highest_point_of_pin_bore = round(
                                    pilot_bore_depth - abs(pilot_to_pin) - pin_hole_diameter -
                                    ((ledge_counterbore_diameter - pin_hole_diameter) / 2), 4)
                                print("distance_from_origin_to_highest_point_of_pin_bore: ",
                                      distance_from_origin_to_highest_point_of_pin_bore)
                            else:
                                # ==================================================================================|
                                # Calculation of distance from origin of the piston to the highest point of pinhole:|
                                #       PilotBoreDepth - PilotToPin - PinHoleDiameter                               |
                                # ==================================================================================|
                                distance_from_origin_to_highest_point_of_pin_bore = round(
                                    (pilot_bore_depth - abs(pilot_to_pin) - pin_hole_diameter), 4)
                                print("distance_from_origin_to_highest_point_of_pin_bore: ",
                                      distance_from_origin_to_highest_point_of_pin_bore)
                        else:
                            # ========================================================================================|
                            # if one of (PilotBoreDepth,PilotToPin,PinHoleDiameter) missing set the distance to be <0>|
                            # it will use to let user decide the status.                                              |
                            # ========================================================================================|
                            distance_from_origin_to_highest_point_of_pin_bore = 0

                        # ===================================================================================|
                        # [#] For 'ledge_cut_availability_status' explanation check:                         |
                        #     ledge_cut_availability_status <Note> in [Four Cycle Pin Bore Function] Section.|
                        # ===================================================================================|
                        global ledge_cut_availability_status

                        # ==============================================|
                        # [#] If Statement to decide the LadgeCutStatus.|
                        # ==============================================|
                        # ==================================================================================|
                        #   [#] When the APP can't detect LadgeCutStatus, user will decide and try to create|
                        #       the program again, as a result the status WILL change , this 'pass' below to|
                        #       prevent ask about the status again.                                         |
                        # ==================================================================================|
                        if (ledge_cut_availability_status != ""):
                            pass

                        # ==============================================================================|
                        #   [#] If (distance_from_origin_to...) didn't calculated above, keep the Status|
                        #       without change to let user decide the LadgeCutStatus.                   |
                        # ==============================================================================|
                        elif (distance_from_origin_to_highest_point_of_pin_bore == 0):
                            ledge_cut_availability_status = ""

                        # =============================================================================|
                        #   [#] If Forging Boss Outside Spacing is <0> or 'None', it's mean the Forging|
                        #       is Fully around --> which it doesn't need to Use the Ledge Cut Tool.   |
                        # =============================================================================|
                        elif (forging_outside_boss_spacing == 0 or forging_outside_boss_spacing is None):
                            print("We DON'T Need to Use Ledge cut tool")
                            ledge_cut_availability_status = 0

                        # =================================================================================|
                        #   [#] if (distance_from_origin_to_highest_point_of_pin_bore - 0.05) <GreaterThan>|
                        #       forging_out_side_ring_belt_height:                                         |
                        #       -->   it's NO Need to Use Ledge cut tool.                                  |
                        # =================================================================================|
                        elif ((forging_out_side_ring_belt_height != "" and
                               forging_out_side_ring_belt_height is not None and forging_out_side_ring_belt_height != 0)
                              and ((distance_from_origin_to_highest_point_of_pin_bore - 0.05) >
                                   forging_out_side_ring_belt_height)):
                            print("We DON'T Need to Use Ledge cut tool")
                            ledge_cut_availability_status = 0

                        # ========================================================================================|
                        #   [#] if (distance_from_origin_to_highest_point_of_pin_bore - 0.05) <SmallerThanOrEqual>|
                        #       forging_out_side_ring_belt_height:                                                |
                        #       -->   it's Need to Use Ledge cut tool.                                            |
                        # ========================================================================================|
                        elif ((forging_out_side_ring_belt_height != "" and
                               forging_out_side_ring_belt_height is not None and
                               forging_out_side_ring_belt_height != 0) and
                              ((distance_from_origin_to_highest_point_of_pin_bore - 0.05) <=
                               forging_out_side_ring_belt_height)):
                            print("Need to Use Ledge cut tool")
                            ledge_cut_availability_status = 1

                            # ================================================================|
                            #   [#] If for any reason can't detect the status, keep the Status|
                            #       without change to let user decide the LadgeCutStatus.     |
                            # ================================================================|
                        else:
                            ledge_cut_availability_status = ""

                        # ========================================================================|
                        # [#] If the status didn't detect, Ask the user decide the LadgeCutStatus.|
                        # ========================================================================|
                        if (ledge_cut_availability_status == ""):

                            # =========================================================|
                            # [#] Set verification_message to be the current Situation.|
                            # =========================================================|
                            verification_messages_of_creating_old_horizontal_machine_program = [
                                "[u]Ledge status[/u] can't detect," + '\n' + "Does this job need to use Ladge Tool ?"]

                            # ==============================================================|
                            # [#] Add verification_message to the Confirmation Details list.|
                            # ==============================================================|
                            old_horizontal_program_confirmation_email_message_list.append(
                                "# It was need confirmation of :" + "\n" + (
                                    "**Ledge status** can't detect, Does this job need to use Ladge Tool ?"))

                            # ==================================================================|
                            # [#] title: set it to have the header and the verification_message.|
                            # ==================================================================|
                            self.title = '[color=0066ff]Confirmation Message[/color] ' + '\n' + '\n' + \
                                         '[b][i][color=ffffff]' + \
                                         verification_messages_of_creating_old_horizontal_machine_program[0] + \
                                         '[/color][/i][/b]'

                            # ============================================================================|
                            # [#] Set option of 'Need to Use Ledge cut tool'.                             |
                            # [#] OldHorizontalMachineItem:                                               |
                            #     The Class that created above to be able to create List's Items.         |
                            # [#] need_to_use_ledge_tool_for_old_horizontal_machine:                      |
                            #     The Function that created below to set the status to use Ledge cut tool.|
                            # [#] need_ledge_tool_status:                                                 |
                            #     Set Status to be 'False' by default to indicate is not picked yet,      |
                            #     and it will change to be 'True' when user choose this option.           |
                            # ============================================================================|
                            self.need_to_use_ledge_tool_option_for_old_horizontal_machine = OldHorizontalMachineItem(
                                text="Yes, It is Need",
                                on_release=self.need_to_use_ledge_tool_for_old_horizontal_machine)
                            self.need_ledge_tool_status = False

                            # =======================================================================================|
                            # [#] Set option of 'NO Need to Use Ledge cut tool'.                                     |
                            # [#] OldHorizontalMachineItem:                                                          |
                            #     The Class that created above to be able to create List's Items.                    |
                            # [#] does_not_need_to_use_ledge_tool_for_old_horizontal_machine:                        |
                            #     The Function that created below to set the status to NO need to use Ledge cut tool.|
                            # [#] does_not_need_ledge_tool_status:                                                   |
                            #     Set Status to be 'False' by default to indicate is not picked yet,                 |
                            #     and it will change to be 'True' when user choose this option.                      |
                            # =======================================================================================|
                            self.does_not_need_to_use_ledge_tool_option_for_old_horizontal_machine =\
                                OldHorizontalMachineItem(
                                    text="No, It Does NOT Need",
                                    on_release=self.does_not_need_to_use_ledge_tool_for_old_horizontal_machine)
                            self.does_not_need_ledge_tool_status = False

                            # ===========================================|
                            # [#] Create items list to have both options.|
                            # ===========================================|
                            self.items = [self.need_to_use_ledge_tool_option_for_old_horizontal_machine,
                                          self.does_not_need_to_use_ledge_tool_option_for_old_horizontal_machine]

                            # =============================================================================|
                            # [#] Call Function of (need_confirmation) by pass the parameters that created:|
                            #   [#] title: self.title                                                      |
                            #   [#] sub_function: self.decide_ledge_tool_status_for_old_horizontal_machine |
                            #                     (it's Function created below to set the LedgeCutStatus). |
                            #   [#] dialog_type: "confirmation" (it used to indicate that user             |
                            #                     needs to choose Option).                                 |
                            #   [#] content: self.items (the ItemsList that have the options).             |
                            # =============================================================================|
                            need_confirmation_to_create_old_horizontal_machine_program(
                                self, self.title, self.decide_ledge_tool_status_for_old_horizontal_machine,
                                "confirmation", self.items)
                            print("LEDGE_CUT_AVAILABILITY_STATUS in create function: ", ledge_cut_availability_status)
                            return

                        # =============================================================================|
                        # [#] LedgeCutTool could be <0.625> or <0.375> diameter size, therefor it needs|
                        #     a logic to decide which one can use.                                     |
                        # =============================================================================|
                        if (ledge_cut_availability_status == 1 or (ledge_counterbore_diameter != 0 and
                                                                   ledge_counterbore_diameter is not None)):

                            # =========================================================================================|
                            # [#] If PilotHoleDiameter value (from Piston DataBase) is GreaterThanOrEqual <0.629>:     |
                            #    -Add the tool description of (0.625 Diameter) from tool list from the Excel Sheet.    |
                            #    -NO needs to change the tool number in 'CONSTANTS' section in the template because    |
                            #     it use tool number <06> by default which is the tool number of (0.625 Diameter) tool.|
                            #    -Add <1> to the tool_list_index to add the next applicable tool after this tool.      |
                            # =========================================================================================|
                            if (pin_hole_diameter >= 0.629):

                                # =====================================================================================|
                                # [#] Steps to add the tool description from the Excel Sheet to ToolList in template:  |
                                #  [#] Needs to find the index of the tool in the 'miscellaneous_tool_list' (that set  |
                                #      earlier in [Load Horizontal Sheets Function] section) by use the                |
                                #      ToolUsage of the tool which is ('LEDGE TOOL 0.625 DIA').                        |
                                #  [#] Store the INDEX founded above in variable of 'miscellaneous_tool_list_index'.   |
                                #  [#] Add the tool description to the ToolList of the template by:                    |
                                #   -Use <insert> method to OldHorizontalTemplateLines list in current tool_list_index.|
                                #   -Use <PandasLibrary> method to locate the tool description by Access the Excel File|
                                #    'horizontal_tool_list_file' and Sheet of name ['MISCELLANEOUS_TOOL_LIST'] and     |
                                #    column of ['DESCRIPTION(FOR_MACHINES_27/28/32)'] in location of                   |
                                #    'miscellaneous_tool_list_index' that founded above.                               |
                                # =====================================================================================|
                                miscellaneous_tool_list_index = miscellaneous_tool_list.index('LEDGE TOOL 0.625 DIA')
                                pin_bore_program_lines_of_old_horizontal_machine.insert(
                                    tool_list_index, horizontal_tool_list_file['MISCELLANEOUS_TOOL_LIST'].loc[
                                        miscellaneous_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                                tool_list_index += 1

                            # =====================================================================================|
                            # [#] If PilotHoleDiameter value (from Piston DataBase) is SmallerThan <0.629>:        |
                            #    -Add the tool description of (0.375 Diameter) from tool list from the Excel Sheet.|
                            #    -Needs to change the tool number in 'CONSTANTS' section in the template to use    |
                            #     the tool number of (0.375 Diameter) tool because the default tool number is '06' |
                            #     which is the tool number of (0.625 Diameter) tool.                               |
                            #    -Add <1> to the tool_list_index to add the next applicable tool after this tool.  |
                            # =====================================================================================|
                            elif (pin_hole_diameter < 0.629):

                                # =====================================================================================|
                                # [#] Steps to add the tool description from the Excel Sheet to ToolList in template:  |
                                #  [#] Needs to find the index of the tool in the 'miscellaneous_tool_list' (that set  |
                                #      earlier in [Load Horizontal Sheets Function] section) by use the                |
                                #      ToolUsage of the tool which is ('LEDGE TOOL 0.375 DIA').                        |
                                #  [#] Store the INDEX founded above in variable of 'miscellaneous_tool_list_index'.   |
                                #  [#] Add the tool description to the ToolList of the template by:                    |
                                #   -Use <insert> method to OldHorizontalTemplateLines list in current tool_list_index.|
                                #   -Use <PandasLibrary> method to locate the tool description by Access the Excel File|
                                #    'horizontal_tool_list_file' and Sheet of name ['MISCELLANEOUS_TOOL_LIST'] and     |
                                #    column of ['DESCRIPTION(FOR_MACHINES_27/28/32)'] in location of                   |
                                #    'miscellaneous_tool_list_index' that founded above.                               |
                                # =====================================================================================|
                                miscellaneous_tool_list_index = miscellaneous_tool_list.index('LEDGE TOOL 0.375 DIA')
                                pin_bore_program_lines_of_old_horizontal_machine.insert(
                                    tool_list_index, horizontal_tool_list_file['MISCELLANEOUS_TOOL_LIST'].loc[
                                        miscellaneous_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                                tool_list_index += 1

                                # ====================================================================================|
                                # [#] Steps to change the tool description in (**CONSTANTS**) section in the template:|
                                #   [#] Needs to find the index of the default tool description which is:             |
                                #       '(T06 IS THE STD .625 CARBIDE END MILL LEDGE TOOL)'.                          |
                                #   [#] Store the INDEX founded above in variable of 'ledge_tool_note_index'.         |
                                #   [#] Change the tool description by:                                               |
                                #      -Use 'ledge_tool_note_index' that founded above.                               |
                                #      -Use <PandasLibrary> method to locate the tool description by Access the Excel |
                                #       File 'horizontal_tool_list_file' and Sheet of name ['MISCELLANEOUS_TOOL_LIST']|
                                #       and column of ['DESCRIPTION(FOR_MACHINES_27/28/32)'] in location of           |
                                #       'miscellaneous_tool_list_index' that founded above.                           |
                                # ====================================================================================|
                                ledge_tool_note_index = pin_bore_program_lines_of_old_horizontal_machine.index(
                                    '(T06 IS THE STD .625 CARBIDE END MILL LEDGE TOOL)')
                                pin_bore_program_lines_of_old_horizontal_machine[ledge_tool_note_index] = \
                                    horizontal_tool_list_file['MISCELLANEOUS_TOOL_LIST'].loc[
                                        miscellaneous_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)']

                                # ===================================================================================|
                                # [#] Steps to change the tool number in (**CONSTANTS**) section in the template:    |
                                #   [#] Needs to find the index of the default tool number which is:                 |
                                #       'VC103=06'.                                                                  |
                                #   [#] Store the INDEX founded above in variable of 'VC103_variable_index'.         |
                                #   [#] Change the tool number by:                                                   |
                                #      -Use 'VC103_variable_index' that founded above.                               |
                                #      -Use <PandasLibrary> method to locate the tool number by Access the Excel File|
                                #       'horizontal_tool_list_file' and Sheet of name ['MISCELLANEOUS_TOOL_LIST'] and|
                                #       column of ['TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)'] in location of        |
                                #       'miscellaneous_tool_list_index' that founded above.                          |
                                # ===================================================================================|
                                VC103_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(
                                    'VC103=06')
                                pin_bore_program_lines_of_old_horizontal_machine[VC103_variable_index] = \
                                    ('VC103=' + str(horizontal_tool_list_file['MISCELLANEOUS_TOOL_LIST'].loc[
                                        miscellaneous_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']))

                        # endregion  <<<<============================[Ledge Cut Tool]============================>>>>

                        # region  <<<<============================[Rough Bore Tool]============================>>>>

                        # ===========================================================================================|
                        # [#] The RoughBoreTool is the tool used to rough the PinHole before it finished or add      |
                        #     any other PinBore Features, tool diameter will be different according to the Finished  |
                        #     PinHoleDiameter, therefore it needs logic to decide which tool can use with each size. |
                        # [#] The PinHoleDiameter could be very variant depends on the Piston Design, therefore it   |
                        #     needs to use 'Range' number check with (for loop) to be always have RoughTool no       |
                        #     matter what the Finish PinHoleDiameter size, the safe Range of the RoughTool is:       |
                        #      -Minimum PinHoleDiameter : RoughBoreTool(mm) + ~0.43(mm) -->                          |
                        #        EX: Use 11(mm) RoughTool to rough 11.43(mm) FinishPinHoleDiameter                   |
                        #      -Maximum PinHoleDiameter : RoughBoreTool(mm) + ~1.60(mm) -->                          |
                        #         EX: Use 11(mm) RoughTool to rough 12.599(mm) FinishPinHoleDiameter                 |
                        #      On other Words:                                                                       |
                        #      -Minimum PinHoleDiameter: RoughBoreTool(inch) + ~0.017(inch) -->                      |
                        #        EX: Use 0.433(inch) RoughTool to rough 0.45(inch) FinishPinHoleDiameter             |
                        #      -Maximum PinHoleDiameter : RoughBoreTool(inch) + ~0.063(inch) -->                     |
                        #         EX: Use 0.433(inch) RoughTool to rough 0.4959(inch) FinishPinHoleDiameter          |
                        # ===========================================================================================|

                        # ========================================================================================|
                        #                         [Steps to Add Description of RoughBoreTool]                     |
                        # [#] Steps to add the tool description from the Excel Sheet to ToolList in template      |
                        #   according to FinishPinHoleDiameter:                                                   |
                        #  [#] Needs to find the index of the tool in the 'rough_bore_tool_list' (that set earlier|
                        #      in [Load Horizontal Sheets Function] section) by use the DrillDiameterSize of the  |
                        #      tool that range between ('11MM' to '26MM').                                        |
                        #  [#] Store the INDEX founded above in variable of 'rough_bore_tool_list_index'.         |
                        #  [#] Add the tool description to the ToolList of the template by:                       |
                        #   -Use <insert> method to OldHorizontalTemplateLines list in current tool_list_index.   |
                        #   -Use <PandasLibrary> method to locate the tool description by Access the Excel File   |
                        #    'horizontal_tool_list_file' and Sheet of name ['ROUGH_BORE_TOOL_LIST'] and           |
                        #    column of ['DESCRIPTION(FOR_MACHINES_27/28/32)'] in location of                      |
                        #    'rough_bore_tool_list_index' that founded above.                                     |
                        #  [#] Add <1> to the tool_list_index to add the next applicable tool after this tool.    |
                        # ========================================================================================|

                        # ========================================================================================|
                        #                         [Steps to Set the Number of RoughBoreTool]                      |
                        # [#] Needs to set the ToolNumber of the RoughBoreTool to use it later on "TOOLS NUMBER"  |
                        #     section in OldHorizontalTemplate.                                                   |
                        # [#] Steps to set the ToolNumber from the Excel Sheet according to FinishPinHoleDiameter:|
                        #  [#] Use the index of 'rough_bore_tool_list_index' that founded above.                  |
                        #  [#] Define variable 'rough_bore_tool_number' to set the ToolNumber from Excel Sheet by:|
                        #   -Use <PandasLibrary> method to locate the ToolNumber by Access the Excel File         |
                        #    'horizontal_tool_list_file' and Sheet of name ['ROUGH_BORE_TOOL_LIST'] and           |
                        #    column of ['TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)'] in location of                |
                        #    'rough_bore_tool_list_index' that founded above.                                     |
                        # ========================================================================================|
                        global rough_bore_tool_number
                        if (0.4500 <= pin_hole_diameter < 0.4960):
                            rough_bore_tool_list_index = rough_bore_tool_list.index('11MM')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                    rough_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            rough_bore_tool_number = horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                rough_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.4960 <= pin_hole_diameter < 0.5355):
                            rough_bore_tool_list_index = rough_bore_tool_list.index('12MM')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                    rough_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            rough_bore_tool_number = horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                rough_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.5355 <= pin_hole_diameter < 0.5750):
                            rough_bore_tool_list_index = rough_bore_tool_list.index('13MM')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                    rough_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            rough_bore_tool_number = horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                rough_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.5750 <= pin_hole_diameter < 0.6142):
                            rough_bore_tool_list_index = rough_bore_tool_list.index('14MM')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                    rough_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            rough_bore_tool_number = horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                rough_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.6142 <= pin_hole_diameter < 0.6536):
                            rough_bore_tool_list_index = rough_bore_tool_list.index('15MM')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                    rough_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            rough_bore_tool_number = horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                rough_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.6536 <= pin_hole_diameter < 0.6930):
                            rough_bore_tool_list_index = rough_bore_tool_list.index('16MM')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                    rough_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            rough_bore_tool_number = horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                rough_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.6930 <= pin_hole_diameter < 0.7323):
                            rough_bore_tool_list_index = rough_bore_tool_list.index('17MM')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                    rough_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            rough_bore_tool_number = horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                rough_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.7323 <= pin_hole_diameter < 0.7717):
                            rough_bore_tool_list_index = rough_bore_tool_list.index('18MM')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                    rough_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            rough_bore_tool_number = horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                rough_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.7717 <= pin_hole_diameter < 0.8111):
                            rough_bore_tool_list_index = rough_bore_tool_list.index('19MM')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                    rough_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            rough_bore_tool_number = horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                rough_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.8111 <= pin_hole_diameter < 0.8504):
                            rough_bore_tool_list_index = rough_bore_tool_list.index('20MM')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                    rough_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            rough_bore_tool_number = horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                rough_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.8504 <= pin_hole_diameter < 0.8898):
                            rough_bore_tool_list_index = rough_bore_tool_list.index('21MM')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                    rough_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            rough_bore_tool_number = horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                rough_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.8898 <= pin_hole_diameter < 0.9292):
                            rough_bore_tool_list_index = rough_bore_tool_list.index('22MM')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                    rough_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            rough_bore_tool_number = horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                rough_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.9292 <= pin_hole_diameter < 0.9686):
                            rough_bore_tool_list_index = rough_bore_tool_list.index('23MM')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                    rough_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            rough_bore_tool_number = horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                rough_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.9686 <= pin_hole_diameter < 1.0079):
                            rough_bore_tool_list_index = rough_bore_tool_list.index('24MM')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                    rough_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            rough_bore_tool_number = horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                rough_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (1.0079 <= pin_hole_diameter < 1.0473):
                            rough_bore_tool_list_index = rough_bore_tool_list.index('1_INCH')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                    rough_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            rough_bore_tool_number = horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                rough_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        # ===========================================================================================|
                        # [#] Even if the FinishPinHoleDiameter BiggerThan <1.0942>, Use this RoughTool, and the User|
                        #     will be informed to use the LedgeCounterbore tool to add the needed passes manually.   |
                        # ===========================================================================================|
                        elif (1.0473 <= pin_hole_diameter < 1.0942 or pin_hole_diameter >= 1.0942):
                            rough_bore_tool_list_index = rough_bore_tool_list.index('26MM')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                    rough_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            rough_bore_tool_number = horizontal_tool_list_file['ROUGH_BORE_TOOL_LIST'].loc[
                                rough_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1

                        # ==========================================================================================|
                        # [#] If the FinishPinHoleDiameter IsEqual <0> or 'None', Just set the ToolNumber to be <0>,|
                        #     because the ToolNumber can not be nothing, should have a value, and the User will be  |
                        #     informed to fix the issue.                                                            |
                        # ==========================================================================================|
                        elif (pin_hole_diameter == 0.0 or pin_hole_diameter is None):
                            print("PIN HOLE DIAMETER CAN'T BE <0> OR NOTHING, SEE ENGINEERING")
                            rough_bore_tool_number = 0
                        # =============================================================================================|
                        # [#] If Can't find RoughTool fit the FinishPinHoleDiameter, Just set the ToolNumber to be <0>,|
                        #     because the ToolNumber can not be nothing, should have a value, and the User will be     |
                        #     informed to fix the issue.                                                               |
                        # =============================================================================================|
                        else:
                            print("Can't find RoughTool fit the FinishPinHoleDiameter, MAYBE WE CAN USE SWAP TOOL "
                                  "T60 OR JUST PUT MESSAGE SEE PROGRAMMING")
                            # global rough_bore_tool_number
                            rough_bore_tool_number = 0

                        # endregion  <<<<============================[Rough Bore Tool]============================>>>>

                        # region  <<<<============================[Finish Bore Tool]============================>>>>

                        # ==========================================================================================|
                        # [#] The FinishBoreTool is the tool used to finish the PinHole before adding               |
                        #     any other PinBore Features, tool diameter will be different according to the Finished |
                        #     PinHoleDiameter, therefore it needs logic to decide which tool can use with each size.|
                        # [#] The PinHoleDiameter could be very variant depends on the Piston Design, sometimes it  |
                        #     use MapelTool if it's available, or MapleTool(T45) for some sizes, otherwise it use   |
                        #     the BoringBarTool, therefore it's need to use 'Range' number check with (for loop)    |
                        #     to decide the FinishBoreTool, the safe Range of the FinishTool is:                    |
                        #     (PinHoleDiameter - 0.001) <= PinHoleDiameter <= (PinHoleDiameter + 0.001)             |
                        #         Ex: Use (T08 0.827 MAPAL REAMER) to finish 0.826 PinHoleDiameter in WD-14949,     |
                        #             Use (T03 0.927 MAPAL REAMER) to finish 0.928 PinHoleDiameter in WD-16057.     |
                        # ==========================================================================================|

                        # =========================================================================================|
                        #                         [Steps to Add Description of FinishBoreTool]                     |
                        # [#] Steps to add the tool description from the Excel Sheet to ToolList in template       |
                        #   according to FinishPinHoleDiameter:                                                    |
                        #  [#] Needs to find the index of the tool in the 'finish_bore_tool_list' (that set earlier|
                        #      in [Load Horizontal Sheets Function] section) by use the PinHoleDiameter of the     |
                        #      tool that will start with (0.4724) and end with ('BORING_BAR_TOOL').                |
                        #  [#] Store the INDEX founded above in variable of 'finish_bore_tool_list_index'.         |
                        #  [#] Add the tool description to the ToolList of the template by:                        |
                        #   -Use <insert> method to OldHorizontalTemplateLines list in current tool_list_index.    |
                        #   -Use <PandasLibrary> method to locate the tool description by Access the Excel File    |
                        #    'horizontal_tool_list_file' and Sheet of name ['FINISH_BORE_TOOL_LIST'] and           |
                        #    column of ['DESCRIPTION(FOR_MACHINES_27/28/32)'] in location of                       |
                        #    'finish_bore_tool_list_index' that founded above.                                     |
                        #  [#] Add <1> to the tool_list_index to add the next applicable tool after this tool.     |
                        # =========================================================================================|

                        # =========================================================================================|
                        #                         [Steps to Set the Number of FinishBoreTool]                      |
                        # [#] Needs to set the ToolNumber of the FinishBoreTool to use it later on "TOOLS NUMBER"  |
                        #     section in OldHorizontalTemplate.                                                    |
                        # [#] Steps to set the ToolNumber from the Excel Sheet according to FinishPinHoleDiameter: |
                        #  [#] Use the index of 'finish_bore_tool_list_index' that founded above.                  |
                        #  [#] Define variable 'finish_bore_tool_number' to set the ToolNumber from Excel Sheet by:|
                        #   -Use <PandasLibrary> method to locate the ToolNumber by Access the Excel File          |
                        #    'horizontal_tool_list_file' and Sheet of name ['FINISH_BORE_TOOL_LIST'] and           |
                        #    column of ['TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)'] in location of                 |
                        #    'finish_bore_tool_list_index' that founded above.                                     |
                        # =========================================================================================|

                        # NEED TO CHECK IF WE GONNA USE MAPEL TOOL OR BORING BAR
                        if (0.471 <= pin_hole_diameter <= 0.473):
                            # pin_hole_diameter = 0.472
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.4724)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        # NEED TO CHECK IF WE GONNA USE MAPEL TOOL OR BORING BAR
                        elif (0.489 <= pin_hole_diameter <= 0.491):
                            # pin_hole_diameter = 0.490
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.49)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.511 <= pin_hole_diameter <= 0.513):
                            # pin_hole_diameter = 0.512
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.512)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.550 <= pin_hole_diameter <= 0.552):
                            # pin_hole_diameter = 0.551
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.551)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.590 <= pin_hole_diameter <= 0.592):
                            # pin_hole_diameter = 0.591
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.591)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.629 <= pin_hole_diameter <= 0.631):
                            # pin_hole_diameter = 0.630
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.63)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.668 <= pin_hole_diameter <= 0.670):
                            # pin_hole_diameter = 0.669
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.669)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.671 <= pin_hole_diameter <= 0.673):
                            # pin_hole_diameter = 0.672
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.672)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.708 <= pin_hole_diameter <= 0.710):
                            # pin_hole_diameter = 0.709
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.709)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.747 <= pin_hole_diameter <= 0.749):
                            # pin_hole_diameter = 0.748
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.748)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.786 <= pin_hole_diameter <= 0.788):
                            # pin_hole_diameter = 0.787
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.787)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        # =========================================================================|
                        # [#] Use Bigger Range because it use the same tool for 0.791 & 0.792 size.|
                        # =========================================================================|
                        elif (0.790 <= pin_hole_diameter <= 0.793):
                            # pin_hole_diameter = 0.792
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.792)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                            # NEED TO CHECK IF WE CAN USE ROUGH TOOL OF 20MM INSTEAD 19MM AS OLD PROGRAMS DONE
                        elif (0.811 <= pin_hole_diameter <= 0.813):
                            # pin_hole_diameter = 0.812
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.8124)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.826 <= pin_hole_diameter <= 0.828):
                            # pin_hole_diameter = 0.827
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.827)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.865 <= pin_hole_diameter <= 0.867):
                            # pin_hole_diameter = 0.866
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.866)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.904 <= pin_hole_diameter <= 0.906):
                            # pin_hole_diameter = 0.905
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.905)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.911 <= pin_hole_diameter <= 0.913):
                            # pin_hole_diameter = 0.912
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.912)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.926 <= pin_hole_diameter <= 0.928):
                            # pin_hole_diameter = 0.927
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.927)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1

                            # ******************MAYBE NEED TO ADD 0.943 AS MAPLE USE T45**********************
                            # MAYBE DESCRIPTION NEEDS TO BE (...... MAPAL REAMER - RUN AS BORING BAR),
                            # MAYBE SOMETIMES ADD (.........  MAPAL REAMER - NO TRAILING INSERT)

                        elif (0.944 <= pin_hole_diameter <= 0.946):
                            # pin_hole_diameter = 0.945
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.945)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.974 <= pin_hole_diameter <= 0.976):
                            # pin_hole_diameter = 0.975
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.975)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.979 <= pin_hole_diameter <= 0.981):
                            # pin_hole_diameter = 0.980
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.98)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.983 <= pin_hole_diameter <= 0.985):
                            # pin_hole_diameter = 0.984
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.984)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.989 <= pin_hole_diameter <= 0.991):
                            # pin_hole_diameter = 0.990
                            finish_bore_tool_list_index = finish_bore_tool_list.index(0.99)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1
                        elif (0.999 <= pin_hole_diameter <= 1.001):
                            # pin_hole_diameter = 1.000
                            finish_bore_tool_list_index = finish_bore_tool_list.index(1.00)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1

                            # MAYBE NEED TO ADD ANOTHER LOGIC TO USE 1.094 MAPEL TOOL IF NEED IT

                        elif (1.093 <= pin_hole_diameter <= 1.095):
                            # pin_hole_diameter = 1.094
                            finish_bore_tool_list_index = finish_bore_tool_list.index('1.094_BORING')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1

                        # ==========================================================================================|
                        # [#] If the FinishPinHoleDiameter IsEqual <0> or 'None', Just set the ToolNumber to be <0>,|
                        #     because the ToolNumber can not be nothing, should have a value, and the User will be  |
                        #     informed to fix the issue.                                                            |
                        # ==========================================================================================|
                        elif (pin_hole_diameter == 0.0 or pin_hole_diameter is None):
                            print("PIN HOLE DIAMETER CAN'T BE <0> OR NOTHING, SEE ENGINEERING")
                            finish_bore_tool_number = 0

                        # ===========================================================================================|
                        # [#] If there is NO MapelTool founded for the FinishPinHoleDiameter size, Use BoringBar tool|
                        #     with the given PinHoleDiameter size.                                                   |
                        # ===========================================================================================|
                        else:
                            # pin_hole_diameter = float(str(pin_hole_diameter)[:5])
                            finish_bore_tool_list_index = finish_bore_tool_list.index('BORING_BAR_TOOL')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                    finish_bore_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'] + format(
                                    pin_hole_diameter) + ' BORING BAR)')
                            finish_bore_tool_number = horizontal_tool_list_file['FINISH_BORE_TOOL_LIST'].loc[
                                finish_bore_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1

                        # =============================================================================================|
                        # [#] Needs to check which PinHoleDiameter Sizes will use Swap MAPEL Tool T45 to add them here |
                        #     and maybe on Excel sheet also.                                                           |
                        # =============================================================================================|
                        # =============================================================================================|
                        # NEEDS TO CHECK IF WE NEED TO ADD NOTE OF (WILL HONE TO .928 FOR EXAMPLE)
                        # =============================================================================================|

                        # endregion  <<<<============================[Finish Bore Tool]============================>>>>

                        # region  <<<<============================[LockRing/C-Fren Tool]============================>>>>

                        # =========================================================================================|
                        # [#] The LockRing/C-Fren Tool will be different according to the width of LockRing/C-Fren,|
                        #     therefore it needs logic to choose the tool with the same width.                     |
                        # =========================================================================================|

                        # =========================================================================================|
                        #                     [Steps to Add Description of LockRing/C-Fren Tool]                   |
                        # [#] Steps to add the tool description from the Excel Sheet to ToolList in template       |
                        #     according to LockRing/C-Fren Width:                                                  |
                        #  [#] Needs to find the index of the tool in the 'lock_ring_and_cfren_tool_list' (that set|
                        #      earlier in [Load Horizontal Sheets Function] section) by use the TOOL_WIDTH of the  |
                        #      tool that range between ('0.039' to '0.11').                                        |
                        #  [#] Store the INDEX founded above in variable of 'lock_ring_and_cfren_tool_list_index'. |
                        #  [#] Add the tool description to the ToolList of the template by:                        |
                        #   -Use <insert> method to OldHorizontalTemplateLines list in current tool_list_index.    |
                        #   -Use <PandasLibrary> method to locate the tool description by Access the Excel File    |
                        #    'horizontal_tool_list_file' and Sheet of name ['LOCK_RING_AND_CFREN_TOOL_LIST'] and   |
                        #    column of ['DESCRIPTION(FOR_MACHINES_27/28/32)'] in location of                       |
                        #    'lock_ring_and_cfren_tool_list_index' that founded above.                             |
                        #  [#] Add <1> to the tool_list_index to add the next applicable tool after this tool.     |
                        # =========================================================================================|

                        # ========================================================================================|
                        #                      [Steps to Set the Number of LockRing/C-Fren Tool]                  |
                        # [#] Needs to set the ToolNumber of the LockRing/C-Fren Tool to use it later on          |
                        #     "TOOLS NUMBER" section in OldHorizontalTemplate.                                    |
                        # [#] Steps to set the ToolNumber from the Excel Sheet according to LockRing/C-Fren Width:|
                        #  [#] Use the index of 'lock_ring_and_cfren_tool_list_index' that founded above.         |
                        #  [#] Define variable 'lock_ring_tool_number' to set the ToolNumber from Excel Sheet by: |
                        #   -Use <PandasLibrary> method to locate the ToolNumber by Access the Excel File         |
                        #    'horizontal_tool_list_file' and Sheet of name ['LOCK_RING_AND_CFREN_TOOL_LIST'] and  |
                        #    column of ['TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)'] in location of                |
                        #    'lock_ring_and_cfren_tool_list_index' that founded above.                            |
                        # ========================================================================================|

                        # ===========================================================================================|
                        #                      [Steps to Set the Diameter of LockRing/C-Fren Tool]                   |
                        # [#] Needs to set the Diameter of the LockRing/C-Fren Tool to use it later on               |
                        #     the OldHorizontalTemplate.                                                             |
                        # [#] Steps to set the ToolDiameter from the Excel Sheet according to LockRing/C-Fren Width: |
                        #  [#] Use the index of 'lock_ring_and_cfren_tool_list_index' that founded above.            |
                        #  [#] Define variable 'lock_ring_tool_diameter' to set the ToolDiameter from Excel Sheet by:|
                        #   -Use <PandasLibrary> method to locate the ToolDiameter by Access the Excel File          |
                        #    'horizontal_tool_list_file' and Sheet of name ['LOCK_RING_AND_CFREN_TOOL_LIST'] and     |
                        #    column of ['TOOL_DIAMETER'] in location of 'lock_ring_and_cfren_tool_list_index'        |
                        #    that founded above.                                                                     |
                        # ===========================================================================================|

                        # NEED LOGIC TO ADD BOTH LOCKRING AND CFREN TOOL DESCRIPTION TO THE TOOL LIST IN THE PROGRAM IF
                        # THEY ARE NOT THE SAME TOOL EX WD-13100
                        # NEED TO ADJUST HORIZONTAL TEMPLATE ALSO

                        # ============================================================================================|
                        # [#] If LockRing/C-Fren Width is <0> or 'None', that's mean job doesn't have LockRing/C-Fren,|
                        #     and no need to add anything.                                                            |
                        # [#] Set the ToolNumber to be <0>, because the ToolNumber can not be nothing, should         |
                        #     have a value.                                                                           |
                        # ============================================================================================|
                        if ((lock_ring_cutter_width == 0.0 or lock_ring_cutter_width is None) and
                                (cfren_cutter_width == 0.0 or cfren_cutter_width is None)):
                            pass
                            lock_ring_tool_number = 0
                        # ===========================================================================================|
                        # [#] Needs to separate logic of LockRing/C-Fren for the next 'elif Statement', because needs|
                        #     to use the ToolWidth in the description (T43 is Swap tool use for uncommon ToolWidth). |
                        # ===========================================================================================|
                        elif (lock_ring_cutter_width == 0.039):
                            lock_ring_and_cfren_tool_list_index = lock_ring_and_cfren_tool_list.index(0.039)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                    lock_ring_and_cfren_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'] + format(
                                    lock_ring_cutter_width) + ' SQ X .465 DIA. LOCK RING)')
                            lock_ring_tool_number = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            lock_ring_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            tool_list_index += 1
                        elif (cfren_cutter_width == 0.039):
                            lock_ring_and_cfren_tool_list_index = lock_ring_and_cfren_tool_list.index(0.039)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                    lock_ring_and_cfren_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'] + format(
                                    cfren_cutter_width) + ' SQ X .465 DIA. LOCK RING)')
                            lock_ring_tool_number = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            cfren_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            tool_list_index += 1

                            # NEED TO CHECK WITCH TOOL WITH 0.042 WIDTH WILL USE WHILE WE HAVE THREE TOOLS WITH
                            # DIFFERENT DIAMETER (OR ADD LOGIC TO CHECK WHICH TOOL WILL USE),
                            # NEED TO ADD ANOTHER TOOL TO EXCEL SHEET AS WELL
                        elif (lock_ring_cutter_width == 0.042):
                            lock_ring_and_cfren_tool_list_index = lock_ring_and_cfren_tool_list.index(0.042)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                    lock_ring_and_cfren_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'] + format(
                                    lock_ring_cutter_width) + ' SQ X .465 PH HORN LOCK RING)')
                            lock_ring_tool_number = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            lock_ring_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            tool_list_index += 1
                        elif (cfren_cutter_width == 0.042):
                            lock_ring_and_cfren_tool_list_index = lock_ring_and_cfren_tool_list.index(0.042)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                    lock_ring_and_cfren_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'] + format(
                                    cfren_cutter_width) + ' SQ X .465 PH HORN LOCK RING)')
                            lock_ring_tool_number = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            cfren_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            tool_list_index += 1

                        # NEED TO CHECK WITCH TOOL WITH 0.044 WIDTH WILL USE WHILE WE HAVE THREE TOOLS WITH DIFFERENT
                        # DIAMETER (OR ADD LOGIC TO CHECK WITCH TOOL WILL USE),
                        # FOR NOW WE WILL USE (T25 IS A .044 RAD X .465 PH HORN LOCK RING)
                        elif (lock_ring_cutter_width == 0.044 or cfren_cutter_width == 0.044):
                            lock_ring_and_cfren_tool_list_index = lock_ring_and_cfren_tool_list.index(
                                '0.044(DIA 0.465)')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                    lock_ring_and_cfren_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            lock_ring_tool_number = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            lock_ring_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            cfren_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            tool_list_index += 1

                            # MAYBE NEEDS TO ADD LOCKRING TOOL WITH 0.046 WIDTH

                        elif (lock_ring_cutter_width == 0.047 or cfren_cutter_width == 0.047):
                            lock_ring_and_cfren_tool_list_index = lock_ring_and_cfren_tool_list.index(0.047)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                    lock_ring_and_cfren_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            lock_ring_tool_number = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            lock_ring_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            cfren_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            tool_list_index += 1

                        # ===========================================================================================|
                        # [#] Needs to separate logic of LockRing/C-Fren for the next 'elif Statement', because needs|
                        #     to use the ToolWidth in the description (T43 is Swap tool use for uncommon ToolWidth). |
                        # ===========================================================================================|
                        elif (lock_ring_cutter_width == 0.048):
                            lock_ring_and_cfren_tool_list_index = lock_ring_and_cfren_tool_list.index(0.048)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                    lock_ring_and_cfren_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'] + format(
                                    lock_ring_cutter_width) + ' RAD X .575 DIA. LOCK RING)')
                            lock_ring_tool_number = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            lock_ring_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            tool_list_index += 1
                        elif (cfren_cutter_width == 0.048):
                            lock_ring_and_cfren_tool_list_index = lock_ring_and_cfren_tool_list.index(0.048)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                    lock_ring_and_cfren_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'] + format(
                                    cfren_cutter_width) + ' RAD X .575 DIA. LOCK RING)')
                            lock_ring_tool_number = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            cfren_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            tool_list_index += 1

                        # NEED TO CHECK WITCH TOOL WITH 0.053 WIDTH WILL USE WHILE WE HAVE THREE TOOLS WITH DIFFERENT
                        # DIAMETER (OR ADD LOGIC TO CHECK WICTH TOOL WILL USE),
                        # FOR NOW WE WILL USE (T09 IS A .053 RAD X .618 PH HORN LOCK RING)
                        elif (lock_ring_cutter_width == 0.053 or cfren_cutter_width == 0.053):
                            lock_ring_and_cfren_tool_list_index = lock_ring_and_cfren_tool_list.index(
                                '0.053(DIA 0.618)')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                    lock_ring_and_cfren_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            lock_ring_tool_number = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            lock_ring_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            cfren_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            tool_list_index += 1
                        elif (lock_ring_cutter_width == 0.059 or cfren_cutter_width == 0.059):
                            lock_ring_and_cfren_tool_list_index = lock_ring_and_cfren_tool_list.index(0.059)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                    lock_ring_and_cfren_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            lock_ring_tool_number = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            lock_ring_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            cfren_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            tool_list_index += 1
                        elif (lock_ring_cutter_width == 0.063 or cfren_cutter_width == 0.063):
                            lock_ring_and_cfren_tool_list_index = lock_ring_and_cfren_tool_list.index(0.063)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                    lock_ring_and_cfren_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            lock_ring_tool_number = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            lock_ring_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            cfren_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            tool_list_index += 1

                        # MAYBE NEEDS TO ADD LOCKRING TOOL WITH 0.064 WIDTH

                        elif (lock_ring_cutter_width == 0.065 or cfren_cutter_width == 0.065):
                            lock_ring_and_cfren_tool_list_index = lock_ring_and_cfren_tool_list.index(0.065)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                    lock_ring_and_cfren_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            lock_ring_tool_number = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            lock_ring_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            cfren_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            tool_list_index += 1
                        elif (lock_ring_cutter_width == 0.067 or cfren_cutter_width == 0.067):
                            lock_ring_and_cfren_tool_list_index = lock_ring_and_cfren_tool_list.index(0.067)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                    lock_ring_and_cfren_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            lock_ring_tool_number = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            lock_ring_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            cfren_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            tool_list_index += 1

                        # MAYBE NEEDS TO ADD LOCKRING TOOL WITH 0.071 WIDTH

                        # MAYBE NEEDS TO ADD LOCKRING TOOL WITH 0.073 WIDTH

                        elif (lock_ring_cutter_width == 0.076 or cfren_cutter_width == 0.076):
                            lock_ring_and_cfren_tool_list_index = lock_ring_and_cfren_tool_list.index(0.076)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                    lock_ring_and_cfren_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            lock_ring_tool_number = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            lock_ring_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            cfren_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            tool_list_index += 1
                        elif (lock_ring_cutter_width == 0.077 or cfren_cutter_width == 0.077):
                            lock_ring_and_cfren_tool_list_index = lock_ring_and_cfren_tool_list.index(0.077)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                    lock_ring_and_cfren_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            lock_ring_tool_number = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            lock_ring_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            cfren_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            tool_list_index += 1
                        elif (lock_ring_cutter_width == 0.088 or cfren_cutter_width == 0.088):
                            lock_ring_and_cfren_tool_list_index = lock_ring_and_cfren_tool_list.index(0.088)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                    lock_ring_and_cfren_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            lock_ring_tool_number = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            lock_ring_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            cfren_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            tool_list_index += 1
                        elif (lock_ring_cutter_width == 0.11 or cfren_cutter_width == 0.11):
                            lock_ring_and_cfren_tool_list_index = lock_ring_and_cfren_tool_list.index(0.11)
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                    lock_ring_and_cfren_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            lock_ring_tool_number = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            lock_ring_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            cfren_tool_diameter = horizontal_tool_list_file['LOCK_RING_AND_CFREN_TOOL_LIST'].loc[
                                lock_ring_and_cfren_tool_list_index, 'TOOL_DIAMETER']
                            tool_list_index += 1

                        # NEED TO BACK AND WORK HERE
                        # NEED TO CHECK IF WE CAN USE T43 WITH ANY TOOL WIDTH COMES WITH JOB(LIKE WHAT WE DO FOR
                        # BORING BAR), OR NEED TO PUT MESSAGE ON TOP OF PROGRAM OR ON DIALOG OF APP TO QUESTION
                        # THE TOOL AVAILABILITY

                        else:
                            print("LOCK RING TOOL IS NOT ON THE TOOL LIST, SEE PROGRAMMING")
                            lock_ring_tool_number = 0

                        # endregion  <<<<==========================[LockRing/C-Fren Tool]==========================>>>>

                        # region  <<<<========================[Double Oil Hole Slot Tool]========================>>>>

                        # =============================================================================================|
                        # [#] Define Variable to set DOHS Availability Status.                                         |
                        # [#] Set the Value to be <0> as a default to use it in [VC123] Variable in HorizontalTemplate.|
                        # [#] Change the Value to be <1> when DOHS ID Spacing Detected, and PinHoleDiameter is         |
                        #     GreaterThanOrEqual <0.901>.                                                              |
                        # =============================================================================================|

                        # =====================================================================================|
                        #                   [Steps to Add Description of DoubleOilHoleSlotTool]                |
                        # [#] Steps to add the tool description from the Excel Sheet to ToolList in template:  |
                        #  [#] Needs to find the index of the tool in the 'miscellaneous_tool_list' (that set  |
                        #      earlier in [Load Horizontal Sheets Function] section) by use the                |
                        #      ToolUsage of the tool which is ('DOUBLE OIL HOLES SLOTS(DOHS) 0.750 PH').       |
                        #  [#] Store the INDEX founded above in variable of 'miscellaneous_tool_list_index'.   |
                        #  [#] Add the tool description to the ToolList of the template by:                    |
                        #   -Use <insert> method to OldHorizontalTemplateLines list in current tool_list_index.|
                        #   -Use <PandasLibrary> method to locate the tool description by Access the Excel File|
                        #    'horizontal_tool_list_file' and Sheet of name ['MISCELLANEOUS_TOOL_LIST'] and     |
                        #    column of ['DESCRIPTION(FOR_MACHINES_27/28/32)'] in location of                   |
                        #    'miscellaneous_tool_list_index' that founded above.                               |
                        #  [#] Add <1> to the tool_list_index to add the next applicable tool after this tool. |
                        # =====================================================================================|

                        # ==================================================================================|
                        #                    [Steps to Set the Number of DoubleOilHoleSlotTool]             |
                        # [#] Needs to set the ToolNumber of the DoubleOilHoleSlotTool to use it later on   |
                        #     "TOOLS NUMBER" section in OldHorizontalTemplate.                              |
                        # [#] Steps to set the ToolNumber from the Excel Sheet:                             |
                        #  [#] Use the index of 'miscellaneous_tool_list_index' that founded above.         |
                        #  [#] Define variable 'double_oil_hole_slot_tool_number' to set the ToolNumber from|
                        #      Excel Sheet by:                                                              |
                        #     -Use <PandasLibrary> method to locate the ToolNumber by Access the Excel File |
                        #     'horizontal_tool_list_file' and Sheet of name ['MISCELLANEOUS_TOOL_LIST'] and |
                        #     column of ['TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)'] in location of         |
                        #     'miscellaneous_tool_list_index' that founded above.                           |
                        # ==================================================================================|

                        global double_oil_hole_slot_availability_status
                        double_oil_hole_slot_availability_status = 0
                        if (double_oil_hole_slot_ID_spacing != 0 and double_oil_hole_slot_ID_spacing is not None and
                                pin_hole_diameter >= 0.901):
                            double_oil_hole_slot_availability_status = 1
                            miscellaneous_tool_list_index = miscellaneous_tool_list.index(
                                'DOUBLE OIL HOLES SLOTS(DOHS) 0.750 PH')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['MISCELLANEOUS_TOOL_LIST'].loc[
                                    miscellaneous_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            double_oil_hole_slot_tool_number = horizontal_tool_list_file['MISCELLANEOUS_TOOL_LIST'].loc[
                                miscellaneous_tool_list_index, 'TOOL_NUMBER(FOR_MACHINES_27/28/32) (T00)']
                            tool_list_index += 1

                        # endregion  <<<<=======================[Double Oil Hole Slot Tool]=======================>>>>

                        # region  <<<<========================[Notch Tool]========================>>>>

                        # =============================================================================================|
                        # [#] Define Variable to set Notch Availability Status.                                        |
                        # [#] Set the Value to be <0> as a default to use it in [VC124] Variable in HorizontalTemplate.|
                        # [#] Change the Value to be <1> when Notch Angle First Location Detected.                     |
                        # =============================================================================================|

                        # =====================================================================================|
                        #                        [Steps to Add Description of NotchTool]                       |
                        # [#] Steps to add the tool description from the Excel Sheet to ToolList in template:  |
                        #  [#] Needs to find the index of the tool in the 'miscellaneous_tool_list' (that set  |
                        #      earlier in [Load Horizontal Sheets Function] section) by use the                |
                        #      ToolUsage of the tool which is ('NOTCH TOOL 5/32 DIA').                         |
                        #  [#] Store the INDEX founded above in variable of 'miscellaneous_tool_list_index'.   |
                        #  [#] Add the tool description to the ToolList of the template by:                    |
                        #   -Use <insert> method to OldHorizontalTemplateLines list in current tool_list_index.|
                        #   -Use <PandasLibrary> method to locate the tool description by Access the Excel File|
                        #    'horizontal_tool_list_file' and Sheet of name ['MISCELLANEOUS_TOOL_LIST'] and     |
                        #    column of ['DESCRIPTION(FOR_MACHINES_27/28/32)'] in location of                   |
                        #    'miscellaneous_tool_list_index' that founded above.                               |
                        #  [#] Add <1> to the tool_list_index to add the next applicable tool after this tool. |
                        # =====================================================================================|

                        global notch_availability_status
                        notch_availability_status = 0
                        if ((notch_angle_first_location != 0 and notch_angle_first_location is not None) or
                                (notch_angle_second_location != 0 and notch_angle_second_location is not None)):
                            notch_availability_status = 1
                            miscellaneous_tool_list_index = miscellaneous_tool_list.index('NOTCH TOOL 5/32 DIA')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['MISCELLANEOUS_TOOL_LIST'].loc[
                                    miscellaneous_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            tool_list_index += 1

                        # endregion  <<<<========================[Notch Tool]========================>>>>

                        # region  <<<<========================[Horizontal Pin Slot Tool]========================>>>>

                        # =============================================================================================|
                        # [#] Define Variable to set Horizontal Pin Slot Availability Status.                          |
                        # [#] Set the Value to be <0> as a default to use it in [VC125] Variable in HorizontalTemplate.|
                        # [#] Change the Value to be <1> when horizontal_slots_arc_diameter GreaterThan <0.375>.       |
                        #       or if it just detected (need to check that)                                            |
                        # =============================================================================================|

                        # =====================================================================================|
                        #                        [Steps to Add Description of H-Slot Tool]                     |
                        # [#] Steps to add the tool description from the Excel Sheet to ToolList in template:  |
                        #  [#] Needs to find the index of the tool in the 'miscellaneous_tool_list' (that set  |
                        #      earlier in [Load Horizontal Sheets Function] section) by use the                |
                        #      ToolUsage of the tool which is ('HORIZONTAL SLOTS TOOL 0.375 DIA').             |
                        #  [#] Store the INDEX founded above in variable of 'miscellaneous_tool_list_index'.   |
                        #  [#] Add the tool description to the ToolList of the template by:                    |
                        #   -Use <insert> method to OldHorizontalTemplateLines list in current tool_list_index.|
                        #   -Use <PandasLibrary> method to locate the tool description by Access the Excel File|
                        #    'horizontal_tool_list_file' and Sheet of name ['MISCELLANEOUS_TOOL_LIST'] and     |
                        #    column of ['DESCRIPTION(FOR_MACHINES_27/28/32)'] in location of                   |
                        #    'miscellaneous_tool_list_index' that founded above.                               |
                        #  [#] Add <1> to the tool_list_index to add the next applicable tool after this tool. |
                        # =====================================================================================|

                        global horizontal_slots_availability_status
                        horizontal_slots_availability_status = 0

                        global horizontal_slots_straight_through_availability_status
                        horizontal_slots_straight_through_availability_status = 0

                        if ((horizontal_slots_diameter_depth != 0 and horizontal_slots_diameter_depth is not None)):
                            miscellaneous_tool_list_index = miscellaneous_tool_list.index(
                                'HORIZONTAL SLOTS TOOL 0.375 DIA')
                            pin_bore_program_lines_of_old_horizontal_machine.insert(
                                tool_list_index, horizontal_tool_list_file['MISCELLANEOUS_TOOL_LIST'].loc[
                                    miscellaneous_tool_list_index, 'DESCRIPTION(FOR_MACHINES_27/28/32)'])
                            tool_list_index += 1

                            # MAYBE WE NEED TO FIX THIS CONDITION LATER
                            # WE STILL NEED TO BACK HERE AND WORK ON CONDITION
                            if (horizontal_slots_arc_diameter != 0.375):
                                horizontal_slots_availability_status = 1
                            elif (horizontal_slots_arc_diameter == 0.375):
                                horizontal_slots_straight_through_availability_status = 1

                            # =======================================================================================|
                            # [#] Check if the Horizontal Slots OD Spacing didn't detect or found (Still equal "") to|
                            #     ask the user to enter the value.                                                   |
                            # =======================================================================================|
                            if (horizontal_slots_OD_spacing == ""):

                                # =========================================================|
                                # [#] Set verification_message to be the current Situation.|
                                # =========================================================|
                                verification_messages_of_creating_old_horizontal_machine_program = [
                                    "[u]Horizontal OD Spacing[/u] does NOT found, Please Enter the value"]

                                # ==============================================================|
                                # [#] Add verification_message to the Confirmation Details list.|
                                # ==============================================================|
                                old_horizontal_program_confirmation_email_message_list.append(
                                    "# It was need confirmation of :" + "\n" + (
                                        "**Horizontal OD Spacing** does NOT found, Please Enter the value"))

                                # ==================================================================|
                                # [#] title: set it to have the header and the verification_message.|
                                # ==================================================================|
                                self.title = '[color=0066ff]Confirmation Message[/color]' + '\n' + '\n' + \
                                             '[b][i][color=ffffff]' + \
                                             verification_messages_of_creating_old_horizontal_machine_program[0] + \
                                             '[/color][/i][/b]'

                                # ====================================================================|
                                # [#] Create the window_message that contain the Confirmation Details.|
                                # ====================================================================|
                                self.Dialog_BoxLayout = BoxLayout(height=30)

                                # ===========================================================|
                                # [#] Create the TextInput field to let user enter the value.|
                                # [#] (input_filter="float"): to accept Numeric input Only.  |
                                # ===========================================================|
                                self.horiz_slots_confirmation_text_field = TextInput(
                                    hint_text="Enter Value", multiline=False, input_filter="float",
                                    background_color=[60 / 255.0, 60 / 255.0, 60 / 255.0, 1],
                                    foreground_color=[1, 1, 1, 1], size_hint_y=None, height=30)

                                # ===========================================================================|
                                # [#] Add the TextInput field created above to be part of the window_message.|
                                # ===========================================================================|
                                self.Dialog_BoxLayout.add_widget(self.horiz_slots_confirmation_text_field)

                                # ===================================================================================|
                                # [#] Call Function of (need_confirmation) by pass the parameters that created:      |
                                #   [#] title: self.title                                                            |
                                #   [#] sub_function: self.enter_horizontal_slots_OD_spacing_value (it's the Function|
                                #       created below to set the Horizontal Slots OD Spacing).                       |
                                #   [#] dialog_type: "custom" (it used to indicate that user needs to enter value).  |
                                #   [#] content: self.Dialog_BoxLayout (the window_message that have the Details).   |
                                # ===================================================================================|
                                need_confirmation_to_create_old_horizontal_machine_program(
                                    self, self.title, self.enter_horizontal_slots_OD_spacing_value, "custom",
                                    self.Dialog_BoxLayout)
                                print("Horizontal OD Spacing BEFORE User Input: ", horizontal_slots_OD_spacing)
                                return

                            # MAYBE WE DON'T WANT THIS SECTION IF WE MAKE HORIZ SLOT STANDER WITH BORE SIZE
                        # +++++------------********************--------------------------------------+++++#
                        # +++++--LOGIC TO SET HORIZONTAL SLOT NUMBERS ACCORDING TO PIN_HOLE_DIAMETER TO USE THEM LATER
                        # IN VARIABLES(VC176,VC177, AND VC178) SECTIONS-------+++++#
                        # +++++-------------------------***************--------------------------------------+++++#
                        # MAYBE THE WAY WE PROGRAM HORIZONTAL SLOTS WILL CHANGE SOON , BUT FOR NOW WE WILL USE THE WAY
                        # WE HAVE
                        # SOME OF THESE NUMBERS ARE MISSING, NEED TO BACK AND FIGURE OUT THESE NUMBERS OR MAKE
                        # MESSAGE LIKE: CAN'T FIND THEM OR SOMETHING SIMILAR,
                        # ALSO NEED A LOGIC TO CHECK WHEN WE HAVE TO MAKE MANUAL HORIZ SLOT PROGRAM BY MASTERCAM AND
                        # MAKE THESE NUMBER 1,1,1

                        # region  <<<<========================[Horizontal Pin Slot Numbers]========================>>>>

                        # ==========================================================================================|
                        # [#] Horizontal Slot Numbers could be stander numbers according to PinHoleDiameter size or,|
                        #     Set them to equal <1> to indicate it will done manually.                              |
                        # [#] Declare the variables as 'global' again to be able to use them on Entire code.        |
                        # ==========================================================================================|
                        global i_start_horizontal_slot
                        global j_start_horizontal_slot
                        global horizontal_slot_radius

                        # =============================================================================================|
                        # Just For Now                                                                                 |
                        # Cases that needed manual H-Slots.                                                            |
                        # [#] Set Horizontal Slot Numbers to equal <1> if:                                             |
                        #  -Pressure Fed Holes intersect the H-Slot.                                                   |
                        #  -Horizontal Slots OD Spacing GreaterThanOrEqual LockRing ID Spacing (The new H-Slot Stander)|
                        #  -Job has NO LockRing and H-Slots Diameter Depth equal <0.015> (The new H-Slot Stander)      |
                        # =============================================================================================|

                        if (pressure_fed_holes_availability_status == 1):
                            i_start_horizontal_slot = 1
                            j_start_horizontal_slot = 1
                            horizontal_slot_radius = 1

                        elif ((horizontal_slots_OD_spacing is not None and horizontal_slots_OD_spacing != "" and
                               horizontal_slots_OD_spacing != 0) and
                              (lock_ring_ID_spacing is not None and
                               lock_ring_ID_spacing != "" and
                               lock_ring_ID_spacing != 0) and
                              (horizontal_slots_OD_spacing >= lock_ring_ID_spacing)):
                            i_start_horizontal_slot = 1
                            j_start_horizontal_slot = 1
                            horizontal_slot_radius = 1

                        elif ((horizontal_slots_OD_spacing is not None and horizontal_slots_OD_spacing != "" and
                               horizontal_slots_OD_spacing != 0) and
                              (lock_ring_ID_spacing is None or
                               lock_ring_ID_spacing == "" or
                               lock_ring_ID_spacing == 0) and
                              (lock_ring_cutter_width is None or lock_ring_cutter_width == "" or
                               lock_ring_cutter_width == 0) and (horizontal_slots_diameter_depth == 0.015)):
                            i_start_horizontal_slot = 1
                            j_start_horizontal_slot = 1
                            horizontal_slot_radius = 1

                        # ==========================================================================================|
                        # [#] If none of the cases above apply, Set the Horizontal Slot Numbers from the Excel Sheet|
                        #     according to FinishPinHoleDiameter, therefore it's need to use 'Range' number check   |
                        #     to decide the H-Slots Numbers.                                                        |
                        # ==========================================================================================|

                        # ==========================================================================================|
                        #                             [Steps to Set H-Slot Numbers]                                 |
                        # [#] Steps to set the H-Slot Numbers from the Excel Sheet to use them in HorizontalTemplate|
                        #   according to FinishPinHoleDiameter:                                                     |
                        #  [#] Needs to find the index of the H-Slot Numbers in the 'horizontal_slot_numbers' list  |
                        #      (that set earlier in [Load Horizontal Sheets Function] section) by use the           |
                        #      PinHoleDiameter that will start with (0.4724) and end with (1.094).                  |
                        #  [#] Store the INDEX founded above in variable of 'horizontal_slot_numbers_index'.        |
                        #  [#] Get the H-Slot Numbers by:                                                           |
                        #   -Use <PandasLibrary> method to locate the H-Slot Number by Access the Excel File        |
                        #    'horizontal_tool_list_file' and Sheet of name ['HORIZONTAL_SLOT_NUMBERS'] and          |
                        #    columns of ['i_START_HORIZONTAL_SLOT'], ['j_START_HORIZONTAL_SLOT'] and                |
                        #    ['HORIZONTAL_SLOT_RADIUS'] in location of 'horizontal_slot_numbers_index' that         |
                        #    founded above.                                                                         |
                        # ==========================================================================================|

                        elif (0.471 <= pin_hole_diameter <= 0.473):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.4724)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.489 <= pin_hole_diameter <= 0.491):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.49)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.511 <= pin_hole_diameter <= 0.513):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.512)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.550 <= pin_hole_diameter <= 0.552):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.551)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.590 <= pin_hole_diameter <= 0.592):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.591)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.629 <= pin_hole_diameter <= 0.631):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.63)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.668 <= pin_hole_diameter <= 0.670):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.669)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.671 <= pin_hole_diameter <= 0.673):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.672)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.708 <= pin_hole_diameter <= 0.710):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.709)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.727 <= pin_hole_diameter <= 0.729):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.7283)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.747 <= pin_hole_diameter <= 0.749):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.748)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.786 <= pin_hole_diameter <= 0.788):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.787)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        # ============================================================================|
                        # [#] Use Bigger Range because it use the same Numbers for 0.791 & 0.792 size.|
                        # ============================================================================|
                        elif (0.790 <= pin_hole_diameter <= 0.793):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.791)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.799 <= pin_hole_diameter <= 0.801):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.800)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.811 <= pin_hole_diameter <= 0.813):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.8124)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.826 <= pin_hole_diameter <= 0.828):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.827)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.865 <= pin_hole_diameter <= 0.867):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.866)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.874 <= pin_hole_diameter <= 0.876):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.875)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.900 <= pin_hole_diameter <= 0.902):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.901)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.904 <= pin_hole_diameter <= 0.906):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.905)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.911 <= pin_hole_diameter <= 0.913):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.912)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.926 <= pin_hole_diameter <= 0.928):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.927)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.939 <= pin_hole_diameter <= 0.941):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.94)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.942 <= pin_hole_diameter < 0.944):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.943)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.944 <= pin_hole_diameter <= 0.946):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.945)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.974 <= pin_hole_diameter <= 0.976):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.975)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.979 <= pin_hole_diameter <= 0.981):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.98)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.9811 <= pin_hole_diameter < 0.983):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.9825)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.983 <= pin_hole_diameter <= 0.985):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.984)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.989 <= pin_hole_diameter <= 0.991):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(0.99)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (0.999 <= pin_hole_diameter <= 1.001):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(1.00)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (1.030 <= pin_hole_diameter <= 1.032):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(1.031)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']
                        elif (1.093 <= pin_hole_diameter <= 1.095):
                            horizontal_slot_numbers_index = horizontal_slot_numbers.index(1.094)
                            i_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'i_START_HORIZONTAL_SLOT']
                            j_start_horizontal_slot = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'j_START_HORIZONTAL_SLOT']
                            horizontal_slot_radius = horizontal_tool_list_file['HORIZONTAL_SLOT_NUMBERS'].loc[
                                horizontal_slot_numbers_index, 'HORIZONTAL_SLOT_RADIUS']

                        # ============================================================================================|
                        # [#] If nothing of above apply, that's mean H-Slot Numbers of certain PinHole Size not exist |
                        #     in the Excel Sheet, therefore set the H-Slots Numbers to be "" and the user will fix it.|
                        # ============================================================================================|
                        else:
                            i_start_horizontal_slot = ""
                            j_start_horizontal_slot = ""
                            horizontal_slot_radius = ""

                        # endregion  <<<<======================[Horizontal Pin Slot Numbers]======================>>>>

                        # endregion  <<<<=======================[Horizontal Pin Slot Tool]=======================>>>>

                        # =========================================================================================|
                        # [#] After All Tools have been added to the ToolList of the Template, change the variable |
                        #     (that indicate the end of the tool list) to be <1> to avoid enter the (for) loop over|
                        #     and over again.                                                                      |
                        # =========================================================================================|
                        end_of_tool_list_of_old_horizontal_program = 1

                # endregion  <<<<============================[Tool List]============================>>>>

                # =====================================================================================================|
                # [#] Use (for) loop with (if statement) to access and modify each Variable in (**FEATURE LIST**),     |
                #     (**TOOLS NUMBERS**), and (**DIMENSIONAL VARIABLES**) sections of OldHorizontalTemplate according |
                #     to the Job Information.                                                                          |
                # [#] Use (for) loop to access and read each Line of the Horizontal template, on other words, to access|
                #     and read each element of the list of OldHorizontalTemplateLines (pin_bore_program_lines_of....). |
                # [#] To access and read each line of the template, needs to define variable (call it:'substr') and    |
                #     set it to be the <Variable Name> of the beginning of each line in the Horizontal template        |
                #     (Ex: 'VC118','VC119','VC132','VC155','VC172'...Etc).                                             |
                # [#] Use (find) Method (built-in Function) to search for the Specific Text (that's set on 'substr'    |
                #     variable) in each line of the template.                                                          |
                #     ['find()' is a function that return location(index) of the SupString when find it in the line,   |
                #      if doesn't find anything it will return (index = -1).]                                          |
                # [#] When 'find()' method found and detect the text of 'substr' in Specific line, it will return <0>  |
                #     because all the Variables in the template are locate in the beginning of the line.               |
                #     (Ex: If 'substr' set to be "VC118", the index value of "VC118=0  (LedgeCut)" line is <0> because |
                #      "VC118" is locate in the beginning of the line.                                                 |
                # [#] Define and set variable 'horizontal_program_line_index' to be equal the index value that's found |
                #     (which is equal <0>).                                                                            |
                # [#] Use (if Statement) to check if 'horizontal_program_line_index' is equal <0> to access the line,  |
                #     then define Variable to set the location(index) of the *TemplateLine* that's contain the 'substr'|
                #     of the <Variable Name> (Ex:'VC118_variable_index' is set to indicate the location of line of     |
                #     "VC118=0  (LedgeCut)" in OldHorizontalTemplateLines List).                                       |
                # [#] Use the TemplateLine index that is found to access and modify the line with the necessary changes|
                #     according to the Job information.                                                                |
                # =====================================================================================================|
                for line in pin_bore_program_lines_of_old_horizontal_machine:

                    # region  <<<<============================[Feature List]============================>>>>

                    # region  <<<<=======================[Ledge Cut Status Variable <VC118>]=======================>>>>

                    # ============================================================================================|
                    # [#] To access and read the TemplateLine, needs to define variable (call it:'substr') and set|
                    #     it to be "VC118".                                                                       |
                    # ============================================================================================|
                    substr = "VC118"

                    # ================================================================================================|
                    # [#] Use (find) Method to search for the Specific Text 'VC118' (that's set on 'substr' variable).|
                    # [#] When 'find()' method found and detect the text of 'VC118', it will return <0> because it    |
                    #     locate in the beginning of the line.                                                        |
                    # [#] Define and set variable 'horizontal_program_line_index' to be equal the index value that's  |
                    #     found (which is equal <0>).                                                                 |
                    # ================================================================================================|
                    horizontal_program_line_index = line.find(substr)

                    # ============================================================================================|
                    # [#] Use (if Statement) to check if 'horizontal_program_line_index' is equal <0> (because all|
                    #     variable in template locate in the beginning of the line) to access the TemplateLine.   |
                    # ============================================================================================|
                    if (horizontal_program_line_index == 0):

                        # =======================================================================================|
                        # [#] Define Variable to set the location(index) of the *TemplateLine* that's contain the|
                        #     <Variable Name> (which is 'VC118') in OldHorizontalTemplateLines List.             |
                        # =======================================================================================|
                        VC118_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC118_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC118_variable_index])

                        # ===================================================================================|
                        # [#] Use the TemplateLine index that is found to access and modify the line with the|
                        #     necessary changes according to the Job information.                            |
                        # ===================================================================================|
                        pin_bore_program_lines_of_old_horizontal_machine[VC118_variable_index] = (
                                'VC118=' + format(ledge_cut_availability_status) + '  (LedgeCut)')
                        print("LEDGE_CUT_AVAILABILITY_STATUS in program: ", ledge_cut_availability_status)

                    # endregion  <<<<=====================[Ledge Cut Status Variable <VC118>]======================>>>>

                    # region  <<<<=======================[Pilot Status Variable <VC119>]=======================>>>>

                    substr = "VC119"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0):
                        VC119_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC119_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC119_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC119_variable_index] = (
                                'VC119=' + format(pilot_availability_status) + '  (Pilot)')

                    # endregion  <<<<=======================[Pilot Status Variable <VC119>]=======================>>>>

                    # region  <<<<=======================[LockRing Status Variable <VC121>]=======================>>>>

                    # =====================================================================|
                    # [#] Define Variable to set LockRing Availability Status.             |
                    # [#] Set the Value to be <0> by default and Change it to be <1> when  |
                    #     LockRingCutterWidth Detected.                                    |
                    # [#] Reason to set Status here not above in ToolList Section, because |
                    #     it use same (if statement) to check both of LockRing and C-Fren. |
                    # =====================================================================|
                    lock_ring_availability_status = 0

                    substr = "VC121"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and
                            (lock_ring_cutter_width != 0 and lock_ring_cutter_width is not None)):
                        lock_ring_availability_status = 1
                        VC121_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC121_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC121_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC121_variable_index] = (
                                'VC121=' + format(lock_ring_availability_status) + '  (LockRing)')

                    # endregion  <<<<=====================[LockRing Status Variable <VC121>]======================>>>>

                    # region  <<<<=======================[C-Fren Status Variable <VC122>]=======================>>>>

                    # =====================================================================|
                    # [#] Define Variable to set C/Fren Availability Status.               |
                    # [#] Set the Value to be <0> by default and Change it to be <1> when  |
                    #     C/FrenCutterWidth Detected.                                      |
                    # [#] Reason to set Status here not above in ToolList Section, because |
                    #     it use same (if statement) to check both of LockRing and C-Fren. |
                    # =====================================================================|
                    cfren_availability_status = 0

                    substr = "VC122"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and
                            (cfren_cutter_width != 0 and cfren_cutter_width is not None)):
                        cfren_availability_status = 1
                        VC122_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC122_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC122_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC122_variable_index] = (
                                'VC122=' + format(cfren_availability_status) + '  (C-Fren)')

                    # endregion  <<<<=======================[C-Fren Status Variable <VC122>]=======================>>>>

                    # region  <<<<=======================[DOHS Status Variable <VC123>]=======================>>>>

                    substr = "VC123"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0):
                        VC123_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC123_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC123_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC123_variable_index] = (
                                'VC123=' + format(double_oil_hole_slot_availability_status) + '  (DOHS)')

                    # endregion  <<<<=======================[DOHS Status Variable <VC123>]=======================>>>>

                    # region  <<<<=======================[Notch Status Variable <VC124>]=======================>>>>

                    substr = "VC124"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0):
                        VC124_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC124_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC124_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC124_variable_index] = (
                                'VC124=' + format(notch_availability_status) + '  (CirclipNotch)')

                    # endregion  <<<<=======================[Notch Status Variable <VC124>]=======================>>>>

                    # region  <<<<=======================[H-Slots Status Variable <VC125>]=======================>>>>

                    substr = "VC125"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0):
                        VC125_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC125_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC125_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC125_variable_index] = (
                                'VC125=' + format(horizontal_slots_availability_status) + '  (H-Slots)')

                    # endregion  <<<<======================[H-Slots Status Variable <VC125>]======================>>>>

                    # region  <<<<===================[LedgeCounterbore Status Variable <VC126>]===================>>>>

                    # =====================================================================|
                    # [#] Define Variable to set LedgeCounterbore Availability Status.     |
                    # [#] Set the Value to be <0> by default and Change it to be <1> when  |
                    #     LedgeCounterboreDiameter Detected.                               |
                    # =====================================================================|
                    ledge_counterbore_availability_status = 0

                    substr = "VC126"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and ledge_counterbore_diameter != 0 and
                            ledge_counterbore_diameter is not None):
                        ledge_counterbore_availability_status = 1
                        VC126_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC126_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC126_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC126_variable_index] = (
                                'VC126=' + format(ledge_counterbore_availability_status) + '  (LedgeCounterbore)')

                    # endregion  <<<<==================[LedgeCounterbore Status Variable <VC126>]==================>>>>

                    # region  <<<<===================[Double Notch Status Variable <VC127>]===================>>>>

                    # =====================================================================|
                    # [#] Define Variable to set Double Notch Availability Status.         |
                    # [#] Set the Value to be <0> by default and Change it to be <1> when  |
                    #     NotchAngleSecondLocation Detected.                               |
                    # =====================================================================|
                    double_notch_availability_status = 0
                    substr = "VC127"
                    horizontal_program_line_index = line.find(substr)

                    # =================================================================================================|
                    # [#] Add this condition (notch_angle_second_location != 135) to avoid set the status              |
                    #     to be <1> (when is not needed) because Engineers sometimes add <135> for both Notch Location.|
                    # =================================================================================================|
                    if (horizontal_program_line_index == 0 and
                            (notch_angle_second_location != 0 and notch_angle_second_location is not None
                             and notch_angle_second_location != 135)):
                        double_notch_availability_status = 1
                        VC127_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC127_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC127_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC127_variable_index] = (
                                'VC127=' + format(double_notch_availability_status) + '   (Double Notch)')

                    # endregion  <<<<===================[Double Notch Status Variable <VC127>]===================>>>>

                    # region  <<<<===================[H-Slots Through Status Variable <VC129>]===================>>>>

                    substr = "VC129"
                    horizontal_program_line_index = line.find(substr)
                    # still needs to check the condition
                    if (horizontal_program_line_index == 0 and horizontal_slots_arc_diameter == 0.375):
                        VC129_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC129_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC129_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC129_variable_index] = ('VC129=' + format(
                            horizontal_slots_straight_through_availability_status) + '   (.375SlotsStraightThough)')

                    # endregion  <<<<==================[H-Slots Through Status Variable <VC129>]==================>>>>

                    # endregion  <<<<============================[Feature List]============================>>>>

                    # region  <<<<============================[Tool Number List]============================>>>>

                    # region  <<<<====================[Rough Bore Tool Number Variable <VC130>]====================>>>>

                    substr = "VC130"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0):
                        VC130_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC130_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC130_variable_index])

                        # ===================================================================================|
                        # [#] Use if statement to make the tool number always printed with two digits even it|
                        #     was one digit (like: 1,3,5,7).                                                 |
                        # ===================================================================================|
                        if (rough_bore_tool_number < 10):
                            pin_bore_program_lines_of_old_horizontal_machine[VC130_variable_index] = (
                                    'VC130=0' + format(rough_bore_tool_number) + '  (RoughBoreToolNo)')
                        elif (rough_bore_tool_number >= 10):
                            pin_bore_program_lines_of_old_horizontal_machine[VC130_variable_index] = (
                                    'VC130=' + format(rough_bore_tool_number) + '  (RoughBoreToolNo)')

                    # endregion  <<<<==================[Rough Bore Tool Number Variable <VC130>]===================>>>>

                    # region  <<<<===================[Finish Bore Tool Number Variable <VC132>]===================>>>>

                    substr = "VC132"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0):
                        VC132_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC132_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC132_variable_index])

                        # ===================================================================================|
                        # [#] Use if statement to make the tool number always printed with two digits even it|
                        #     was one digit (like: 1,3,5,7).                                                 |
                        # ===================================================================================|
                        if (finish_bore_tool_number < 10):
                            pin_bore_program_lines_of_old_horizontal_machine[VC132_variable_index] = (
                                    'VC132=0' + format(finish_bore_tool_number) + '  (FinishBoreToolNo)')
                        elif (finish_bore_tool_number >= 10):
                            pin_bore_program_lines_of_old_horizontal_machine[VC132_variable_index] = (
                                    'VC132=' + format(finish_bore_tool_number) + '  (FinishBoreToolNo)')

                    # endregion  <<<<=================[Finish Bore Tool Number Variable <VC132>]==================>>>>

                    # region  <<<<================[LockRing/C-Fren Tool Number Variable <VC134>]=================>>>>

                    substr = "VC134"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0):
                        VC134_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC134_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC134_variable_index])

                        # ===================================================================================|
                        # [#] Use if statement to make the tool number always printed with two digits even it|
                        #     was one digit (like: 1,3,5,7).                                                 |
                        # ===================================================================================|
                        if (lock_ring_tool_number < 10):
                            pin_bore_program_lines_of_old_horizontal_machine[VC134_variable_index] = (
                                    'VC134=0' + format(lock_ring_tool_number) + '  (LockRingToolNo)')
                        elif (lock_ring_tool_number >= 10):
                            pin_bore_program_lines_of_old_horizontal_machine[VC134_variable_index] = (
                                    'VC134=' + format(lock_ring_tool_number) + '  (LockRingToolNo)')

                    # endregion  <<<<===============[LockRing/C-Fren Tool Number Variable <VC134>]================>>>>

                    # region  <<<<================[DOHS Tool Number Variable <VC136>]=================>>>>

                    substr = "VC136"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and double_oil_hole_slot_availability_status == 1):
                        VC136_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC136_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC136_variable_index])

                        # ===================================================================================|
                        # [#] Use if statement to make the tool number always printed with two digits even it|
                        #     was one digit (like: 1,3,5,7).                                                 |
                        # [#] ToolNumber for DoubleOilHoleSlot is always <5>, but use (if statement) just in |
                        #     case the tool number change on future.                                         |
                        # ===================================================================================|
                        if (double_oil_hole_slot_tool_number < 10):
                            pin_bore_program_lines_of_old_horizontal_machine[VC136_variable_index] = (
                                    'VC136=0' + format(double_oil_hole_slot_tool_number) + '  (DOHSToolNo)')
                        elif (double_oil_hole_slot_tool_number >= 10):
                            pin_bore_program_lines_of_old_horizontal_machine[VC136_variable_index] = (
                                    'VC136=' + format(double_oil_hole_slot_tool_number) + '  (DOHSToolNo)')

                    # endregion  <<<<================[DOHS Tool Number Variable <VC136>]=================>>>>

                    # endregion  <<<<============================[Tool Number List]============================>>>>

                    # region <<<<================[Dimensional Variables]=================>>>>

                    # region <<<<================[Pin Hole Diameter Variable <VC149>]=================>>>>

                    substr = "VC149"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and
                            (pin_hole_diameter != 0 and pin_hole_diameter is not None)):
                        VC149_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC149_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC149_variable_index])

                        # ============================================================================================|
                        # [#] Use (format(pin_hole_diameter, '.3f')) because it needs to set The PinHoleDiameter to be|
                        #     always 3-digits to make sure it's working with H-Slot and DOHS logic in the template.   |
                        # Still needs to check with Programming                                                       |
                        # ============================================================================================|
                        pin_bore_program_lines_of_old_horizontal_machine[VC149_variable_index] = (
                                'VC149=' + format(pin_hole_diameter, '.3f') + '  (PinHoleDiameter)')

                    # endregion <<<<================[Pin Hole Diameter Variable <VC149>]=================>>>>

                    # region <<<<================[Forge Ref Length Variable <VC150>]=================>>>>

                    substr = "VC150"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0):

                        # ==========================================================================================|
                        # [#] Use (if-else statements) because Some Forging (Especially JE Forging) is missing some |
                        #     Information(dimension), therefore it needs to check if the dimension is missing to ask|
                        #     the user to enter the value.                                                          |
                        # ==========================================================================================|
                        if ((forge_ref_length != 0 and forge_ref_length is not None and forge_ref_length != "")):
                            VC150_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                            print(VC150_variable_index)
                            print(pin_bore_program_lines_of_old_horizontal_machine[VC150_variable_index])

                            pin_bore_program_lines_of_old_horizontal_machine[VC150_variable_index] = (
                                    'VC150=' + format(forge_ref_length) + '  (ForgeRefLength)')

                        # =====================================================================|
                        # if the Forging dimension is missing, ask the user to enter the value.|
                        # =====================================================================|
                        else:
                            # =========================================================|
                            # [#] Set verification_message to be the current Situation.|
                            # =========================================================|
                            verification_messages_of_creating_old_horizontal_machine_program = [
                                "[u]Forge Ref Length[/u] " +
                                "does NOT found, Please take a look to Forging Info or "
                                "Model and Enter the value :"]
                            # ==============================================================|
                            # [#] Add verification_message to the Confirmation Details list.|
                            # ==============================================================|
                            old_horizontal_program_confirmation_email_message_list.append(
                                "# It was need confirmation of :" + "\n" + (
                                    "**Forge Ref Length** does NOT found, Please take a look to Forging Info or "
                                    "Model and Enter the value"))
                            # ==================================================================|
                            # [#] title: set it to have the header and the verification_message.|
                            # ==================================================================|
                            self.title = '[color=0066ff]Confirmation Message[/color]' + '\n' + '\n' + \
                                         '[b][i][color=ffffff]' + \
                                         verification_messages_of_creating_old_horizontal_machine_program[0] + \
                                         '[/color][/i][/b]'
                            # ====================================================================|
                            # [#] Create the window_message that contain the Confirmation Details.|
                            # ====================================================================|
                            self.Dialog_BoxLayout = BoxLayout(height=30)
                            # ===========================================================|
                            # [#] Create the TextInput field to let user enter the value.|
                            # [#] (input_filter="float"): to accept Numeric input Only.  |
                            # ===========================================================|
                            self.forge_ref_length_confirmation_text_field = TextInput(
                                hint_text="Enter Value", multiline=False, input_filter="float",
                                background_color=[60 / 255.0, 60 / 255.0, 60 / 255.0, 1],
                                foreground_color=[1, 1, 1, 1], size_hint_y=None, height=30)
                            # ===========================================================================|
                            # [#] Add the TextInput field created above to be part of the window_message.|
                            # ===========================================================================|
                            self.Dialog_BoxLayout.add_widget(self.forge_ref_length_confirmation_text_field)
                            # ===================================================================================|
                            # [#] Call Function of (need_confirmation) by pass the parameters that created:      |
                            #   [#] title: self.title                                                            |
                            #   [#] sub_function: self.enter_forge_ref_length_value (it's the Function created   |
                            #        below to set the Forge Ref Length).                                         |
                            #   [#] dialog_type: "custom" (it used to indicate that user needs to enter value).  |
                            #   [#] content: self.Dialog_BoxLayout (the window_message that have the Details).   |
                            # ===================================================================================|
                            need_confirmation_to_create_old_horizontal_machine_program(
                                self, self.title, self.enter_forge_ref_length_value, "custom", self.Dialog_BoxLayout)
                            print("Forge Ref Length BEFORE User Input: ", forge_ref_length)
                            return

                    # endregion <<<<================[Forge Ref Length Variable <VC150>]=================>>>>

                    # region <<<<================[Pilot Bore Depth Variable <VC151>]=================>>>>

                    substr = "VC151"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and
                            (pilot_bore_depth != 0 and pilot_bore_depth is not None)):
                        VC151_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC151_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC151_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC151_variable_index] = (
                                'VC151=' + format(pilot_bore_depth) + '  (PilotBoreDepth)')

                    # endregion <<<<================[Pilot Bore Depth Variable <VC151>]=================>>>>

                    # region <<<<================[Rough Bore Speed Variable <VC152>]=================>>>>

                    substr = "VC152"
                    horizontal_program_line_index = line.find(substr)
                    # =====================================================================================|
                    # [#] Use (finish_bore_tool_number == 28) to indicate it use BoringBar tool which needs|
                    #     to slow the speed to 6000, otherwise leave it as it is in the template (8000).   |
                    # maybe needs to add T45 to the condition (check with programming).                    |
                    # =====================================================================================|
                    if (horizontal_program_line_index == 0 and finish_bore_tool_number == 28):
                        VC152_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC152_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC152_variable_index])

                        rough_bore_speed = 6000
                        pin_bore_program_lines_of_old_horizontal_machine[VC152_variable_index] = (
                                'VC152=' + format(rough_bore_speed) + '  (RoughBoreSpeed)')

                    # endregion <<<<================[Rough Bore Speed Variable <VC152>]=================>>>>

                    # region <<<<================[Rough Bore Feed Variable <VC153>]=================>>>>

                    substr = "VC153"
                    horizontal_program_line_index = line.find(substr)
                    # =====================================================================================|
                    # [#] Use (finish_bore_tool_number == 28) to indicate it use BoringBar tool which needs|
                    #     to slow the FeedRate to 60, otherwise leave it as it is in the template (100).   |
                    # maybe needs to add T45 to the condition (check with programming).                    |
                    # =====================================================================================|
                    if (horizontal_program_line_index == 0 and finish_bore_tool_number == 28):
                        VC153_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC153_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC153_variable_index])

                        rough_bore_feed = 60
                        pin_bore_program_lines_of_old_horizontal_machine[VC153_variable_index] = (
                                'VC153=' + format(rough_bore_feed) + '  (RoughBoreFeed)')

                    # endregion <<<<================[Rough Bore Feed Variable <VC153>]=================>>>>

                    # region <<<<============[X_distance_from_origin_to_pin_center Variable <VC154>]=============>>>>

                    substr = "VC154"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and
                            (pilot_bore_depth != 0 and pilot_bore_depth is not None) and
                            (pilot_to_pin != 0 and pilot_to_pin is not None) and
                            (pin_hole_diameter != 0 and pin_hole_diameter is not None)):
                        VC154_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC154_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC154_variable_index])

                        # ==========================================================================================|
                        # [#] Calculation of the X_distance from the Piston origin to the center of the PinHole:    |
                        #     (X_distance from the Piston origin to the Piston Pilot) -                             |
                        #       (X_distance from the Piston Pilot to closest point of PinHole) -                    |
                        #          (PinHole Diameter divided by 2).                                                 |
                        # [#] Needs to take the Absolute Value of PilotToPin to have always Positive Value, needs to|
                        #     use Absolute because most PilotToPin values in the DataBase are Negative.             |
                        # ==========================================================================================|
                        global X_distance_from_origin_to_pin_center
                        X_distance_from_origin_to_pin_center = format(
                            pilot_bore_depth - abs(pilot_to_pin) - (pin_hole_diameter / 2), '.4f')
                        pin_bore_program_lines_of_old_horizontal_machine[VC154_variable_index] = (
                                'VC154=-' + format(X_distance_from_origin_to_pin_center) + '  (xPinCenter)')

                    # endregion <<<<===========[X_distance_from_origin_to_pin_center Variable <VC154>]============>>>>

                    # region <<<<================[Offset Variable <VC155>]=================>>>>

                    substr = "VC155"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0):

                        # =======================================================================================|
                        # [#] Use (if statement) to check if the Offset Amount is missing, therefore it needs to |
                        #     ask the user to enter the value.                                                   |
                        # =======================================================================================|
                        if (offset_amount is None):
                            # =========================================================|
                            # [#] Set verification_message to be the current Situation.|
                            # =========================================================|
                            verification_messages_of_creating_old_horizontal_machine_program = [
                                "[u]Offset Amount[/u] does NOT found, Please Enter the value"]
                            # ==============================================================|
                            # [#] Add verification_message to the Confirmation Details list.|
                            # ==============================================================|
                            old_horizontal_program_confirmation_email_message_list.append(
                                "# It was need confirmation of :" + "\n" + (
                                    "**Offset Amount** does NOT found, Please Enter the value"))
                            # ==================================================================|
                            # [#] title: set it to have the header and the verification_message.|
                            # ==================================================================|
                            self.title = '[color=0066ff]Confirmation Message[/color]' + '\n' + '\n' + \
                                         '[b][i][color=ffffff]' + \
                                         verification_messages_of_creating_old_horizontal_machine_program[0] + \
                                         '[/color][/i][/b]'
                            # ====================================================================|
                            # [#] Create the window_message that contain the Confirmation Details.|
                            # ====================================================================|
                            self.Dialog_BoxLayout = BoxLayout(height=30)
                            # ===========================================================|
                            # [#] Create the TextInput field to let user enter the value.|
                            # [#] (input_filter="float"): to accept Numeric input Only.  |
                            # ===========================================================|
                            self.offset_amount_confirmation_text_field = TextInput(
                                hint_text="Enter Value", multiline=False, input_filter="float",
                                background_color=[60 / 255.0, 60 / 255.0, 60 / 255.0, 1],
                                foreground_color=[1, 1, 1, 1], size_hint_y=None, height=30)
                            # ===========================================================================|
                            # [#] Add the TextInput field created above to be part of the window_message.|
                            # ===========================================================================|
                            self.Dialog_BoxLayout.add_widget(self.offset_amount_confirmation_text_field)
                            # =================================================================================|
                            # [#] Call Function of (need_confirmation) by pass the parameters that created:    |
                            #   [#] title: self.title                                                          |
                            #   [#] sub_function: self.enter_offset_value (it's the Function created           |
                            #        below to set the Offset Amount).                                          |
                            #   [#] dialog_type: "custom" (it used to indicate that user needs to enter value).|
                            #   [#] content: self.Dialog_BoxLayout (the window_message that have the Details). |
                            # =================================================================================|
                            need_confirmation_to_create_old_horizontal_machine_program(
                                self, self.title, self.enter_offset_value, "custom", self.Dialog_BoxLayout)

                            print("offset value before ", offset_amount)
                            return

                        # =======================================================================================|
                        # [#] Use (if statement) to check if the Offset Direction is missing, therefore it needs |
                        #     to ask the user to choose the Offset Direction option.                             |
                        # =======================================================================================|
                        if (offset_direction == "" and offset_amount != 0):
                            # =========================================================|
                            # [#] Set verification_message to be the current Situation.|
                            # =========================================================|
                            verification_messages_of_creating_old_horizontal_machine_program = [
                                "[u]Offset Direction[/u] does NOT found, Please Choose the direction"]
                            # ==============================================================|
                            # [#] Add verification_message to the Confirmation Details list.|
                            # ==============================================================|
                            old_horizontal_program_confirmation_email_message_list.append(
                                "# It was need confirmation of :" + "\n" + (
                                    "**Offset Direction** does NOT found, Please Choose the direction"))
                            # ==================================================================|
                            # [#] title: set it to have the header and the verification_message.|
                            # ==================================================================|
                            self.title = '[color=0066ff]Confirmation Message[/color] ' + '\n' + '\n' + \
                                         '[b][i][color=ffffff]' + \
                                         verification_messages_of_creating_old_horizontal_machine_program[0] + \
                                         '[/color][/i][/b]'
                            # ======================================================================|
                            # [#] Set option of the Offset Direction To0.                           |
                            # [#] OldHorizontalMachineItem:                                         |
                            #     The Class that created above to be able to create List's Items.   |
                            # [#] set_offset_direction_To0:                                         |
                            #     The Function that created below to set the Offset Direction To0.  |
                            # [#] offset_To0_option_status:                                         |
                            #     Set Status to be 'False' by default to indicate is not picked yet,|
                            #     and it will change to be 'True' when user choose this option.     |
                            # ======================================================================|
                            self.offset_To0_option = OldHorizontalMachineItem(
                                text="OFFSET To0", on_release=self.set_offset_direction_To0)
                            self.offset_To0_option_status = False
                            # ======================================================================|
                            # [#] Set option of the Offset Direction To180.                         |
                            # [#] OldHorizontalMachineItem:                                         |
                            #     The Class that created above to be able to create List's Items.   |
                            # [#] set_offset_direction_To180:                                       |
                            #     The Function that created below to set the Offset Direction To180.|
                            # [#] offset_To180_option_status:                                       |
                            #     Set Status to be 'False' by default to indicate is not picked yet,|
                            #     and it will change to be 'True' when user choose this option.     |
                            # ======================================================================|
                            self.offset_To180_option = OldHorizontalMachineItem(
                                text="OFFSET To180", on_release=self.set_offset_direction_To180)
                            self.offset_To180_option_status = False
                            # ========================================================================|
                            # [#] Set option of the Offset Each Way Direction.                        |
                            # [#] OldHorizontalMachineItem:                                           |
                            #     The Class that created above to be able to create List's Items.     |
                            # [#] set_offset_direction_each_way:                                      |
                            #     The Function that created below to set the Offset Direction EachWay.|
                            # [#] offset_To0_option_status:                                           |
                            #     Set Status to be 'False' by default to indicate is not picked yet,  |
                            #     and it will change to be 'True' when user choose this option.       |
                            # ========================================================================|
                            self.offset_each_way_option = OldHorizontalMachineItem(
                                text="OFFSET EACH WAY", on_release=self.set_offset_direction_each_way)
                            self.offset_each_way_option_status = False
                            # ==========================================|
                            # [#] Create items list to have All options.|
                            # ==========================================|
                            self.items = [self.offset_To0_option, self.offset_To180_option, self.offset_each_way_option]
                            # =============================================================================|
                            # [#] Call Function of (need_confirmation) by pass the parameters that created:|
                            #   [#] title: self.title                                                      |
                            #   [#] sub_function: self.choose_offset_direction (it's Function              |
                            #                     created below to set the Offset Direction).              |
                            #   [#] dialog_type: "confirmation" (it used to indicate that user             |
                            #                     needs to choose Option).                                 |
                            #   [#] content: self.items (the ItemsList that have the options).             |
                            # =============================================================================|
                            need_confirmation_to_create_old_horizontal_machine_program(
                                self, self.title, self.choose_offset_direction, "confirmation", self.items)

                            print("offset direction before ", offset_direction)
                            return

                        # =====================================================================================|
                        # [#] Use (if statement) to Set the Offset Direction that will use in math calculation.|
                        # [#] If can't set the value, needs to ask the User to choose Offset direction again.  |
                        # =====================================================================================|
                        if (offset_amount == 0):
                            offset_direction_for_math = 1
                        elif (offset_direction == "OFFSET EACH WAY" or offset_direction == "OFFSET To0"):
                            offset_direction_for_math = 1
                        elif (offset_direction == "OFFSET To180"):
                            offset_direction_for_math = -1
                        else:
                            need_confirmation_to_create_old_horizontal_machine_program(
                                self, self.title, self.choose_offset_direction, "confirmation", self.items)
                            return

                        global VC155_variable_index
                        VC155_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC155_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC155_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC155_variable_index] = (
                                'VC155=' + format(offset_amount * offset_direction_for_math) + '  (Offset)')

                    # endregion <<<<================[Offset Variable <VC155>]=================>>>>

                    # region <<<<================[Z_Value Of Top Of Pin Bore Variable <VC156>]=================>>>>

                    substr = "VC156"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0):
                        VC156_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC156_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC156_variable_index])

                        # =========================================================================================|
                        # [#] Use (if-else statements) because Some Forging (Especially JE Forging) is missing some|
                        #     Information(dimension), therefore it needs to check if the dimension is missing to   |
                        #     ask the user to enter the value.                                                     |
                        # =========================================================================================|
                        if (forging_diameter != 0 and forging_diameter is not None):
                            pin_bore_program_lines_of_old_horizontal_machine[VC156_variable_index] = (
                                    'VC156=[' + format(forging_diameter) + '/2]' +
                                    '  (zPinBoreTop - ? IS Forging Diameter)')
                        else:
                            # =========================================================|
                            # [#] Set verification_message to be the current Situation.|
                            # =========================================================|
                            verification_messages_of_creating_old_horizontal_machine_program = [
                                "[u]Forging Diameter[/u] " +
                                "does NOT found, Please take a look to Forging Info or "
                                "Model and Enter the value :"]
                            # ==============================================================|
                            # [#] Add verification_message to the Confirmation Details list.|
                            # ==============================================================|
                            old_horizontal_program_confirmation_email_message_list.append(
                                "# It was need confirmation of :" + "\n" + (
                                    "**Forging Diameter** does NOT found, Please take a look to Forging Info or "
                                    "Model and Enter the value"))
                            # ==================================================================|
                            # [#] title: set it to have the header and the verification_message.|
                            # ==================================================================|
                            self.title = '[color=0066ff]Confirmation Message[/color]' + '\n' + '\n' + \
                                         '[b][i][color=ffffff]' + \
                                         verification_messages_of_creating_old_horizontal_machine_program[0] + \
                                         '[/color][/i][/b]'
                            # ====================================================================|
                            # [#] Create the window_message that contain the Confirmation Details.|
                            # ====================================================================|
                            self.Dialog_BoxLayout = BoxLayout(height=30)
                            # ===========================================================|
                            # [#] Create the TextInput field to let user enter the value.|
                            # [#] (input_filter="float"): to accept Numeric input Only.  |
                            # ===========================================================|
                            self.forging_diameter_confirmation_text_field = TextInput(
                                hint_text="Enter Value", multiline=False, input_filter="float",
                                background_color=[60 / 255.0, 60 / 255.0, 60 / 255.0, 1],
                                foreground_color=[1, 1, 1, 1], size_hint_y=None, height=30)
                            # ===========================================================================|
                            # [#] Add the TextInput field created above to be part of the window_message.|
                            # ===========================================================================|
                            self.Dialog_BoxLayout.add_widget(self.forging_diameter_confirmation_text_field)
                            # ===================================================================================|
                            # [#] Call Function of (need_confirmation) by pass the parameters that created:      |
                            #   [#] title: self.title                                                            |
                            #   [#] sub_function: self.enter_forging_diameter_value (it's the Function created   |
                            #        below to set the Forging Diameter).                                         |
                            #   [#] dialog_type: "custom" (it used to indicate that user needs to enter value).  |
                            #   [#] content: self.Dialog_BoxLayout (the window_message that have the Details).   |
                            # ===================================================================================|
                            need_confirmation_to_create_old_horizontal_machine_program(
                                self, self.title, self.enter_forging_diameter_value, "custom", self.Dialog_BoxLayout)
                            print("Forging Diameter BEFORE User Input: ", forging_diameter)
                            return

                    # endregion <<<<================[Z_Value Of Top Of Pin Bore <VC156>]=================>>>>

                    # region <<<<===============[Z_Value Of Bottom Of Rough Bore Variable <VC157>]================>>>>
                    # ===================================================================================|
                    # [#] No need to do anything here, it will just use how it is in Horizontal template.|
                    # ===================================================================================|
                    # endregion <<<<==============[Z_Value Of Bottom Of Rough Bore Variable <VC157>]==============>>>>

                    # region <<<<================[Z_Value Finish Bore Bottom Variable <VC158>]=================>>>>

                    substr = "VC158"
                    horizontal_program_line_index = line.find(substr)
                    # =========================================================================================|
                    # [#] Use (finish_bore_tool_number == 28) to indicate it use BoringBar tool which needs    |
                    #     to change the value to be <0.1>, otherwise leave it as it is in the template (1.156).|
                    # maybe needs to add T45 to the condition (check with programming).                        |
                    # =========================================================================================|
                    if (horizontal_program_line_index == 0 and finish_bore_tool_number == 28):
                        VC158_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC158_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC158_variable_index])

                        value_used_in_Z_value_finish_bore_bottom = 0.1
                        pin_bore_program_lines_of_old_horizontal_machine[VC158_variable_index] = (
                                'VC158=-[VC156+' + format(value_used_in_Z_value_finish_bore_bottom) + ']' +
                                '  (zFinishBoreBottom)')

                    # endregion <<<<===============[Z_Value Finish Bore Bottom Variable <VC158>]================>>>>

                    # region <<<<================[Forging Outside Boss Spacing Variable <VC159>]=================>>>>

                    substr = "VC159"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0):
                        VC159_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC159_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC159_variable_index])

                        # ===========================================================================================|
                        # [#] Use (if statement) to set value of ForgingOutsideBossSpacing according to Forging Type,|
                        #     if it is NOT Fully Around set the value as it is stored in the Forging DataBase, if it |
                        #     is Fully Around set the value to be the Forging Diameter.                              |
                        # [#] Use (else statement) because Some Forging (Especially JE Forging) is missing some      |
                        #     Information(dimension), therefore it needs to check if the dimension is missing to     |
                        #     ask the user to enter the value.                                                       |
                        # ===========================================================================================|
                        if (forging_outside_boss_spacing != 0 and forging_outside_boss_spacing is not None):
                            pin_bore_program_lines_of_old_horizontal_machine[VC159_variable_index] = (
                                    'VC159=' + format(forging_outside_boss_spacing) + '  (OutsideBossSpacing)')
                        # =========================================================================================|
                        # [#] If the Forging is Fully Around (ie: NO value of 'forging_outside_boss_spacing'), then|
                        #     set the value to be the Forging Diameter.                                            |
                        # =========================================================================================|
                        elif (forging_outside_boss_spacing == 0 or forging_outside_boss_spacing is None):
                            pin_bore_program_lines_of_old_horizontal_machine[VC159_variable_index] = (
                                    'VC159=' + format(forging_diameter) + '  (OutsideBossSpacing)')
                        # =============================================================================|
                        # [#] Use (else statement) to ask the user to enter the value if it is missing.|
                        # =============================================================================|
                        else:
                            # =========================================================|
                            # [#] Set verification_message to be the current Situation.|
                            # =========================================================|
                            verification_messages_of_creating_old_horizontal_machine_program = [
                                "[u]Forging Outside Boss Spacing[/u] " +
                                "does NOT found, Please take a look to Forging Info or Model and Enter the value :"]
                            # ==============================================================|
                            # [#] Add verification_message to the Confirmation Details list.|
                            # ==============================================================|
                            old_horizontal_program_confirmation_email_message_list.append(
                                "# It was need confirmation of :" + "\n" + (
                                    "**Forging Outside Boss Spacing** does NOT found, "
                                    "Please take a look to Forging Info or Model and Enter the value"))
                            # ==================================================================|
                            # [#] title: set it to have the header and the verification_message.|
                            # ==================================================================|
                            self.title = '[color=0066ff]Confirmation Message[/color]' + '\n' + '\n' + \
                                         '[b][i][color=ffffff]' + \
                                         verification_messages_of_creating_old_horizontal_machine_program[0] + \
                                         '[/color][/i][/b]'
                            # ====================================================================|
                            # [#] Create the window_message that contain the Confirmation Details.|
                            # ====================================================================|
                            self.Dialog_BoxLayout = BoxLayout(height=30)
                            # ===========================================================|
                            # [#] Create the TextInput field to let user enter the value.|
                            # [#] (input_filter="float"): to accept Numeric input Only.  |
                            # ===========================================================|
                            self.forging_outside_boss_spacing_confirmation_text_field = TextInput(
                                hint_text="Enter Value", multiline=False, input_filter="float",
                                background_color=[60 / 255.0, 60 / 255.0, 60 / 255.0, 1],
                                foreground_color=[1, 1, 1, 1], size_hint_y=None, height=30)
                            # ===========================================================================|
                            # [#] Add the TextInput field created above to be part of the window_message.|
                            # ===========================================================================|
                            self.Dialog_BoxLayout.add_widget(self.forging_outside_boss_spacing_confirmation_text_field)
                            # ====================================================================================|
                            # [#] Call Function of (need_confirmation) by pass the parameters that created:       |
                            #   [#] title: self.title                                                             |
                            #   [#] sub_function: self.enter_forging_outside_boss_spacing_value (it's the Function|
                            #        created below to set the Forging Outside Boss Spacing).                      |
                            #   [#] dialog_type: "custom" (it used to indicate that user needs to enter value).   |
                            #   [#] content: self.Dialog_BoxLayout (the window_message that have the Details).    |
                            # ====================================================================================|
                            need_confirmation_to_create_old_horizontal_machine_program(
                                self, self.title, self.enter_forging_outside_boss_spacing_value,
                                "custom", self.Dialog_BoxLayout)
                            print("Forging Outside Boss Spacing BEFORE User Input: ", forging_outside_boss_spacing)
                            return

                    # endregion <<<<================[Forging Outside Boss Spacing Variable <VC159>]=================>>>>

                    # region <<<<================[Ledge Tool Diameter Variable <VC160>]=================>>>>

                    substr = "VC160"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and
                            (ledge_cut_availability_status == 1 or
                             (ledge_counterbore_diameter != 0 and ledge_counterbore_diameter is not None))):
                        VC160_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC160_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC160_variable_index])
                        # ======================================================================================|
                        # [#] If the PinHoleDiameter is SmallerThan <0.629>, needs to change the diameter to    |
                        #     be <0.375> (the Diameter of the Ledge Tool that used with the small PinHole Size).|
                        #     Otherwise leave it as it is in the template (0.625).                              |
                        # ======================================================================================|
                        if (pin_hole_diameter < 0.629):
                            global ledge_tool_diameter
                            ledge_tool_diameter = 0.375
                            pin_bore_program_lines_of_old_horizontal_machine[VC160_variable_index] = (
                                    'VC160=' + format(ledge_tool_diameter) + '  (LedgeToolDiameter)')

                    # endregion <<<<================[Ledge Tool Diameter Variable <VC160>]=================>>>>

                    # region <<<<================[Z_Value Of Top Of LockRing Variable <VC161>]=================>>>>

                    substr = "VC161"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and
                            (lock_ring_cutter_width != 0 and lock_ring_cutter_width is not None)):

                        # =======================================================================================|
                        # [#] Use (if-else statements) to check if the LockRing ID Spacing is missing, therefore |
                        #     it needs to ask the user to enter the value.                                       |
                        # =======================================================================================|
                        if (lock_ring_ID_spacing != 0 and lock_ring_ID_spacing is not None):
                            VC161_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                            print(VC161_variable_index)
                            print(pin_bore_program_lines_of_old_horizontal_machine[VC161_variable_index])

                            pin_bore_program_lines_of_old_horizontal_machine[VC161_variable_index] = (
                                    'VC161=[' + format(lock_ring_ID_spacing) + '/2]' +
                                    '  (zTopLockRing = Replace 0 with Lockring ID Spacing)')
                        else:
                            # =========================================================|
                            # [#] Set verification_message to be the current Situation.|
                            # =========================================================|
                            verification_messages_of_creating_old_horizontal_machine_program = [
                                "[u]Lock Ring ID Spacing[/u] " +
                                "does NOT found, Please Enter the value :"]
                            # ==============================================================|
                            # [#] Add verification_message to the Confirmation Details list.|
                            # ==============================================================|
                            old_horizontal_program_confirmation_email_message_list.append(
                                "# It was need confirmation of :" + "\n" + (
                                    "**Lock Ring ID Spacing** does NOT found, Please Enter the value"))
                            # ==================================================================|
                            # [#] title: set it to have the header and the verification_message.|
                            # ==================================================================|
                            self.title = '[color=0066ff]Confirmation Message[/color]' + '\n' + '\n' + \
                                         '[b][i][color=ffffff]' + \
                                         verification_messages_of_creating_old_horizontal_machine_program[0] + \
                                         '[/color][/i][/b]'
                            # ====================================================================|
                            # [#] Create the window_message that contain the Confirmation Details.|
                            # ====================================================================|
                            self.Dialog_BoxLayout = BoxLayout(height=30)
                            # ===========================================================|
                            # [#] Create the TextInput field to let user enter the value.|
                            # [#] (input_filter="float"): to accept Numeric input Only.  |
                            # ===========================================================|
                            self.lock_ring_ID_spacing_confirmation_text_field = TextInput(
                                hint_text="Enter Value", multiline=False, input_filter="float",
                                background_color=[60 / 255.0, 60 / 255.0, 60 / 255.0, 1],
                                foreground_color=[1, 1, 1, 1], size_hint_y=None, height=30)
                            # ===========================================================================|
                            # [#] Add the TextInput field created above to be part of the window_message.|
                            # ===========================================================================|
                            self.Dialog_BoxLayout.add_widget(self.lock_ring_ID_spacing_confirmation_text_field)
                            # =================================================================================|
                            # [#] Call Function of (need_confirmation) by pass the parameters that created:    |
                            #   [#] title: self.title                                                          |
                            #   [#] sub_function: self.enter_lock_ring_ID_spacing_value (it's the Function     |
                            #        created below to set the Lock Ring ID Spacing).                           |
                            #   [#] dialog_type: "custom" (it used to indicate that user needs to enter value).|
                            #   [#] content: self.Dialog_BoxLayout (the window_message that have the Details). |
                            # =================================================================================|
                            need_confirmation_to_create_old_horizontal_machine_program(
                                self, self.title, self.enter_lock_ring_ID_spacing_value,
                                "custom", self.Dialog_BoxLayout)
                            print("Lock Ring ID Spacing BEFORE User Input: ", lock_ring_ID_spacing)
                            return
                    # endregion <<<<================[Z_Value Of Top Of LockRing Variable <VC161>]=================>>>>

                    # region <<<<================[Z_Value Of Bottom Of LockRing Variable <VC162>]=================>>>>

                    substr = "VC162"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and
                            (lock_ring_cutter_width != 0 and lock_ring_cutter_width is not None)):

                        VC162_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC162_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC162_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC162_variable_index] = (
                                'VC162=-[VC161+' + format(lock_ring_cutter_width) + ']' +
                                '  (zBottomLockRing IS zTopLockRing + LR CUTTER WIDTH)')

                    # endregion <<<<==============[Z_Value Of Bottom Of LockRing Variable <VC162>]================>>>>

                    # region <<<<================[LockRing Cut Radius Variable <VC163>]=================>>>>

                    substr = "VC163"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and
                            (lock_ring_cutter_width != 0 and lock_ring_cutter_width is not None)):

                        # =====================================================================================|
                        # [#] Use (if-else statements) to check if the LockRing Diameter is missing, therefore |
                        #     it needs to ask the user to enter the value.                                     |
                        # =====================================================================================|
                        if (lock_ring_diameter != 0 and lock_ring_diameter is not None):
                            VC163_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                            print(VC163_variable_index)
                            print(pin_bore_program_lines_of_old_horizontal_machine[VC163_variable_index])
                            pin_bore_program_lines_of_old_horizontal_machine[VC163_variable_index] = (
                                'VC163=[[' + format(lock_ring_diameter) + '-' + format(lock_ring_tool_diameter) +
                                ']/2]' + '  (LRCutRadius)')
                        else:
                            # =========================================================|
                            # [#] Set verification_message to be the current Situation.|
                            # =========================================================|
                            verification_messages_of_creating_old_horizontal_machine_program = [
                                "[u]Lock Ring Diameter[/u] " +
                                "does NOT found, Please Enter the value :"]
                            # ==============================================================|
                            # [#] Add verification_message to the Confirmation Details list.|
                            # ==============================================================|
                            old_horizontal_program_confirmation_email_message_list.append(
                                "# It was need confirmation of :" + "\n" + (
                                    "**Lock Ring Diameter** does NOT found, Please Enter the value"))
                            # ==================================================================|
                            # [#] title: set it to have the header and the verification_message.|
                            # ==================================================================|
                            self.title = '[color=0066ff]Confirmation Message[/color]' + '\n' + '\n' + \
                                         '[b][i][color=ffffff]' + \
                                         verification_messages_of_creating_old_horizontal_machine_program[0] + \
                                         '[/color][/i][/b]'
                            # ====================================================================|
                            # [#] Create the window_message that contain the Confirmation Details.|
                            # ====================================================================|
                            self.Dialog_BoxLayout = BoxLayout(height=30)
                            # ===========================================================|
                            # [#] Create the TextInput field to let user enter the value.|
                            # [#] (input_filter="float"): to accept Numeric input Only.  |
                            # ===========================================================|
                            self.lock_ring_diameter_confirmation_text_field = TextInput(
                                hint_text="Enter Value", multiline=False, input_filter="float",
                                background_color=[60 / 255.0, 60 / 255.0, 60 / 255.0, 1],
                                foreground_color=[1, 1, 1, 1], size_hint_y=None, height=30)
                            # ===========================================================================|
                            # [#] Add the TextInput field created above to be part of the window_message.|
                            # ===========================================================================|
                            self.Dialog_BoxLayout.add_widget(self.lock_ring_diameter_confirmation_text_field)
                            # =================================================================================|
                            # [#] Call Function of (need_confirmation) by pass the parameters that created:    |
                            #   [#] title: self.title                                                          |
                            #   [#] sub_function: self.enter_lock_ring_diameter_value (it's the Function       |
                            #        created below to set the Lock Ring Diameter).                             |
                            #   [#] dialog_type: "custom" (it used to indicate that user needs to enter value).|
                            #   [#] content: self.Dialog_BoxLayout (the window_message that have the Details). |
                            # =================================================================================|
                            need_confirmation_to_create_old_horizontal_machine_program(
                                self, self.title, self.enter_lock_ring_diameter_value,
                                "custom", self.Dialog_BoxLayout)
                            print("Lock Ring Diameter BEFORE User Input: ", lock_ring_diameter)
                            return

                    # endregion <<<<================[LockRing Cut Radius Variable <VC163>]=================>>>>

                    # region <<<<================[Z_Value Of Top Of C-Fren Variable <VC164>]=================>>>>

                    substr = "VC164"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and
                            (cfren_cutter_width != 0 and cfren_cutter_width is not None)):

                        # =====================================================================================|
                        # [#] Use (if-else statements) to check if the C/Fren ID Spacing is missing, therefore |
                        #     it needs to ask the user to enter the value.                                     |
                        # =====================================================================================|
                        if (cfren_ID_spacing != 0 and cfren_ID_spacing is not None):
                            VC164_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                            print(VC164_variable_index)
                            print(pin_bore_program_lines_of_old_horizontal_machine[VC164_variable_index])

                            pin_bore_program_lines_of_old_horizontal_machine[VC164_variable_index] = (
                                    'VC164=[' + format(cfren_ID_spacing) + '/2]' +
                                    '  (zTopCFREN = Replace 0 with CFren ID Spacing)')
                        else:
                            # =========================================================|
                            # [#] Set verification_message to be the current Situation.|
                            # =========================================================|
                            verification_messages_of_creating_old_horizontal_machine_program = [
                                "[u]C/Fren ID Spacing[/u] " +
                                "does NOT found, Please Enter the value :"]
                            # ==============================================================|
                            # [#] Add verification_message to the Confirmation Details list.|
                            # ==============================================================|
                            old_horizontal_program_confirmation_email_message_list.append(
                                "# It was need confirmation of :" + "\n" + (
                                    "**C/Fren ID Spacing** does NOT found, Please Enter the value"))
                            # ==================================================================|
                            # [#] title: set it to have the header and the verification_message.|
                            # ==================================================================|
                            self.title = '[color=0066ff]Confirmation Message[/color]' + '\n' + '\n' + \
                                         '[b][i][color=ffffff]' + \
                                         verification_messages_of_creating_old_horizontal_machine_program[0] + \
                                         '[/color][/i][/b]'
                            # ====================================================================|
                            # [#] Create the window_message that contain the Confirmation Details.|
                            # ====================================================================|
                            self.Dialog_BoxLayout = BoxLayout(height=30)
                            # ===========================================================|
                            # [#] Create the TextInput field to let user enter the value.|
                            # [#] (input_filter="float"): to accept Numeric input Only.  |
                            # ===========================================================|
                            self.cfren_ID_spacing_confirmation_text_field = TextInput(
                                hint_text="Enter Value", multiline=False, input_filter="float",
                                background_color=[60 / 255.0, 60 / 255.0, 60 / 255.0, 1],
                                foreground_color=[1, 1, 1, 1], size_hint_y=None, height=30)
                            # ===========================================================================|
                            # [#] Add the TextInput field created above to be part of the window_message.|
                            # ===========================================================================|
                            self.Dialog_BoxLayout.add_widget(self.cfren_ID_spacing_confirmation_text_field)
                            # =================================================================================|
                            # [#] Call Function of (need_confirmation) by pass the parameters that created:    |
                            #   [#] title: self.title                                                          |
                            #   [#] sub_function: self.enter_cfren_ID_spacing_value (it's the Function         |
                            #        created below to set the C/Fren ID Spacing).                              |
                            #   [#] dialog_type: "custom" (it used to indicate that user needs to enter value).|
                            #   [#] content: self.Dialog_BoxLayout (the window_message that have the Details). |
                            # =================================================================================|
                            need_confirmation_to_create_old_horizontal_machine_program(
                                self, self.title, self.enter_cfren_ID_spacing_value,
                                "custom", self.Dialog_BoxLayout)
                            print("C/Fren ID Spacing BEFORE User Input: ", cfren_ID_spacing)
                            return

                    # endregion <<<<================[Z_Value Of Top Of C-Fren Variable <VC164>]=================>>>>

                    # region <<<<================[Z_Value Of Bottom Of C-Fren Variable <VC165>]=================>>>>

                    substr = "VC165"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and
                            (cfren_ID_spacing != 0 and cfren_ID_spacing is not None)):

                        # =======================================================================================|
                        # [#] Use (if-else statements) to check if the C/Fren Cutter Width is missing, therefore |
                        #     it needs to ask the user to enter the value.                                       |
                        # =======================================================================================|
                        if (cfren_cutter_width != 0 and cfren_cutter_width is not None):
                            VC165_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                            print(VC165_variable_index)
                            print(pin_bore_program_lines_of_old_horizontal_machine[VC165_variable_index])

                            pin_bore_program_lines_of_old_horizontal_machine[VC165_variable_index] = (
                                    'VC165=-[VC164+' + format(cfren_cutter_width) + ']' +
                                    '  (zBottomCFREN IS zTopCFREN + LR CUTTER WIDTH)')
                        else:
                            # =========================================================|
                            # [#] Set verification_message to be the current Situation.|
                            # =========================================================|
                            verification_messages_of_creating_old_horizontal_machine_program = [
                                "[u]C/Fren Cutter Width[/u] " +
                                "does NOT found, Please Enter the value :"]
                            # ==============================================================|
                            # [#] Add verification_message to the Confirmation Details list.|
                            # ==============================================================|
                            old_horizontal_program_confirmation_email_message_list.append(
                                "# It was need confirmation of :" + "\n" + (
                                    "**C/Fren Cutter Width** does NOT found, Please Enter the value"))
                            # ==================================================================|
                            # [#] title: set it to have the header and the verification_message.|
                            # ==================================================================|
                            self.title = '[color=0066ff]Confirmation Message[/color]' + '\n' + '\n' + \
                                         '[b][i][color=ffffff]' + \
                                         verification_messages_of_creating_old_horizontal_machine_program[0] + \
                                         '[/color][/i][/b]'
                            # ====================================================================|
                            # [#] Create the window_message that contain the Confirmation Details.|
                            # ====================================================================|
                            self.Dialog_BoxLayout = BoxLayout(height=30)
                            # ===========================================================|
                            # [#] Create the TextInput field to let user enter the value.|
                            # [#] (input_filter="float"): to accept Numeric input Only.  |
                            # ===========================================================|
                            self.cfren_cutter_width_confirmation_text_field = TextInput(
                                hint_text="Enter Value", multiline=False, input_filter="float",
                                background_color=[60 / 255.0, 60 / 255.0, 60 / 255.0, 1],
                                foreground_color=[1, 1, 1, 1], size_hint_y=None, height=30)
                            # ===========================================================================|
                            # [#] Add the TextInput field created above to be part of the window_message.|
                            # ===========================================================================|
                            self.Dialog_BoxLayout.add_widget(self.cfren_cutter_width_confirmation_text_field)
                            # =================================================================================|
                            # [#] Call Function of (need_confirmation) by pass the parameters that created:    |
                            #   [#] title: self.title                                                          |
                            #   [#] sub_function: self.enter_cfren_cutter_width_value (it's the Function       |
                            #        created below to set the C/Fren Cutter Width).                            |
                            #   [#] dialog_type: "custom" (it used to indicate that user needs to enter value).|
                            #   [#] content: self.Dialog_BoxLayout (the window_message that have the Details). |
                            # =================================================================================|
                            need_confirmation_to_create_old_horizontal_machine_program(
                                self, self.title, self.enter_cfren_cutter_width_value,
                                "custom", self.Dialog_BoxLayout)
                            print("C/Fren Cutter Width BEFORE User Input: ", cfren_cutter_width)
                            return

                    # endregion <<<<================[Z_Value Of Bottom Of C-Fren Variable <VC165>]=================>>>>

                    # region <<<<================[C-Fren Cut Radius Variable <VC166>]=================>>>>

                    substr = "VC166"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and
                            (cfren_cutter_width != 0 and cfren_cutter_width is not None)):

                        # ===================================================================================|
                        # [#] Use (if-else statements) to check if the C/Fren Diameter is missing, therefore |
                        #     it needs to ask the user to enter the value.                                   |
                        # ===================================================================================|
                        if (cfren_diameter != 0 and cfren_diameter is not None):
                            VC166_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                            print(VC166_variable_index)
                            print(pin_bore_program_lines_of_old_horizontal_machine[VC166_variable_index])

                            pin_bore_program_lines_of_old_horizontal_machine[VC166_variable_index] = (
                                    'VC166=[[' + format(cfren_diameter) + '-' + format(cfren_tool_diameter) +
                                    ']/2]' + '  (CFCutRadius)')
                        else:
                            # =========================================================|
                            # [#] Set verification_message to be the current Situation.|
                            # =========================================================|
                            verification_messages_of_creating_old_horizontal_machine_program = [
                                "[u]C/Fren Diameter[/u] " +
                                "does NOT found, Please Enter the value :"]
                            # ==============================================================|
                            # [#] Add verification_message to the Confirmation Details list.|
                            # ==============================================================|
                            old_horizontal_program_confirmation_email_message_list.append(
                                "# It was need confirmation of :" + "\n" + (
                                    "**C/Fren Diameter** does NOT found, Please Enter the value"))
                            # ==================================================================|
                            # [#] title: set it to have the header and the verification_message.|
                            # ==================================================================|
                            self.title = '[color=0066ff]Confirmation Message[/color]' + '\n' + '\n' + \
                                         '[b][i][color=ffffff]' + \
                                         verification_messages_of_creating_old_horizontal_machine_program[0] + \
                                         '[/color][/i][/b]'
                            # ====================================================================|
                            # [#] Create the window_message that contain the Confirmation Details.|
                            # ====================================================================|
                            self.Dialog_BoxLayout = BoxLayout(height=30)
                            # ===========================================================|
                            # [#] Create the TextInput field to let user enter the value.|
                            # [#] (input_filter="float"): to accept Numeric input Only.  |
                            # ===========================================================|
                            self.cfren_diameter_confirmation_text_field = TextInput(
                                hint_text="Enter Value", multiline=False, input_filter="float",
                                background_color=[60 / 255.0, 60 / 255.0, 60 / 255.0, 1],
                                foreground_color=[1, 1, 1, 1], size_hint_y=None, height=30)
                            # ===========================================================================|
                            # [#] Add the TextInput field created above to be part of the window_message.|
                            # ===========================================================================|
                            self.Dialog_BoxLayout.add_widget(self.cfren_diameter_confirmation_text_field)
                            # =================================================================================|
                            # [#] Call Function of (need_confirmation) by pass the parameters that created:    |
                            #   [#] title: self.title                                                          |
                            #   [#] sub_function: self.enter_cfren_diameter_value (it's the Function           |
                            #        created below to set the C/Fren Diameter).                                |
                            #   [#] dialog_type: "custom" (it used to indicate that user needs to enter value).|
                            #   [#] content: self.Dialog_BoxLayout (the window_message that have the Details). |
                            # =================================================================================|
                            need_confirmation_to_create_old_horizontal_machine_program(
                                self, self.title, self.enter_cfren_diameter_value, "custom",
                                self.Dialog_BoxLayout)
                            print("C/Fren Diameter BEFORE User Input: ", cfren_diameter)
                            return

                    # endregion <<<<================[C-Fren Cut Radius Variable <VC166>]=================>>>>

                    # region <<<<================[C-Fren Cut Offset Variable <VC167>]=================>>>>
                    # ===================================================================================|
                    # [#] No need to do anything here, it will just use how it is in Horizontal template.|
                    # ===================================================================================|
                    # endregion <<<<================[C-Fren Cut Offset Variable <VC167>]=================>>>>

                    # region <<<<===============[Double Oil Hole Slot ID Spacing Variable <VC168>]================>>>>

                    substr = "VC168"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and
                            (double_oil_hole_slot_ID_spacing != 0 and double_oil_hole_slot_ID_spacing is not None)
                            and pin_hole_diameter >= 0.901):
                        VC168_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC168_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC168_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC168_variable_index] = (
                                'VC168=' + format(double_oil_hole_slot_ID_spacing) + '  (DOHS_ID_Spacing)')

                    # endregion <<<<=============[Double Oil Hole Slot ID Spacing Variable <VC168>]===============>>>>

                    # region <<<<===========[Notch Calculation To Use On Variables (VC172 AND VC173)]============>>>>

                    # =========================================================================================|
                    # [#] It was figured out by drawing triangle in SolidWorks and using Trigonometric math    |
                    #     and the Pythagorean theorem.                                                         |
                    # [#] Check the Word file of the Calculation (located on:..........)  for more Explanation.|
                    # =========================================================================================|
                    if (((notch_angle_first_location != 0 and notch_angle_first_location is not None)) and
                            offset_amount is not None):
                        # =========================================================================================|
                        # [#] Use (if statement) to make sure the Offset Direction is Set to use it in calculation.|
                        # =========================================================================================|
                        if (offset_amount == 0):
                            offset_direction_for_math = 1
                        elif (offset_direction == "OFFSET EACH WAY" or offset_direction == "OFFSET To0"):
                            offset_direction_for_math = 1
                        elif (offset_direction == "OFFSET To180"):
                            offset_direction_for_math = -1
                        else:
                            # ==============================================================|
                            # [#] To make sure the code work if Offset Direction is missing.|
                            # ==============================================================|
                            offset_direction_for_math = 1

                        # ===================================================================|
                        # [#] Use (math.radians()) because having the Notch Angle in Degrees.|
                        # ===================================================================|
                        Y_value_for_notch_math = float(
                            round((float(pin_hole_diameter) / 2) * (
                                math.sin(math.radians((180 - float(notch_angle_first_location))))), 4))

                        X_value_for_notch_math = float(
                            round((float(pin_hole_diameter) / 2) * (
                                math.cos(math.radians((180 - float(notch_angle_first_location))))), 4))

                        # ====================================================================|
                        # [#] Make Variables 'global' to be able to use it in other functions.|
                        # ====================================================================|
                        global Y_distance_from_origin_to_circlip_notch
                        Y_distance_from_origin_to_circlip_notch = format(
                            float(Y_value_for_notch_math) + ((offset_amount) * offset_direction_for_math), '.4f')

                        global X_distance_from_origin_to_circlip_notch
                        X_distance_from_origin_to_circlip_notch = format(
                            float(X_distance_from_origin_to_pin_center) + float(X_value_for_notch_math), '.4f')

                    # endregion <<<<==========[Notch Calculation To Use On Variables (VC172 AND VC173)]===========>>>>

                    # region <<<<===========[X_Distance From Origin To Circlip Notch Variable <VC172>]============>>>>

                    substr = "VC172"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and
                            (notch_angle_first_location != 0 and notch_angle_first_location is not None)):
                        VC172_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC172_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC172_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC172_variable_index] = (
                                'VC172=-' + format(X_distance_from_origin_to_circlip_notch) + '  (xCirclipNotch)')

                    # endregion <<<<=========[X_Distance From Origin To Circlip Notch Variable <VC172>]==========>>>>

                    # region <<<<===========[Y_Distance From Origin To Circlip Notch Variable <VC173>]============>>>>

                    substr = "VC173"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and
                            (notch_angle_first_location != 0 and notch_angle_first_location is not None)):
                        global VC173_variable_index
                        VC173_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC173_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC173_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC173_variable_index] = (
                                'VC173=' + format(Y_distance_from_origin_to_circlip_notch) + '  (yCirclipNotch)')

                    # endregion <<<<=========[Y_Distance From Origin To Circlip Notch Variable <VC173>]==========>>>>

                    # region <<<<==============[H-Slots OD Spacing Variable <VC175>]===============>>>>

                    substr = "VC175"
                    horizontal_program_line_index = line.find(substr)

                    # (and horizontal_slots_arc_diameter != 0.375) : ***JUST FOR NOW*** CHECK IF JOB HAS
                    # H-SLOTS WITH RADIUS NOT EQUAL 0.375, BECAUSE H-SLOTS WITH 0.375 RADIUS CONSIDER
                    # (HORIZONTAL SLOTS STRAIGHT_THROUGH), WHICH NEED TO ADJUST OTHER VARIABLES

                    if (horizontal_program_line_index == 0 and
                            (horizontal_slots_diameter_depth != 0 and horizontal_slots_diameter_depth is not None)
                            and horizontal_slots_arc_diameter != 0.375):
                        VC175_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC175_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC175_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC175_variable_index] = (
                                'VC175=' + format(horizontal_slots_OD_spacing) + '  (HSlot_ID_Spacing)')

                    # endregion <<<<==============[H-Slots OD Spacing Variable <VC175>]===============>>>>

                    # region <<<<==============[i_Start Horizontal Slot Variable <VC176>]===============>>>>

                    substr = "VC176"
                    horizontal_program_line_index = line.find(substr)

                    # (and horizontal_slots_arc_diameter != 0.375) : ***JUST FOR NOW*** CHECK IF JOB HAS
                    # H-SLOTS WITH RADIUS NOT EQUAL 0.375, BECAUSE H-SLOTS WITH 0.375 RADIUS CONSIDER
                    # (HORIZONTAL SLOTS STRAIGHT_THROUGH), WHICH NEED TO ADJUST OTHER VARIABLES

                    if (horizontal_program_line_index == 0 and
                            (horizontal_slots_diameter_depth != 0 and horizontal_slots_diameter_depth is not None)
                            and horizontal_slots_arc_diameter != 0.375):
                        VC176_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC176_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC176_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC176_variable_index] = (
                                'VC176=' + format(i_start_horizontal_slot) + '  (iStartHSlot)')

                    # endregion <<<<==============[i_Start Horizontal Slot Variable <VC176>]===============>>>>

                    # region <<<<==============[j_Start Horizontal Slot Variable <VC177>]===============>>>>

                    substr = "VC177"
                    horizontal_program_line_index = line.find(substr)

                    # (and horizontal_slots_arc_diameter != 0.375) : ***JUST FOR NOW*** CHECK IF JOB HAS
                    # H-SLOTS WITH RADIUS NOT EQUAL 0.375, BECAUSE H-SLOTS WITH 0.375 RADIUS CONSIDER
                    # (HORIZONTAL SLOTS STRAIGHT_THROUGH), WHICH NEED TO ADJUST OTHER VARIABLES

                    if (horizontal_program_line_index == 0 and
                            (horizontal_slots_diameter_depth != 0 and horizontal_slots_diameter_depth is not None)
                            and horizontal_slots_arc_diameter != 0.375):
                        VC177_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC177_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC177_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC177_variable_index] = (
                                'VC177=' + format(j_start_horizontal_slot) + '  (jStartHSlot)')

                    # endregion <<<<==============[j_Start Horizontal Slot Variable <VC177>]===============>>>>

                    # region <<<<==============[Horizontal Slot Radius Variable <VC178>]===============>>>>

                    substr = "VC178"
                    horizontal_program_line_index = line.find(substr)

                    # (and horizontal_slots_arc_diameter != 0.375) : ***JUST FOR NOW*** CHECK IF JOB HAS
                    # H-SLOTS WITH RADIUS NOT EQUAL 0.375, BECAUSE H-SLOTS WITH 0.375 RADIUS CONSIDER
                    # (HORIZONTAL SLOTS STRAIGHT_THROUGH), WHICH NEED TO ADJUST OTHER VARIABLES

                    if (horizontal_program_line_index == 0 and
                            (horizontal_slots_diameter_depth != 0 and horizontal_slots_diameter_depth is not None)
                            and horizontal_slots_arc_diameter != 0.375):
                        VC178_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC178_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC178_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC178_variable_index] = (
                                'VC178=' + format(horizontal_slot_radius) + '  (HSlotRadius)')

                    # endregion <<<<==============[Horizontal Slot Radius Variable <VC178>]===============>>>>

                    # region <<<<==============[Counterbore Ledge Diameter Variable <VC179>]===============>>>>

                    substr = "VC179"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and
                            ledge_counterbore_diameter != 0 and ledge_counterbore_diameter is not None):
                        VC179_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(VC179_variable_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[VC179_variable_index])

                        pin_bore_program_lines_of_old_horizontal_machine[VC179_variable_index] = (
                                'VC179=' + format(ledge_counterbore_diameter) + '  (CounterboreLedgeDiameter)')

                    # endregion <<<<==============[Counterbore Ledge Diameter Variable <VC179>]===============>>>>

                    # region <=[MAYBE NEED TO ADD VARIABLE OF DISTANCE OF COUNTERBORE TO LOCKRING TO H-TEMPLATE]=>

                    # endregion <=[MAYBE NEED TO ADD VARIABLE OF DISTANCE OF COUNTERBORE TO LOCKRING TO H-TEMPLATE]=>

                    # region <<<<=[X Distance 375 Slots From Center Of Bore To Center Of H-Slot Variable <VC183>]=>>>>
                        # =====================================================|
                        # [#] Needs to figure out the math we want to use here.|
                        # =====================================================|
                    # endregion <<<<[X Distance 375 Slots From Center Of Bore To Center Of H-Slot Variable <VC183>]>>>>

                    # region <<<<=[Y Distance 375 Slots From Center Of Bore To Center Of H-Slot Variable <VC184>]=>>>>
                        # =====================================================|
                        # [#] Needs to figure out the math we want to use here.|
                        # =====================================================|
                    # endregion <<<<[y Distance 375 Slots From Center Of Bore To Center Of H-Slot Variable <VC184>]>>>>

                    # region <<<<=========[Forging Or Finished Boss Width For 375_Slots Variable <VC185>]==========>>>>

                    substr = "VC185"
                    horizontal_program_line_index = line.find(substr)
                    # (and horizontal_slots_arc_diameter == 0.375) :
                    # CHECK IF JOB HAS (HORIZONTAL SLOTS STRAIGHT_THROUGH)
                    if (horizontal_program_line_index == 0 and horizontal_slots_arc_diameter == 0.375):
                        # ==========================================================================================|
                        # [#] Use (if-else statements) because Some Forging (Especially JE Forging) is missing some |
                        #     Information(dimension), therefore it needs to check if the dimension is missing to ask|
                        #     the user to enter the value.                                                          |
                        # ==========================================================================================|
                        if (forging_inside_boss_spacing != 0 and forging_inside_boss_spacing is not None):
                            VC185_variable_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                            print(VC185_variable_index)
                            print(pin_bore_program_lines_of_old_horizontal_machine[VC185_variable_index])

                            pin_bore_program_lines_of_old_horizontal_machine[VC185_variable_index] = (
                                    'VC185=' + format(forging_inside_boss_spacing) +
                                    '  (ForgingOrFinishedBossWidthFor375Slots)')
                        else:
                            # =========================================================|
                            # [#] Set verification_message to be the current Situation.|
                            # =========================================================|
                            verification_messages_of_creating_old_horizontal_machine_program = [
                                "[u]Forging Inside Boss Spacing[/u] " +
                                "does NOT found, Please take a look to Forging Info or Model and Enter the value :"]
                            # ==============================================================|
                            # [#] Add verification_message to the Confirmation Details list.|
                            # ==============================================================|
                            old_horizontal_program_confirmation_email_message_list.append(
                                "# It was need confirmation of :" + "\n" + (
                                    "**Forging Inside Boss Spacing** does NOT found, "
                                    "Please take a look to Forging Info or Model and Enter the value"))
                            # ==================================================================|
                            # [#] title: set it to have the header and the verification_message.|
                            # ==================================================================|
                            self.title = '[color=0066ff]Confirmation Message[/color]' + '\n' + '\n' + \
                                         '[b][i][color=ffffff]' + \
                                         verification_messages_of_creating_old_horizontal_machine_program[0] + \
                                         '[/color][/i][/b]'
                            # ====================================================================|
                            # [#] Create the window_message that contain the Confirmation Details.|
                            # ====================================================================|
                            self.Dialog_BoxLayout = BoxLayout(height=30)
                            # ===========================================================|
                            # [#] Create the TextInput field to let user enter the value.|
                            # [#] (input_filter="float"): to accept Numeric input Only.  |
                            # ===========================================================|
                            self.forging_inside_boss_spacing_confirmation_text_field = TextInput(
                                hint_text="Enter Value", multiline=False, input_filter="float",
                                background_color=[60 / 255.0, 60 / 255.0, 60 / 255.0, 1],
                                foreground_color=[1, 1, 1, 1], size_hint_y=None, height=30)
                            # ===========================================================================|
                            # [#] Add the TextInput field created above to be part of the window_message.|
                            # ===========================================================================|
                            self.Dialog_BoxLayout.add_widget(self.forging_inside_boss_spacing_confirmation_text_field)
                            # ===================================================================================|
                            # [#] Call Function of (need_confirmation) by pass the parameters that created:      |
                            #   [#] title: self.title                                                            |
                            #   [#] sub_function: self.enter_forging_inside_boss_spacing_value (it's the Function|
                            #        created below to set the Forging Inside Boss Spacing).                      |
                            #   [#] dialog_type: "custom" (it used to indicate that user needs to enter value).  |
                            #   [#] content: self.Dialog_BoxLayout (the window_message that have the Details).   |
                            # ===================================================================================|
                            need_confirmation_to_create_old_horizontal_machine_program(
                                self, self.title, self.enter_forging_inside_boss_spacing_value,
                                "custom", self.Dialog_BoxLayout)
                            print("Forging Inside Boss Spacing BEFORE User Input: ", forging_inside_boss_spacing)
                            return

                    # endregion <<<<=======[Forging Or Finished Boss Width For 375_Slots Variable <VC185>]========>>>>

                    # region <<<<=========[Probe Program Call <CALL OXXXX>]==========>>>>

                    substr = "CALL OXXXX"
                    horizontal_program_line_index = line.find(substr)
                    if (horizontal_program_line_index == 0 and
                            (forging_number != "" and forging_number is not None)
                            and probe_program_of_old_horizontal_machine is not None):
                        probe_program_index = pin_bore_program_lines_of_old_horizontal_machine.index(line)
                        print(probe_program_index)
                        print(pin_bore_program_lines_of_old_horizontal_machine[probe_program_index])
                        pin_bore_program_lines_of_old_horizontal_machine[probe_program_index] = (
                                'CALL ' + format(probe_program_of_old_horizontal_machine))

                    # endregion <<<<=========[Probe Program Call <CALL OXXXX>]==========>>>>

                    # endregion <<<<================[Dimensional Variables]=================>>>>

        # =================================================================================================|
        # [#] Use (except) Block to Handle any Error may occur when Accessing or Finding Files and Folders.|
        # [#] (IOError) is error of not find the file of template.                                         |
        # [#] Call (failed_to_create_old_horizontal_machine_program) Function to warn the user.            |
        # [#] Use (return) to stop the function (APP as a result) from executing the rest of the code.     |
        # =================================================================================================|
        except IOError as error:
            fail_messages_of_creating_old_horizontal_machine_program.append(
                "Failed to Find, Load, or Access" + '[b][u][color=ffffff] Horizontal Template [/color][/u][/b]' +
                "File to Create the Program." + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]' + str(
                    error) + '[/color]' + "\n" + "Double Check Network, and File Location." + "\n")
            email_messages_of_creating_old_horizontal_machine_program.append(
                "Failed to Find, Load, or Access Horizontal Template File to Create the Program." + "\n" +
                "Double Check Network, and File Location." + "\n")
            failed_to_create_old_horizontal_machine_program(self)
            return

        # endregion <<<<============================[Old Horizontal For-Loop Template]============================>>>>

        # region  <<<<========================[Old Horizontal Checking Create Program]==========================>>>>

        # =========================================================================================================|
        # [#] Use (if-else statements) to check Status of creating the Program:                                    |
        #    [#] If Lists of FailMessageOfCreatingProgram and VerificationMessageOfCreatingProgram have no elements|
        #        (which indicate there are nothing needs to verify or to prevent to create the program), and the   |
        #        TextInput Field of JobNumber has a value , then the APP will try to create the program.           |
        #    [#] If List of VerificationMessageOfCreatingProgram have some elements (which indicate there are      |
        #        some Confirmation details needs the user to finalize to create the program), then the APP will NOT|
        #        create the program and will print message of the Confirmation details that needed. (this condition|
        #        used just as a safety check, Code have been written to get all Confirmation details finalized by  |
        #        User before reach to this line).                                                                  |
        #    [#] If TextInput Field of JobNumber has NO input, then the APP will NOT create the program and will   |
        #        Warn the user to enter the Job Number.                                                            |
        # =========================================================================================================|
        if (fail_messages_of_creating_old_horizontal_machine_program == [] and
                verification_messages_of_creating_old_horizontal_machine_program == [] and
                self.ids["JobNumberForOldHorizontalMachine"].text != ""):

            # region  <<<<========================[Exceptional Programs Cases <Notes>]==========================>>>>

            # ======================================================================================================|
            # [#] Wiseco Brand Built its Reputation on the flexibility of Customize its Pistons to match Customers' |
            #     needs and expectations, because of that there are some Situations and Cases that Needs the User to|
            #     customize the Template or modify the program manually.                                            |
            # [#] The App will create the basic program in the Original folder ONLY with a message of the exceptions|
            #     (on the Top of Program) that need attention from the User to modify the program manually.         |
            # ======================================================================================================|

            # ================================================================================================|
            # [#] Note<2,1>                                                                                   |
            # [#] In Case the Job has H-Slots and PressureFedHoles, the User needs to Add the H-Slots manually|
            #     to make them intersect with the PressureFedHoles (Ex: SUM1109273780-5).                     |
            # Maybe needs to add to this condition to check the H-Slots Numbers is not <1>
            # ================================================================================================|
            if ((horizontal_slots_diameter_depth != 0 and horizontal_slots_diameter_depth is not None)
                    and pressure_fed_holes_availability_status == 1):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, IT HAS H-SLOTS AND PRESSURE FED OIL HOLES FEATURES, NEEDS TO ADD HORIZ SLOTS "
                    "MANUALLY ***DELETE THIS NOTE BEFORE YOU SAVE IT***" "\n" +
                    "(****MANUAL HORIZ SLOTS - WATCH CLOSELY****)" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # =================================================================================================|
            # [#] Note<2,2>                                                                                    |
            # [#] In Case the Job has One H-Slot Only per side, the User needs to delete One of H-Slot manually|
            #     from the Template Logic (Ex: WD-13768).                                                      |
            # =================================================================================================|
            if ((horizontal_slots_diameter_depth != 0 and horizontal_slots_diameter_depth is not None) and
                    (horiz_piston_pin_slots_notes is not None and horiz_piston_pin_slots_notes != "") and
                    (horiz_piston_pin_slots_notes.find('1 SLOT PER SIDE') != -1)):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, DELETE ONE OF HORIZ SLOTS MANUALLY (SEE MODEL TO CHECK IT'S AT 90 OR 270)  "
                    "***DELETE THIS NOTE BEFORE YOU SAVE IT***" "\n" +
                    "(ONE HORIZONTAL SLOT ONLY PER SIDE)" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # ===============================================================================================|
            # [#] Note<2,3>                                                                                  |
            # [#] In Case the H-Slot OD Spacing is BiggerThan <1.0 inch>(length of the tool), the User needs |
            #     to check if it needs to add Extra Pass Manually (Ex:  WD-16033, WD-12138, WD-12143).       |
            # [#] To get this Calculation, Subtract (ForgingInsideBossSpacing/2) from the H-Slots_OD_Spacing.|
            # ===============================================================================================|
            if ((horizontal_slots_diameter_depth != 0 and horizontal_slots_diameter_depth is not None) and
                    (horizontal_slots_OD_spacing != 0 and horizontal_slots_OD_spacing is not None and
                     horizontal_slots_OD_spacing != "") and
                    (forging_inside_boss_spacing != 0 and forging_inside_boss_spacing is not None) and
                    (((horizontal_slots_OD_spacing - forging_inside_boss_spacing)/2)) > 1.00):
                print("HORIZ SLOTS OD SPACING EACH SIDE: ",
                      format((horizontal_slots_OD_spacing - forging_inside_boss_spacing)/2, '.4f'))
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, IT SEEMS HORIZ SLOTS OD SPACING EACH SIDE IS LONGER THAN 1 INCH "
                    "(LENGTH OF THE TOOL), "
                    "CHECK IF IT IS NEED TO ADD EXTRA PASS MANUALLY. ***DELETE THIS NOTE BEFORE YOU SAVE IT***" "\n" +
                    "(****MANUAL HORIZ SLOTS - WATCH CLOSELY****)" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # ==============================================================================================|
            # [#] Note<2,4>                                                                                 |
            # [#] In Case the Job has PinHole Size with NO H-Slot Numbers in the Excel sheet, the User needs|
            #     to find and add the H-Slots Numbers or Add H-Slots manually.                              |
            # ==============================================================================================|
            if ((horizontal_slots_diameter_depth != 0 and horizontal_slots_diameter_depth is not None) and
                    (i_start_horizontal_slot == "" or j_start_horizontal_slot == "" or horizontal_slot_radius == "")):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, CAN'T FIND HORIZ SLOTS NUMBERS (i,j,Radius) FOR THE PIN SIZE, "
                    "ADD THE NUMBERS OR MAKE HORIZ SLOTS MANUALLY "
                    "***DELETE THIS NOTE BEFORE YOU SAVE IT***" "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # ======================================================================================================|
            # [#] Note<2,5>                                                                                         |
            # [#] In Case the Job has Semi-C/Fren, the User needs to add Semi-C/Fren Values manually (Ex: WD-07676).|
            # ======================================================================================================|
            if (semi_cfren_ID_spacing != 0 and semi_cfren_ID_spacing is not None):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, ADD Semi-C/Fren VALUES(VC122,VC134,VC164,VC165,VC166,and VC167)MANUALLY "
                    "***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # ==================================================================================================|
            # [#] Note<2,6>                                                                                     |
            # [#] In Case the Job has different tool width for Lockring and C/Fren, the User needs to add C/Fren|
            #     tool manually (Ex: 12947M10000, 13157M10300).                                                 |
            # ==================================================================================================|
            if ((lock_ring_cutter_width != cfren_cutter_width) and
                    (cfren_ID_spacing != 0 and cfren_ID_spacing is not None) and
                    (lock_ring_cutter_width != 0 and lock_ring_cutter_width is not None)):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, Modify the Template MANUALLY to be able to use different "
                    "tools for Lockring and C/Fren  "
                    "***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # ====================================================================================================|
            # [#] Note<2,7>                                                                                       |
            # [#] In Case PinHoleDiameter didn't detect, the User needs to work with Engineering to fix the issue.|
            #     (this condition used just as a safety check, Code have been written to avoid this issue before  |
            #     reach to this line).                                                                            |
            # ====================================================================================================|
            if (pin_hole_diameter == 0.0 or pin_hole_diameter is None):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, PIN HOLE DIAMETER CAN'T BE EQUAL 0 OR NOTHING, DOUBLE CHECK JOB INFO AND TRY "
                    "TO FIX THE ISSUE, ***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # ==================================================================================================|
            # [#] Note<2,8>                                                                                     |
            # [#] In Case PinHoleDiameter is BiggerThan <1.095>, the User needs to use RoughBoreTool with bigger|
            #     diameter, or use counterbore cut on both sides (Ex: WD-13960, AW-07578, WD-15506).            |
            # ==================================================================================================|
            if (pin_hole_diameter > 1.095):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, USE ROUGH BORE TOOL WITH BIGGER DIAMETER, OR USE COUNTERBORE CUT ON BOTH SIDES  "
                    "***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # =================================================================================================|
            # [#] Note<2,9>                                                                                    |
            # [#] In Case PinHoleDiameter is SmallerThan <0.4710>, the User needs to check if the machine SetUp|
            #     to run job with Small Uncommon PinHoleDiameter size.                                         |
            # =================================================================================================|
            if (pin_hole_diameter < 0.4710):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, SMALL UNCOMMON PIN HOLE DIAMETER SIZE " + pin_hole_diameter +
                    ", DOUBLE CHECK IF WE ARE ABLE TO RUN THIS JOB. ***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # =================================================================================================|
            # [#] Note<2,10>                                                                                   |
            # [#] In Case RoughBoreToolNumber still equal <0> (which indicates that can't find RoughBoreTool   |
            #     that fit the FinishPinHoleDiameter size), the User needs to check if can use Swap Tool (T60).|
            # =================================================================================================|
            if (rough_bore_tool_number == 0):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, CAN'T FIND RoughBoreTool THAT FIT THE FinishPinHoleDiameter SIZE, "
                    "MAYBE WE CAN USE SWAP TOOL T60 IF APPLICABLE  ***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # ================================================================================================|
            # [#] Note<2,11>                                                                                  |
            # [#] In Case PistonOverallLength is LongerThan <4.30>, the User needs to work on Machine Rotation|
            #     manually to avoid crashing the Machine (Ex: WD-13497).                                      |
            # ================================================================================================|
            if (piston_overall_length >= 4.300):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, LONG PART THAT NEEDS TO WORK ON MACHINE ROTATION (M15/M16)   "
                    "***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # =================================================================================================|
            # [#] Note<2,12>                                                                                   |
            # [#] In Case the Job has LedgeCounterbore, the User needs to add LedgeCounterbore values manually.|
            # [#] The Reason behind asking the User to add the LedgeCounterbore values manually because it is  |
            #     hard to get Counterbore diameter easily from the DataBase because it is always stored in     |
            #     comments (there is no specific location can access to get the value from the DataBase).      |
            # =================================================================================================|
            if ((ledge_counterbore_diameter != 0 and ledge_counterbore_diameter is not None and
                 ledge_counterbore_diameter == "")):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, JOB SEEMS NEED LedgeCounterbore,IF SO,ADD COUNTERBORE VALUES "
                    "(VC126 and VC179) MANUALLY "
                    "AND ID SPACING IF NECESSARY ***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # =================================================================================================|
            # [#] Note<2,13>                                                                                   |
            # [#] In Case the Job has LedgeCounterbore with NO LockRing, the User needs to add LedgeCounterbore|
            #     Diameter and ID Spacing manually (Ex: WD-13308).                                             |
            # =================================================================================================|
            if ((ledge_counterbore_diameter != 0 and ledge_counterbore_diameter is not None and
                 ledge_counterbore_diameter == "") and (lock_ring_ID_spacing == 0 or lock_ring_ID_spacing is None)):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, JOB SEEMS NEED LedgeCounterbore ID SPACING BECAUSE HAS NO LOCKRING,IF SO,"
                    "NEEDS TO ADD IT MANUALLY  ***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # =================================================================================================|
            # [#] Note<2,14>                                                                                   |
            # [#] In Case the Job has Forged forging, the User needs to add VC85 value of the forging and      |
            #     distance from Pilot to Top of Forging manually (Ex: WIS-10079).                              |
            # [#] The 'forged_forging_list' (that set earlier in [Load Horizontal Sheets Function] section)    |
            #     contains all Forged forging.                                                                 |
            # =================================================================================================|
            if ((forging_number != "" and forging_number is not None) and (forging_number in forged_forging_list)):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, NEEDS TO ADD BELOW VC85 VALUE OF THE FORGING AND DISTANCE FROM PILOT TO "
                    "TOP OF FORGING ***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n" + "(TRY VC85 = ???)" + "\n" +
                    "(DISTANCE FROM PILOT TO TOP OF FORGING SHOULD MEASURE = ???)" + "\n" +
                    "(MAY HAVE TO ADJUST FOR PILOT TO TOP OF FORGING)" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # =====================================================================================================|
            # [#] Note<2,15>                                                                                       |
            # [#] In Case the PinHoleDiameter size doesn't include in DOHS Logic in the template, the User needs   |
            #     to figure out the DOHS Numbers and add them to Template Logic manually.                          |
            # [#] The 'double_oil_hole_slots_pin_sizes_list' (that set earlier in [Load Horizontal Sheets Function]|
            #     section) contains all PinHoleDiameter sizes that used in Template Logic.                         |
            # =====================================================================================================|
            if ((pin_hole_diameter != 0 and pin_hole_diameter != "" and pin_hole_diameter is not None and
                 pin_hole_diameter >= 0.901) and (double_oil_hole_slot_ID_spacing != 0 and
                                                  double_oil_hole_slot_ID_spacing is not None) and
                    (pin_hole_diameter not in double_oil_hole_slots_pin_sizes_list)):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, DOHS LOGIC DOESN'T INCLUDE THIS PIN SIZE, NEEDS TO FIGURE OUT DOHS NUMBERS, "
                    "AND ADD THEM TO LOGIC MANUALLY  ***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n")
                notes_for_old_horizontal_machine_program_index += 1


            #  Add same Check above but for H-Slot

            # ====================================================================================================|
            # [#] Note<2,16>                                                                                      |
            # [#] In Case PilotBoreDepth didn't detect, the User needs to work with Engineering to fix the issue. |
            #     (this condition used just as a safety check, Code have been written to avoid this issue before  |
            #     reach to this line).                                                                            |
            # ====================================================================================================|
            if (pilot_bore_depth == 0 or pilot_bore_depth is None):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, PILOT BORE DEPTH CAN'T BE EQUAL 0 OR NOTHING, DOUBLE CHECK JOB INFO AND TRY "
                    "TO FIX THE ISSUE, ***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # ====================================================================================================|
            # [#] Note<2,17>                                                                                      |
            # [#] In Case X_DistanceFromOriginToPinCenter didn't calculate correctly or it's GraterThan <-0.3>,   |
            #     the User needs to work with Engineering to fix the issue.                                       |
            # ====================================================================================================|
            if (X_distance_from_origin_to_pin_center == 0 or (
                    float(float(X_distance_from_origin_to_pin_center) * (-1)) > (-0.3))):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, xPinCenter CAN'T BE EQUAL 0 OR BIGGER THAN -.3, DOUBLE CHECK JOB INFO AND "
                    "TRY TO FIX THE ISSUE,  ***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # ====================================================================================================|
            # [#] Note<2,18>                                                                                      |
            # [#] In Case LedgeToolDiameter equal <0>, the User needs to work with Engineering to fix the issue.  |
            # ====================================================================================================|
            if (ledge_tool_diameter == 0):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, Ledge Tool Diameter CAN'T BE EQUAL 0, DOUBLE CHECK JOB INFO AND "
                    "TRY TO FIX THE ISSUE,  ***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # ====================================================================================================|
            # [#] Note<2,19>                                                                                      |
            # [#] In Case LockRingDiameter is SmallerThan PinHoleDiameter, the User needs to work with Engineering|
            #     to fix the issue.                                                                               |
            # ====================================================================================================|
            if ((lock_ring_cutter_width != 0 and lock_ring_cutter_width is not None) and
                    (lock_ring_diameter <= pin_hole_diameter)):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, Lock Ring Diameter CAN'T BE SMALLER THAN Pin Hole Diameter, DOUBLE CHECK JOB INFO "
                    "AND TRY TO FIX THE ISSUE,  ***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # ==================================================================================================|
            # [#] Note<2,20>                                                                                    |
            # [#] In Case C/FrenDiameter is SmallerThan PinHoleDiameter, the User needs to work with Engineering|
            #     to fix the issue.                                                                             |
            # ==================================================================================================|
            if((cfren_cutter_width != 0 and cfren_cutter_width is not None) and
                    (cfren_diameter <= pin_hole_diameter)):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, C/Fren Diameter CAN'T BE SMALLER THAN Pin Hole Diameter, DOUBLE CHECK JOB INFO "
                    "AND TRY TO FIX THE ISSUE,  ***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # ==============================================================================================|
            # [#] Note<2,21>                                                                                |
            # [#] In Case the Job has CirclipNotch with NO LockRing, the User needs to work with Engineering|
            #     to fix the issue.                                                                         |
            # ==============================================================================================|
            if ((notch_angle_first_location != 0 and notch_angle_first_location is not None) and
                    (notch_angle_second_location != 0 and notch_angle_second_location is not None) and
                    (lock_ring_cutter_width == 0 or lock_ring_cutter_width is None)):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, SHOULD NOT HAVE CirclipNotch WHILE HAVE NO LockRing , DOUBLE CHECK JOB INFO "
                    "AND TRY TO FIX THE ISSUE,  ***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # =================================================================================================|
            # [#] Note<2,22>                                                                                   |
            # [#] In Case X_DistanceFromOriginToCirclipNotch is GraterThan <-0.3>, the User needs to work with |
            #     Engineering to fix the issue.                                                                |
            # =================================================================================================|
            if ((X_distance_from_origin_to_circlip_notch != 0 and X_distance_from_origin_to_circlip_notch is not None)
                    and (float(float(X_distance_from_origin_to_circlip_notch) * (-1)) > (-0.3))):
                notes_for_old_horizontal_machine_program.insert(
                    notes_for_old_horizontal_machine_program_index,
                    "NOT READY YET, xCirclipNotch CAN'T BE BIGGER THAN -.3, DOUBLE CHECK JOB INFO AND "
                    "TRY TO FIX THE ISSUE,  ***DELETE THIS NOTE BEFORE YOU SAVE IT***" + "\n")
                notes_for_old_horizontal_machine_program_index += 1

            # endregion  <<<<========================[Exceptional Programs Cases <Notes>]==========================>>>>

            # ==================================================================================|
            # [#] Call the Function to create Horizontal PinBore program in the Original folder.|
            # ==================================================================================|
            create_old_horizontal_machine_program_in_original_folder(self)

        # ======================================================================================================|
        # [#] If List of VerificationMessageOfCreatingProgram have some elements (which indicate there are      |
        #     some Confirmation details needs the user to finalize to create the program), then the APP will NOT|
        #     create the program and will print message of the Confirmation details that needed. (this condition|
        #     used just as a safety check, Code have been written to get all Confirmation details finalized by  |
        #     User before reach to this line).                                                                  |
        # ======================================================================================================|
        elif (verification_messages_of_creating_old_horizontal_machine_program != []):
            print("verification_messages_of_creating_old_horizontal_machine_program STILL HAVE SOMETHING: ",
                  verification_messages_of_creating_old_horizontal_machine_program)

        # ===================================================================================================|
        # [#] If TextInput Field of JobNumber has NO input, then the APP will NOT create the program and will|
        #     Warn the user to enter the Job Number.                                                         |
        # ===================================================================================================|
        elif (self.ids["JobNumberForOldHorizontalMachine"].text == ""):
            fail_messages_of_creating_old_horizontal_machine_program.append("Please Enter Job Number." + "\n")
            failed_to_create_old_horizontal_machine_program(self)
            return

        # endregion <<<<========================[Old Horizontal Checking Create Program]=========================>>>>

    # region <<<<=============================[Old Horizontal Machines Sub Functions]=============================>>>>

    # =======================================================================================================|
    # [#] Create Function to set LedgeCutStatus to be 'Need' to use Ledge cut tool.                          |
    # [#] Set the Color of Option of 'need_to_use_ledge_tool..' to be Green, and Set the Status to be 'True'.|
    # [#] Set the Color of Option of 'does_not_need_to_use_ledge_tool..' to remain same as Background color, |
    #     and Set the Status to be 'False'.                                                                  |
    # =======================================================================================================|
    def need_to_use_ledge_tool_for_old_horizontal_machine(self, obj):
        print("(need_to_use_ledge_tool_for_old_horizontal_machine) Function >> Called")
        self.need_to_use_ledge_tool_option_for_old_horizontal_machine.bg_color = (
            20 / 255, 82 / 255, 20 / 255, 1)
        self.does_not_need_to_use_ledge_tool_option_for_old_horizontal_machine.bg_color = (
            32 / 255.0, 32 / 255.0, 32 / 255.0, 1)
        self.need_ledge_tool_status = True
        self.does_not_need_ledge_tool_status = False

    # =============================================================================================|
    # [#] Create Function to set LedgeCutStatus to be 'NO Need' to use Ledge cut tool.             |
    # [#] Set the Color of Option of 'does_not_need_to_use_ledge_tool....' to be Green, and Set the|
    #     Status to be 'True'.                                                                     |
    # [#] Set the Color of Option of 'need_to_use_ledge_tool..' to remain same as Background color,|
    #     and Set the Status to be 'False'.                                                        |
    # =============================================================================================|
    def does_not_need_to_use_ledge_tool_for_old_horizontal_machine(self, obj):
        print("(does_not_need_to_use_ledge_tool_for_old_horizontal_machine) Function >> calledFunction >> Called")
        self.does_not_need_to_use_ledge_tool_option_for_old_horizontal_machine.bg_color = (
            20 / 255, 82 / 255, 20 / 255, 1)
        self.need_to_use_ledge_tool_option_for_old_horizontal_machine.bg_color = (
            32 / 255.0, 32 / 255.0, 32 / 255.0, 1)
        self.does_not_need_ledge_tool_status = True
        self.need_ledge_tool_status = False

    # ======================================================================================|
    # [#] Create Function to set and decide Ledge Cut Status.                               |
    # [#] Use (else-statement) to make sure the User choose one of the options, otherwise it|
    #     will keep ask and wait for the User input.                                        |
    # ======================================================================================|
    def decide_ledge_tool_status_for_old_horizontal_machine(self, obj):
        print("(decide_ledge_tool_status) Function >> Called")
        global ledge_cut_availability_status
        if (self.need_ledge_tool_status == True):
            ledge_cut_availability_status = 1
        elif (self.does_not_need_ledge_tool_status == True):
            ledge_cut_availability_status = 0
        else:
            print("WAITING FOR USER TO CHOOSE LADGE CUT STATUS")
        print("LEDGE_CUT_AVAILABILITY_STATUS in decide function ", ledge_cut_availability_status)
        # =======================================|
        # [#] To Close the screen message window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

    # =============================================================================================|
    # [#] Create Function to set the Offset Amount value.                                          |
    # [#] Use (if-statement) to make sure the User entered value in the TextInput box, otherwise it|
    #     will keep ask and wait for the User input.                                               |
    # =============================================================================================|
    def enter_offset_value(self, obj):
        print("(enter_offset_value) Function >> Called")
        global offset_amount
        if (self.offset_amount_confirmation_text_field.text != "" and
                self.offset_amount_confirmation_text_field.text is not None):
            offset_amount = float(self.offset_amount_confirmation_text_field.text)
            if (offset_amount == 0.0):
                offset_amount = math.floor(offset_amount)
            else:
                offset_amount = round(float(offset_amount), 4)
        # Most Likely We Don't Needed
        # verification_messages_of_creating_old_horizontal_machine_program = []
        print("offset value after ", offset_amount)
        # =======================================|
        # [#] To Close the screen message window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

    # =====================================================================================================|
    # [#] Create Function to set OffsetDirection to be 'To0'.                                              |
    # [#] Set the Color of Option of 'OffsetDirectionTo0' to be Green, and Set the Status to be 'True'.    |
    # [#] Set the Color of Options of 'OffsetDirectionTo180' and 'OffsetEachWayDirection' to remain same as|
    #     Background color, and Set their Status to be 'False'.                                            |
    # =====================================================================================================|
    def set_offset_direction_To0(self, obj):
        print("(set_offset_direction_To0) Function >> Called")
        self.offset_To0_option.bg_color = (20 / 255, 82 / 255, 20 / 255, 1)
        self.offset_To180_option.bg_color = (32 / 255.0, 32 / 255.0, 32 / 255.0, 1)
        self.offset_each_way_option.bg_color = (32 / 255.0, 32 / 255.0, 32 / 255.0, 1)
        self.offset_To0_option_status = True
        self.offset_To180_option_status = False
        self.offset_each_way_option_status = False
        print(self.offset_To0_option.text)
        print()

    # ===================================================================================================|
    # [#] Create Function to set OffsetDirection to be 'To180'.                                          |
    # [#] Set the Color of Option of 'OffsetDirectionTo180' to be Green, and Set the Status to be 'True'.|
    # [#] Set the Color of Options of 'OffsetDirectionTo0' and 'OffsetEachWayDirection' to remain same as|
    #     Background color, and Set their Status to be 'False'.                                          |
    # ===================================================================================================|
    def set_offset_direction_To180(self, obj):
        print("(set_offset_direction_To180) Function >> Called")
        self.offset_To180_option.bg_color = (20 / 255, 82 / 255, 20 / 255, 1)
        self.offset_To0_option.bg_color = (32 / 255.0, 32 / 255.0, 32 / 255.0, 1)
        self.offset_each_way_option.bg_color = (32 / 255.0, 32 / 255.0, 32 / 255.0, 1)
        self.offset_To180_option_status = True
        self.offset_To0_option_status = False
        self.offset_each_way_option_status = False
        print(self.offset_To180_option.text)
        print()

    # =====================================================================================================|
    # [#] Create Function to set OffsetDirection to be 'EachWay'.                                          |
    # [#] Set the Color of Option of 'OffsetDirectionEachWay' to be Green, and Set the Status to be 'True'.|
    # [#] Set the Color of Options of 'OffsetDirectionTo0' and 'OffsetDirectionTo180' to remain same as    |
    #     Background color, and Set their Status to be 'False'.                                            |
    # =====================================================================================================|
    def set_offset_direction_each_way(self, obj):
        print("(set_offset_direction_each_way) Function >> Called")
        self.offset_each_way_option.bg_color = (20 / 255, 82 / 255, 20 / 255, 1)
        self.offset_To0_option.bg_color = (32 / 255.0, 32 / 255.0, 32 / 255.0, 1)
        self.offset_To180_option.bg_color = (32 / 255.0, 32 / 255.0, 32 / 255.0, 1)
        self.offset_each_way_option_status = True
        self.offset_To180_option_status = False
        self.offset_To0_option_status = False
        print(self.offset_each_way_option.text)

    # ======================================================================================|
    # [#] Create Function to set and choose the OffsetDirection.                            |
    # [#] Use (else-statement) to make sure the User choose one of the options, otherwise it|
    #     will keep ask and wait for the User input.                                        |
    # ======================================================================================|
    def choose_offset_direction(self, obj):
        print("(choose_offset_direction) Function >> Called")
        global offset_direction
        if (self.offset_To0_option_status == True):
            offset_direction = self.offset_To0_option.text
        elif (self.offset_To180_option_status == True):
            offset_direction = self.offset_To180_option.text
        elif (self.offset_each_way_option_status == True):
            offset_direction = self.offset_each_way_option.text
        else:
            print("WAITING USER TO CHOOSE OFFSET_DIRECTION.")
        print("offset Direction after: ", offset_direction)
        # =======================================|
        # [#] To Close the screen message window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

    # ================================================================================================|
    # [#] Create Function to set PilotBoreDiameter to be the 'BigDiameter<2.25>'.                     |
    # [#] Set the Color of Option of 'BigDiameter<2.25>' to be Green, and Set the Status to be 'True'.|
    # [#] Set the Color of Option of 'SmallDiameter<1.70>' to remain same as Background color,        |
    #     and Set the Status to be 'False'.                                                           |
    # ================================================================================================|
    def set_big_pilot_bore_diameter(self, obj):
        print("(set_big_pilot_bore_diameter) Function >> Called")
        self.big_pilot_bore_diameter.bg_color = (20 / 255, 82 / 255, 20 / 255, 1)
        self.small_pilot_bore_diameter.bg_color = (32 / 255.0, 32 / 255.0, 32 / 255.0, 1)
        self.big_pilot_bore_diameter_Option_Status = True
        print("Big Pilot Status: ", self.big_pilot_bore_diameter_Option_Status)
        self.small_pilot_bore_diameter_Option_Status = False
        print("Small Pilot Status: ", self.small_pilot_bore_diameter_Option_Status)

    # ==================================================================================================|
    # [#] Create Function to set PilotBoreDiameter to be the 'SmallDiameter<1.70>'.                     |
    # [#] Set the Color of Option of 'SmallDiameter<1.70>' to be Green, and Set the Status to be 'True'.|
    # [#] Set the Color of Option of 'BigDiameter<2.25>' to remain same as Background color,            |
    #     and Set the Status to be 'False'.                                                             |
    # ==================================================================================================|
    def set_small_pilot_bore_diameter(self, obj):
        print("(set_small_pilot_bore_diameter) Function >> Called")
        self.small_pilot_bore_diameter.bg_color = (20 / 255, 82 / 255, 20 / 255, 1)
        self.big_pilot_bore_diameter.bg_color = (32 / 255.0, 32 / 255.0, 32 / 255.0, 1)
        self.small_pilot_bore_diameter_Option_Status = True
        print("Small Pilot Status: ", self.small_pilot_bore_diameter_Option_Status)
        self.big_pilot_bore_diameter_Option_Status = False
        print("Big Pilot Status: ", self.big_pilot_bore_diameter_Option_Status)

    # ======================================================================================|
    # [#] Create Function to set and choose the PilotBoreDiameter.                          |
    # [#] Use (else-statement) to make sure the User choose one of the options, otherwise it|
    #     will keep ask and wait for the User input.                                        |
    # ======================================================================================|
    def choose_pilot_bore_diameter(self, obj):
        print("(choose_pilot_bore_diameter) Function >> Called")
        global pilot_diameter
        if (self.big_pilot_bore_diameter_Option_Status == True):
            pilot_diameter = 2.25
        elif (self.small_pilot_bore_diameter_Option_Status == True):
            pilot_diameter = 1.70
        else:
            print("WAITING USER TO CHOOSE Pilot Diameter.")
        print("Pilot Diameter after User Choice: ", pilot_diameter)
        # =======================================|
        # [#] To Close the screen message window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

    # =============================================================================================|
    # [#] Create Function to set the HorizontalSlots_OD_Spacing value.                             |
    # [#] Use (if-statement) to make sure the User entered value in the TextInput box, otherwise it|
    #     will keep ask and wait for the User input.                                               |
    # =============================================================================================|
    def enter_horizontal_slots_OD_spacing_value(self, obj):
        print("(enter_horizontal_slots_OD_spacing_value) Function >> Called")
        global horizontal_slots_OD_spacing
        if (self.horiz_slots_confirmation_text_field.text != "" and
                self.horiz_slots_confirmation_text_field.text is not None):
            horizontal_slots_OD_spacing = float(self.horiz_slots_confirmation_text_field.text)
            if (horizontal_slots_OD_spacing == 0.0):
                horizontal_slots_OD_spacing = math.floor(horizontal_slots_OD_spacing)
            else:
                horizontal_slots_OD_spacing = round(float(horizontal_slots_OD_spacing), 4)
        # Most Likely We Don't Needed
        # verification_messages_of_creating_old_horizontal_machine_program = []
        print("Horizontal Slots OD Spacing AFTER User Input: ", horizontal_slots_OD_spacing)
        # =======================================|
        # [#] To Close the screen message window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

    # =============================================================================================|
    # [#] Create Function to set the PilotBoreDepth value.                                         |
    # [#] Use (if-statement) to make sure the User entered value in the TextInput box, otherwise it|
    #     will keep ask and wait for the User input.                                               |
    # =============================================================================================|
    def enter_pilot_bore_depth_value(self, obj):
        print("(enter_pilot_bore_depth_value) Function >> Called")
        global pilot_bore_depth
        global pilot_bore_depth_verifying_status
        if (self.pilot_bore_depth_confirmation_text_field.text != "" and
                self.pilot_bore_depth_confirmation_text_field.text is not None):
            pilot_bore_depth = float(self.pilot_bore_depth_confirmation_text_field.text)
            pilot_bore_depth_verifying_status = True
            if (pilot_bore_depth == 0.0):
                pilot_bore_depth = math.floor(pilot_bore_depth)
            else:
                pilot_bore_depth = round(float(pilot_bore_depth), 4)
        # Most Likely We Don't Needed
        # verification_messages_of_creating_old_horizontal_machine_program = []
        print("Pilot Bore Depth AFTER User Input: ", pilot_bore_depth)
        # =======================================|
        # [#] To Close the screen message window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

    # =============================================================================================|
    # [#] Create Function to set the PilotToPin value.                                             |
    # [#] Use (if-statement) to make sure the User entered value in the TextInput box, otherwise it|
    #     will keep ask and wait for the User input.                                               |
    # =============================================================================================|
    def enter_pilot_to_pin_value(self, obj):
        print("(enter_pilot_to_pin_value) Function >> Called")
        global pilot_to_pin
        global pilot_to_pin_verifying_status
        if (self.pilot_to_pin_confirmation_text_field.text != "" and
                self.pilot_to_pin_confirmation_text_field.text is not None):
            pilot_to_pin = float(self.pilot_to_pin_confirmation_text_field.text)
            pilot_to_pin_verifying_status = True
            if (pilot_to_pin == 0.0):
                pilot_to_pin = math.floor(pilot_to_pin)
            else:
                pilot_to_pin = round(float(pilot_to_pin), 4)
        # Most Likely We Don't Needed
        # verification_messages_of_creating_old_horizontal_machine_program = []
        print("Pilot To Pin AFTER User Input: ", pilot_to_pin)
        # =======================================|
        # [#] To Close the screen message window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

    # =============================================================================================|
    # [#] Create Function to set the PinHoleDiameter value.                                        |
    # [#] Use (if-statement) to make sure the User entered value in the TextInput box, otherwise it|
    #     will keep ask and wait for the User input.                                               |
    # =============================================================================================|
    def enter_pin_hole_diameter_value(self, obj):
        print("(enter_pin_hole_diameter_value) Function >> Called")
        global pin_hole_diameter
        global pin_hole_diameter_verifying_status
        if (self.pin_hole_diameter_confirmation_text_field.text != "" and
                self.pin_hole_diameter_confirmation_text_field.text is not None):
            pin_hole_diameter = float(self.pin_hole_diameter_confirmation_text_field.text)
            pin_hole_diameter_verifying_status = True
            if (pin_hole_diameter == 0.0):
                pin_hole_diameter = math.floor(pin_hole_diameter)
            else:
                pin_hole_diameter = round(float(pin_hole_diameter), 4)
        # Most Likely We Don't Needed
        # verification_messages_of_creating_old_horizontal_machine_program = []
        print("Pin Hole Diameter AFTER User Input: ", pin_hole_diameter)
        # =======================================|
        # [#] To Close the screen message window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

    # =============================================================================================|
    # [#] Create Function to set the LockRing_ID_Spacing value.                                    |
    # [#] Use (if-statement) to make sure the User entered value in the TextInput box, otherwise it|
    #     will keep ask and wait for the User input.                                               |
    # =============================================================================================|
    def enter_lock_ring_ID_spacing_value(self, obj):
        print("(enter_lock_ring_ID_spacing_value) Function >> Called")
        global lock_ring_ID_spacing
        if (self.lock_ring_ID_spacing_confirmation_text_field.text != "" and
                self.lock_ring_ID_spacing_confirmation_text_field.text is not None):
            lock_ring_ID_spacing = float(self.lock_ring_ID_spacing_confirmation_text_field.text)
            if (lock_ring_ID_spacing == 0.0):
                lock_ring_ID_spacing = math.floor(lock_ring_ID_spacing)
            else:
                lock_ring_ID_spacing = round(float(lock_ring_ID_spacing), 4)
        # Most Likely We Don't Needed
        # verification_messages_of_creating_old_horizontal_machine_program = []
        print("Lock Ring ID Spacing value AFTER User Input: :", lock_ring_ID_spacing)
        # =======================================|
        # [#] To Close the screen message window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

    # =============================================================================================|
    # [#] Create Function to set the LockRingDiameter value.                                       |
    # [#] Use (if-statement) to make sure the User entered value in the TextInput box, otherwise it|
    #     will keep ask and wait for the User input.                                               |
    # =============================================================================================|
    def enter_lock_ring_diameter_value(self, obj):
        print("(enter_lock_ring_diameter_value) Function >> Called")
        global lock_ring_diameter
        if (self.lock_ring_diameter_confirmation_text_field.text != "" and
                self.lock_ring_diameter_confirmation_text_field.text is not None):
            lock_ring_diameter = float(self.lock_ring_diameter_confirmation_text_field.text)
            if (lock_ring_diameter == 0.0):
                lock_ring_diameter = math.floor(lock_ring_diameter)
            else:
                lock_ring_diameter = round(float(lock_ring_diameter), 4)
        # Most Likely We Don't Needed
        # verification_messages_of_creating_old_horizontal_machine_program = []
        print("Lock Ring Diameter value AFTER User Input: :", lock_ring_diameter)
        # =======================================|
        # [#] To Close the screen message window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

    # =============================================================================================|
    # [#] Create Function to set the C/Fren_ID_Spacing value.                                      |
    # [#] Use (if-statement) to make sure the User entered value in the TextInput box, otherwise it|
    #     will keep ask and wait for the User input.                                               |
    # =============================================================================================|
    def enter_cfren_ID_spacing_value(self, obj):
        print("(enter_cfren_ID_spacing_value) Function >> Called")
        global cfren_ID_spacing
        if (self.cfren_ID_spacing_confirmation_text_field.text != "" and
                self.cfren_ID_spacing_confirmation_text_field.text is not None):
            cfren_ID_spacing = float(self.cfren_ID_spacing_confirmation_text_field.text)
            if (cfren_ID_spacing == 0.0):
                cfren_ID_spacing = math.floor(cfren_ID_spacing)
            else:
                cfren_ID_spacing = round(float(cfren_ID_spacing), 4)
        # Most Likely We Don't Needed
        # verification_messages_of_creating_old_horizontal_machine_program = []
        print("C/Fren ID Spacing value AFTER User Input: :", cfren_ID_spacing)
        # =======================================|
        # [#] To Close the screen message window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

    # =============================================================================================|
    # [#] Create Function to set the C/FrenCutterWidth value.                                      |
    # [#] Use (if-statement) to make sure the User entered value in the TextInput box, otherwise it|
    #     will keep ask and wait for the User input.                                               |
    # =============================================================================================|
    def enter_cfren_cutter_width_value(self, obj):
        print("(enter_cfren_cutter_width_value) Function >> Called")
        global cfren_cutter_width
        if (self.cfren_cutter_width_confirmation_text_field.text != "" and
                self.cfren_cutter_width_confirmation_text_field.text is not None):
            cfren_cutter_width = float(self.cfren_cutter_width_confirmation_text_field.text)
            if (cfren_cutter_width == 0.0):
                cfren_cutter_width = math.floor(cfren_cutter_width)
            else:
                cfren_cutter_width = round(float(cfren_cutter_width), 4)
        # Most Likely We Don't Needed
        # verification_messages_of_creating_old_horizontal_machine_program = []
        print("C/Fren Cutter Width value AFTER User Input: :", cfren_cutter_width)
        # =======================================|
        # [#] To Close the screen message window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

    # =============================================================================================|
    # [#] Create Function to set the C/FrenDiameter value.                                         |
    # [#] Use (if-statement) to make sure the User entered value in the TextInput box, otherwise it|
    #     will keep ask and wait for the User input.                                               |
    # =============================================================================================|
    def enter_cfren_diameter_value(self, obj):
        print("(enter_cfren_diameter_value) Function >> Called")
        global cfren_diameter
        if (self.cfren_diameter_confirmation_text_field.text != "" and
                self.cfren_diameter_confirmation_text_field.text is not None):
            cfren_diameter = float(self.cfren_diameter_confirmation_text_field.text)
            if (cfren_diameter == 0.0):
                cfren_diameter = math.floor(cfren_diameter)
            else:
                cfren_diameter = round(float(cfren_diameter), 4)
        # Most Likely We Don't Needed
        # verification_messages_of_creating_old_horizontal_machine_program = []
        print("C/Fren Diameter value AFTER User Input: :", cfren_diameter)
        # =======================================|
        # [#] To Close the screen message window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

    # =============================================================================================|
    # [#] Create Function to set the ForgeRefLength value.                                         |
    # [#] Use (if-statement) to make sure the User entered value in the TextInput box, otherwise it|
    #     will keep ask and wait for the User input.                                               |
    # =============================================================================================|
    def enter_forge_ref_length_value(self, obj):
        print("(enter_forge_ref_length_value) Function >> Called")
        global forge_ref_length
        if (self.forge_ref_length_confirmation_text_field.text != "" and
                self.forge_ref_length_confirmation_text_field.text is not None):
            forge_ref_length = float(self.forge_ref_length_confirmation_text_field.text)
            if (forge_ref_length == 0.0):
                forge_ref_length = math.floor(forge_ref_length)
            else:
                forge_ref_length = round(float(forge_ref_length), 4)
        # Most Likely We Don't Needed
        # verification_messages_of_creating_old_horizontal_machine_program = []
        print("Forge Ref Length value AFTER User Input: ", forge_ref_length)
        # =======================================|
        # [#] To Close the screen message window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

    # =============================================================================================|
    # [#] Create Function to set the ForgingDiameter value.                                        |
    # [#] Use (if-statement) to make sure the User entered value in the TextInput box, otherwise it|
    #     will keep ask and wait for the User input.                                               |
    # =============================================================================================|
    def enter_forging_diameter_value(self, obj):
        print("(enter_forging_diameter_value) Function >> Called")
        global forging_diameter
        if (self.forging_diameter_confirmation_text_field.text != "" and
                self.forging_diameter_confirmation_text_field.text is not None):
            forging_diameter = float(self.forging_diameter_confirmation_text_field.text)
            if (forging_diameter == 0.0):
                forging_diameter = math.floor(forging_diameter)
            else:
                forging_diameter = round(float(forging_diameter), 4)
        # Most Likely We Don't Needed
        # verification_messages_of_creating_old_horizontal_machine_program = []
        print("Forging Diameter value AFTER User Input: ", forging_diameter)
        # =======================================|
        # [#] To Close the screen message window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

    # =============================================================================================|
    # [#] Create Function to set the ForgingOutsideBossSpacing value.                              |
    # [#] Use (if-statement) to make sure the User entered value in the TextInput box, otherwise it|
    #     will keep ask and wait for the User input.                                               |
    # =============================================================================================|
    def enter_forging_outside_boss_spacing_value(self, obj):
        print("(enter_forging_outside_boss_spacing_value) Function >> Called")
        global forging_outside_boss_spacing
        if (self.forging_outside_boss_spacing_confirmation_text_field.text != "" and
                self.forging_outside_boss_spacing_confirmation_text_field.text is not None):
            forging_outside_boss_spacing = float(self.forging_outside_boss_spacing_confirmation_text_field.text)
            if (forging_outside_boss_spacing == 0.0):
                forging_outside_boss_spacing = math.floor(forging_outside_boss_spacing)
            else:
                forging_outside_boss_spacing = round(float(forging_outside_boss_spacing), 4)
        # Most Likely We Don't Needed
        # verification_messages_of_creating_old_horizontal_machine_program = []
        print("Forging Outside Boss Spacing value AFTER User Input: :", forging_outside_boss_spacing)
        # =======================================|
        # [#] To Close the screen message window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

    # =============================================================================================|
    # [#] Create Function to set the ForgingInsideBossSpacing value.                               |
    # [#] Use (if-statement) to make sure the User entered value in the TextInput box, otherwise it|
    #     will keep ask and wait for the User input.                                               |
    # =============================================================================================|
    def enter_forging_inside_boss_spacing_value(self, obj):
        print("(enter_forging_inside_boss_spacing_value) Function >> Called")
        global forging_inside_boss_spacing
        if (self.forging_inside_boss_spacing_confirmation_text_field.text != "" and
                self.forging_inside_boss_spacing_confirmation_text_field.text is not None):
            forging_inside_boss_spacing = float(self.forging_inside_boss_spacing_confirmation_text_field.text)
            if (forging_inside_boss_spacing == 0.0):
                forging_inside_boss_spacing = math.floor(forging_inside_boss_spacing)
            else:
                forging_inside_boss_spacing = round(float(forging_inside_boss_spacing), 4)
        # Most Likely We Don't Needed
        # verification_messages_of_creating_old_horizontal_machine_program = []
        print("Forging Inside Boss Spacing value AFTER User Input: :", forging_inside_boss_spacing)
        # =======================================|
        # [#] To Close the screen message window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

    # ==========================================================================================|
    # [#] Create Function to Replace existing Horizontal PinBore program in the Original folder.|
    # ==========================================================================================|
    def replace_existing_old_horizontal_machine_program_in_original_folder(self, obj):
        print("(replace_existing_old_horizontal_machine_program_in_original_folder) Function >> Called")
        # =================================================================|
        # [#] To make sure ALL the Screen Message Windows have been Closed.|
        # =================================================================|
        self.old_horizontal_screen_message_window.dismiss()

        # =====================================================================================================|
        # [#] Same Description of the above Function (Note<1,1>), Except it is on the Replace Program Function.|
        # =====================================================================================================|
        global called_need_confirmation_to_create_old_horizontal_machine_program
        called_need_confirmation_to_create_old_horizontal_machine_program = False
        print("called_need_confirmation_to_create_old_horizontal_machine_program on (Replace "
              "Program Function) in Original Folder: ",
              called_need_confirmation_to_create_old_horizontal_machine_program)

        # =====================================================================================================|
        # [#] Same Description of the above Function (Note<1,9>), Except it is on the Replace Program Function.|
        # =====================================================================================================|
        if (offset_direction == "OFFSET EACH WAY" and offset_amount != 0):

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,10>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            global new_horizontal_program_To0_direction_for_old_machine_in_original_folder
            new_horizontal_program_To0_direction_for_old_machine_in_original_folder = (
                    original_folder_path_of_old_horizontal_machine + "\\" + "P" +
                    new_program_number_for_old_horizontal_machine + "TO0.MIN")

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,11>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            horizontal_program_lines_To0_direction_for_old_machine_in_original_folder[0] = (
                    '(PART ' + new_program_number_for_old_horizontal_machine + " TO 0" + ' -- ' + today_date + ' <SYS>' +
                    ')' + "\n".join(notes_for_old_horizontal_machine_program))

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,12>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            try:
                create_new_horizontal_program_To0_direction_for_old_machine_in_original_folder = open(
                    new_horizontal_program_To0_direction_for_old_machine_in_original_folder, "w")
                create_new_horizontal_program_To0_direction_for_old_machine_in_original_folder.write(
                    '\n'.join(horizontal_program_lines_To0_direction_for_old_machine_in_original_folder))
                create_new_horizontal_program_To0_direction_for_old_machine_in_original_folder.close()

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,13>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            except PermissionError or FileNotFoundError as error:
                fail_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find" + '[b][color=ffff00] ORIGINAL [/color][/b]' +
                    "Folder location to Save the Program." + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]'
                    + str(error) + '[/color]' + "\n" + "Double Check Network, and Folder Location." + "\n")
                email_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find Original Folder location to Save the Program." + "\n" +
                    "Double Check Network, and File Location." + "\n")
                failed_to_create_old_horizontal_machine_program(self)
                return

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,14>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            global new_horizontal_program_To180_direction_for_old_machine_in_original_folder
            new_horizontal_program_To180_direction_for_old_machine_in_original_folder = (
                    original_folder_path_of_old_horizontal_machine + "\\" + "P" +
                    new_program_number_for_old_horizontal_machine + "TO180.MIN")

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,15>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            horizontal_program_lines_To180_direction_for_old_machine_in_original_folder[0] = (
                    '(PART ' + new_program_number_for_old_horizontal_machine + " TO 180" + ' -- ' + today_date +
                    ' <SYS>' + ')' + "\n".join(notes_for_old_horizontal_machine_program))

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,16>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            horizontal_program_lines_To180_direction_for_old_machine_in_original_folder[VC155_variable_index] = (
                    'VC155=' + format((-1) * offset_amount) + '  (Offset)')

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,17>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            if (notch_angle_first_location != 0 and notch_angle_first_location is not None):
                horizontal_program_lines_To180_direction_for_old_machine_in_original_folder[VC173_variable_index] = (
                        'VC173=' + format(float(Y_distance_from_origin_to_circlip_notch) - (2 * offset_amount), '.4f') +
                        '  (yCirclipNotch)')

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,18>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            try:
                create_new_horizontal_program_To180_direction_for_old_machine_in_original_folder = open(
                    new_horizontal_program_To180_direction_for_old_machine_in_original_folder, "w")
                create_new_horizontal_program_To180_direction_for_old_machine_in_original_folder.write(
                    '\n'.join(horizontal_program_lines_To180_direction_for_old_machine_in_original_folder))
                create_new_horizontal_program_To180_direction_for_old_machine_in_original_folder.close()

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,19>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            except PermissionError or FileNotFoundError as error:
                fail_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find" + '[b][color=ffff00] ORIGINAL [/color][/b]' +
                    "Folder location to Save the Program." + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]'
                    + str(error) + '[/color]' + "\n" + "Double Check Network, and Folder Location." + "\n")
                email_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find Original Folder location to Save the Program." + "\n" +
                    "Double Check Network, and File Location." + "\n")
                failed_to_create_old_horizontal_machine_program(self)
                return

        # ======================================================================================================|
        # [#] Same Description of the above Function (Note<1,20>), Except it is on the Replace Program Function.|
        # ======================================================================================================|
        else:
            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,21>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            new_horizontal_program_for_old_machine_in_original_folder = (
                    original_folder_path_of_old_horizontal_machine + "\\" + "P" +
                    new_program_number_for_old_horizontal_machine + ".MIN")

            # ==================================================================================================|
            # [#] Access the First Element of List of (Main Program lines) by use ([0] index) to Set the First  |
            #     line of the Program with all changes that needed (Ex: (PART WD-13000 -- 05/14/2022 <SYS>)).   |
            # [#] <SYS>: Used to Indicate this Program is Created By 'WisecoProgramsMaker' APP.                 |
            # [#] (notes_for_old_horizontal_machine_program): To Add any note that needed to the Top of Program.|
            # ==================================================================================================|
            pin_bore_program_lines_of_old_horizontal_machine[0] = (
                    '(PART ' + new_program_number_for_old_horizontal_machine + ' -- ' + today_date + ' <SYS>' + ')' +
                    "\n".join(notes_for_old_horizontal_machine_program))

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,22>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            try:
                create_new_horizontal_program_for_old_machine_in_original_folder = open(
                    new_horizontal_program_for_old_machine_in_original_folder, "w")
                create_new_horizontal_program_for_old_machine_in_original_folder.write(
                    '\n'.join(pin_bore_program_lines_of_old_horizontal_machine))
                create_new_horizontal_program_for_old_machine_in_original_folder.close()

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,23>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            except PermissionError or FileNotFoundError as error:
                fail_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find" + '[b][color=ffff00] ORIGINAL [/color][/b]' +
                    "Folder location to Save the Program." + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]'
                    + str(error) + '[/color]' + "\n" + "Double Check Network, and Folder Location." + "\n")
                email_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find Original Folder location to Save the Program." + "\n" +
                    "Double Check Network, and File Location." + "\n")
                failed_to_create_old_horizontal_machine_program(self)
                return

        # ======================================================================================================|
        # [#] Same Description of the above Function (Note<1,24>), Except it is on the Replace Program Function.|
        # ======================================================================================================|
        if (old_horizontal_program_confirmation_email_message_list != []):
            email_messages_of_creating_old_horizontal_machine_program.append(
                "\n".join(old_horizontal_program_confirmation_email_message_list) + "\n")

        success_messages_of_creating_old_horizontal_machine_program = [
            "\n" + "Program has been **REPLACED** successfully in **ORIGINAL** Folder." + "\n"]
        email_messages_of_creating_old_horizontal_machine_program.append("\n".join(
            success_messages_of_creating_old_horizontal_machine_program))

        # ======================================================================================================|
        # [#] Same Description of the above Function (Note<1,25>), Except it is on the Replace Program Function.|
        # ======================================================================================================|
        close_button = MDRaisedButton(text='Close', on_release=self.close_old_horizontal_window_of_original_folder,
                                      font_size=16)

        # ======================================================================================================|
        # [#] Same Description of the above Function (Note<1,26>), Except it is on the Replace Program Function.|
        # ======================================================================================================|
        self.old_horizontal_screen_message_window = MDDialog(title='[color=248f24]Success Message[/color]', text=(
                '[b][i][u][color=0099ff]' + self.ids["JobNumberForOldHorizontalMachine"].text +
                '[/color][/u][/i][/b]' + '[color=ffffff] Program Has been [/color]' +
                '[b][color=ffffff]REPLACED[/color][/b]' + '[color=ffffff] Successfully in [/color]' +
                '[color=ffff00]ORIGINAL[/color]' + '[color=ffffff] Folder.[/color]'), size_hint=(0.7, 1.0),
                                                             buttons=[close_button], auto_dismiss=False)
        self.old_horizontal_screen_message_window.open()
        print()

    # =========================================================================================|
    # [#] Create Function to Replace existing Horizontal PinBore program in the Running folder.|
    # =========================================================================================|
    def replace_existing_old_horizontal_machine_program_in_running_folder(self, obj):
        print("9replace_existing_old_horizontal_machine_program_in_running_folder) Function >> Called")

        # =================================================================|
        # [#] To make sure ALL the Screen Message Windows have been Closed.|
        # =================================================================|
        self.old_horizontal_screen_message_window.dismiss()

        # =====================================================================================================|
        # [#] Same Description of the above Function (Note<1,1>), Except it is on the Replace Program Function.|
        # =====================================================================================================|
        global called_need_confirmation_to_create_old_horizontal_machine_program
        called_need_confirmation_to_create_old_horizontal_machine_program = False
        print(
            "called_need_confirmation_to_create_old_horizontal_machine_program on Replace Program in Running Folder: ",
            called_need_confirmation_to_create_old_horizontal_machine_program)

        # =====================================================================================================|
        # [#] Same Description of the above Function (Note<1,9>), Except it is on the Replace Program Function.|
        # =====================================================================================================|
        if (offset_direction == "OFFSET EACH WAY" and offset_amount != 0):

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,10>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            global new_horizontal_program_To0_direction_for_old_machine_in_running_folder
            new_horizontal_program_To0_direction_for_old_machine_in_running_folder = (
                    running_folder_path_of_old_horizontal_machine + "\\" + "P" +
                    new_program_number_for_old_horizontal_machine + "TO0.MIN")

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,11>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            horizontal_program_lines_To0_direction_for_old_machine_in_running_folder[0] = (
                    '(PART ' + new_program_number_for_old_horizontal_machine + " TO 0" + ' -- '
                    + today_date + ' <SYS>' + ')')

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,12>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            try:
                create_new_horizontal_program_To0_direction_for_old_machine_in_running_folder = open(
                    new_horizontal_program_To0_direction_for_old_machine_in_running_folder, "w")
                create_new_horizontal_program_To0_direction_for_old_machine_in_running_folder.write(
                    '\n'.join(horizontal_program_lines_To0_direction_for_old_machine_in_running_folder))
                create_new_horizontal_program_To0_direction_for_old_machine_in_running_folder.close()

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,13>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            except PermissionError or FileNotFoundError as error:
                # print("except LOAD_HORIZONTAL_SHEETS_FOR_AUTOMATION")
                fail_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find" + '[b][color=33cc33] RUNNING [/color][/b]' + "Folder location to Save the Program."
                    + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]' + str(
                        error) + '[/color]' + "\n" + "Double Check Network, and Folder Location." + "\n")
                email_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find Running Folder location to Save the Program." + "\n" +
                    "Double Check Network, and File Location." + "\n")
                failed_to_create_old_horizontal_machine_program(self)
                return

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,14>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            global new_horizontal_program_To180_direction_for_old_machine_in_running_folder
            new_horizontal_program_To180_direction_for_old_machine_in_running_folder = (
                    running_folder_path_of_old_horizontal_machine + "\\" + "P" +
                    new_program_number_for_old_horizontal_machine + "TO180.MIN")

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,15>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            horizontal_program_lines_To180_direction_for_old_machine_in_running_folder[0] = (
                    '(PART ' + new_program_number_for_old_horizontal_machine + " TO 180" + ' -- '
                    + today_date + ' <SYS>' + ')')

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,16>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            horizontal_program_lines_To180_direction_for_old_machine_in_running_folder[VC155_variable_index] = (
                    'VC155=' + format((-1) * offset_amount) + '  (Offset)')

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,17>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            if (notch_angle_first_location != 0 and notch_angle_first_location is not None):
                horizontal_program_lines_To180_direction_for_old_machine_in_running_folder[VC173_variable_index] = (
                        'VC173=' + format(float(Y_distance_from_origin_to_circlip_notch) - (2 * offset_amount), '.4f') +
                        '  (yCirclipNotch)')

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,18>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            try:
                create_new_horizontal_program_To180_direction_for_old_machine_in_running_folder = open(
                    new_horizontal_program_To180_direction_for_old_machine_in_running_folder, "w")
                create_new_horizontal_program_To180_direction_for_old_machine_in_running_folder.write(
                    '\n'.join(horizontal_program_lines_To180_direction_for_old_machine_in_running_folder))
                create_new_horizontal_program_To180_direction_for_old_machine_in_running_folder.close()

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,19>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            except PermissionError or FileNotFoundError as error:
                # print("except LOAD_HORIZONTAL_SHEETS_FOR_AUTOMATION")
                fail_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find" + '[b][color=33cc33] RUNNING [/color][/b]' +
                    "Folder location to Save the Program." + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]'
                    + str(error) + '[/color]' + "\n" + "Double Check Network, and Folder Location." + "\n")
                email_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find Running Folder location to Save the Program." + "\n" +
                    "Double Check Network, and File Location." + "\n")
                failed_to_create_old_horizontal_machine_program(self)
                return

        # ======================================================================================================|
        # [#] Same Description of the above Function (Note<1,20>), Except it is on the Replace Program Function.|
        # ======================================================================================================|
        else:
            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,21>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            new_horizontal_program_for_old_machine_in_running_folder = (
                    running_folder_path_of_old_horizontal_machine + "\\" + "P" +
                    new_program_number_for_old_horizontal_machine + ".MIN")

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,22>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            try:
                create_new_horizontal_program_for_old_machine_in_running_folder = open(
                    new_horizontal_program_for_old_machine_in_running_folder, "w")
                create_new_horizontal_program_for_old_machine_in_running_folder.write('\n'.join(
                    pin_bore_program_lines_of_old_horizontal_machine))
                create_new_horizontal_program_for_old_machine_in_running_folder.close()

            # ======================================================================================================|
            # [#] Same Description of the above Function (Note<1,23>), Except it is on the Replace Program Function.|
            # ======================================================================================================|
            except PermissionError or FileNotFoundError as error:
                # print("except LOAD_HORIZONTAL_SHEETS_FOR_AUTOMATION")
                fail_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find" + '[b][color=33cc33] RUNNING [/color][/b]' +
                    "Folder location to Save the Program." + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]'
                    + str(error) + '[/color]' + "\n" + "Double Check Network, and Folder Location." + "\n")
                email_messages_of_creating_old_horizontal_machine_program.append(
                    "Failed to Find Running Folder location to Save the Program." + "\n" +
                    "Double Check Network, and File Location." + "\n")
                failed_to_create_old_horizontal_machine_program(self)
                return

        # ======================================================================================================|
        # [#] Same Description of the above Function (Note<1,24>), Except it is on the Replace Program Function.|
        # ======================================================================================================|
        success_messages_of_creating_old_horizontal_machine_program = [
            "\n" + "Program has been **REPLACED** successfully in **RUNNING** Folder." + "\n"]
        email_messages_of_creating_old_horizontal_machine_program.append(
            "\n".join(success_messages_of_creating_old_horizontal_machine_program))

        # ======================================================================================================|
        # [#] Same Description of the above Function (Note<1,25>), Except it is on the Replace Program Function.|
        # ======================================================================================================|
        close_button = MDRaisedButton(text='Close', on_release=self.close_old_horizontal_window_of_running_folder,
                                      font_size=16)

        # ======================================================================================================|
        # [#] Same Description of the above Function (Note<1,26>), Except it is on the Replace Program Function.|
        # ======================================================================================================|
        self.old_horizontal_screen_message_window = MDDialog(title='[color=248f24]Success Message[/color]', text=(
                '[b][i][u][color=0099ff]' + self.ids["JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                '[color=ffffff] Program Has been [/color]' + '[b][color=ffffff]REPLACED[/color][/b]' +
                '[color=ffffff] Successfully in [/color]' + '[color=33cc33]RUNNING[/color]' +
                '[color=ffffff] Folder.[/color]' + '\n' +
                '[color=ffffff]After closing this window, the program will open on CIMCO Editor.[/color]' +
                '\n' + '[color=ffffff]Please Double Check and Report any Problem on Trello.[/color]'),
                                                             size_hint=(0.7, 1.0), buttons=[close_button],
                                                             auto_dismiss=False)
        self.old_horizontal_screen_message_window.open()
        print()

    # ==================================================================================================|
    # [#] Create Function to Avoid 'SKIP' Creating a program in the Original folder when the User clicks|
    #     on 'No_Button' for (Replace Program Question).                                                |
    # ==================================================================================================|
    def skip_create_old_horizontal_machine_program_in_original_folder(self, obj):
        print("(skip_create_old_horizontal_machine_program_in_original_folder) Function >> Called")

        # ==========================================================================================|
        # [#] Set Status to be 'True' when the User clicks on 'No_Button' to avoid make the program.|
        # ==========================================================================================|
        global dont_create_old_horizontal_machine_program_in_original_folder
        dont_create_old_horizontal_machine_program_in_original_folder = True

    # ============================================================================|
    # [#] Create Function to Close the Message Window that related on the Original|
    #     folder and execute some actions.                                        |
    # ============================================================================|
    def close_old_horizontal_window_of_original_folder(self, obj):
        print("(close_old_horizontal_window_of_original_folder) Function >> Called")

        # NEEDS TO DELETE BELOW ONCE MAKE SURE ALL IS GOOD
        # # =================================================================================================|
        # # [#] As the App shows a messages of the exceptions (on the Top of Program) that need to modify    |
        # #     the program manually, also it will warn the User by display the message of the exceptions on |
        # #     the Screen Message Window of the App.                                                        |
        # # [#] Use (while) loop to display All the messages of the exceptions together on the Screen Message|
        # #     Window of the App.                                                                           |
        # # [#] Define Variable to set Condition of (while) loop to be 'False' by default, and it will change|
        # #     to be 'True' when recognizing and displaying any Exceptional Cases.                          |
        # # =================================================================================================|
        # old_horizontal_exceptional_cases = False
        # while old_horizontal_exceptional_cases is False:
        #
        #     # =================================================================================================|
        #     # [#] Make Variables 'global' again just to make code work to be able to use it in other functions.|
        #     # =================================================================================================|
        #     global dont_create_old_horizontal_machine_program_in_original_folder
        #     global warning_messages_of_creating_old_horizontal_machine_program
        #
        #     # global result_of_existing_programs
        #     # # LIST TO STORE THE FOUND JOBS WITH WHOLE PATH INCLUDING FILE NAME
        #     # result_of_existing_programs = []
        #     # for file in glob.glob(original_folder_path_of_old_horizontal_machine + '\*' +
        #     #                       new_program_number_for_old_horizontal_machine + '*'):
        #     #     # TO APPEND(ADD) EACH PROGRAM THAT FOUND TO THE RESULT LIST
        #     #     result_of_existing_programs.append(file)
        #     # print("result_of_existing_programs: ", result_of_existing_programs)

        # =======================================|
        # [#] To Close the Screen Message Window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

        # region  <<<<========================[Exceptional Programs Cases <Messages>]==========================>>>>

        # =================================================================================================|
        # [#] As the App shows a messages of the exceptions (on the Top of Program) that need to modify    |
        #     the program manually, also it will warn the User by display the message of the exceptions on |
        #     the Screen Message Window of the App.                                                        |
        # [#] Use (while) loop to display All the messages of the exceptions together on the Screen Message|
        #     Window of the App.                                                                           |
        # [#] Define Variable to set Condition of (while) loop to be 'False' by default, and it will change|
        #     to be 'True' when recognizing and displaying any Exceptional Cases.                          |
        # =================================================================================================|
        old_horizontal_exceptional_cases = False
        while old_horizontal_exceptional_cases is False:

            # =================================================================================================|
            # [#] Make Variables 'global' again just to make code work to be able to use it in other functions.|
            # =================================================================================================|
            global dont_create_old_horizontal_machine_program_in_original_folder
            global warning_messages_of_creating_old_horizontal_machine_program

            # ===============================================================================================|
            # [#] Define List to contain ALL the programs that saved on similar names, the main reason to set|
            #     this List to catch any duplicated programs for the same job.                               |
            # ===============================================================================================|
            global result_of_existing_programs
            result_of_existing_programs = []

            # ===============================================================================================|
            # [#] Use (for) loop with (glob.glob) method to find all the programs that saved on similar names|
            #     on the Original folder.                                                                    |
            # [#] Use (*) after Job Number to search the file even with part of the name to find all possible|
            #     results (Ex: name of AW-06048 it gives result of AW-06048, AW-06048TO0 and AW-06048TO180). |
            # ===============================================================================================|
            for file in glob.glob(original_folder_path_of_old_horizontal_machine + '\*' +
                                  new_program_number_for_old_horizontal_machine + '*'):
                result_of_existing_programs.append(file)
            print("result_of_existing_programs: ", result_of_existing_programs)

            # ================================================================================================|
            # [#] Check if the job saved on 'one' and 'Each Way' direction offset at the same time , therefore|
            #     it needs to warn the user to fix the confusing.                                             |
            # ================================================================================================|
            if ((((original_folder_path_of_old_horizontal_machine + "\\" + "P" +
                   new_program_number_for_old_horizontal_machine + "TO0.MIN") in result_of_existing_programs) or
                 ((original_folder_path_of_old_horizontal_machine + "\\" + "P" +
                   new_program_number_for_old_horizontal_machine + "TO180.MIN") in result_of_existing_programs)) and
                    ((original_folder_path_of_old_horizontal_machine + "\\" + "P" +
                      new_program_number_for_old_horizontal_machine + ".MIN") in result_of_existing_programs)):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                    '[color=ffffff] job saved in several ways :[/color]' + '\n' + '\n' + '[color=ffffff]' +
                    ('\n'.join(result_of_existing_programs)) + '[/color]' + '\n' + '\n' +
                    '[color=ffffff]Double check in [/color]' + '[color=ffff00]ORIGINAL[/color]' +
                    '[color=ffffff] and [/color]' + '[color=33cc33]RUNNING[/color]' +
                    '[color=ffffff] Folders , and[/color]' + '\n' +
                    '[color=ffffff]DELETE the Wrong Program.[/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append("Job saved in several ways" + '\n' + (
                    '\n'.join(result_of_existing_programs)) + '\n' + "Double check in ORIGINAL and RUNNING Folders" +
                                                     '\n' + "and DELETE the Wrong Program.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,1>), Except it's display on |
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if ((horizontal_slots_diameter_depth != 0 and horizontal_slots_diameter_depth is not None)
                    and pressure_fed_holes_availability_status == 1):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                     '[color=ffffff] job has Horizontal slots and Pressure fed oil holes features '
                     'which is need to add the horizontal slots manually. [/color]' + '\n' + '\n' +
                     '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                     'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                     'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job has Horizontal slots and Pressure fed oil holes features "
                    "which is need to add the horizontal slots manually." + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,2>), Except it's display on |
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if ((horizontal_slots_diameter_depth != 0 and horizontal_slots_diameter_depth is not None)
                    and (horiz_piston_pin_slots_notes is not None and horiz_piston_pin_slots_notes != "") and
                    (horiz_piston_pin_slots_notes.find('1 SLOT PER SIDE') != -1)):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                     '[color=ffffff] job has One Horizontal slot only per side '
                     'which is need to delete one of the slot in horizontal template manually. [/color]' + '\n' + '\n' +
                     '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                     'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                     'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "job has One Horizontal slot only per side "
                           "which is need to delete one of the slot in horizontal template manually." + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,3>), Except it's display on |
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if ((horizontal_slots_diameter_depth != 0 and horizontal_slots_diameter_depth is not None) and
                    (horizontal_slots_OD_spacing != 0 and horizontal_slots_OD_spacing is not None and
                     horizontal_slots_OD_spacing != "") and
                    (forging_inside_boss_spacing != 0 and forging_inside_boss_spacing is not None) and
                    (((horizontal_slots_OD_spacing - forging_inside_boss_spacing)/2)) > 1.00):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                 "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                 '[color=ffffff] job has Horizontal Slots OD Spacing each side longer than 1 Inch (length of the tool) '
                 'which is need to check if it is need to add extra pass in horizontal template manually. [/color]' +
                 '\n' + '\n' +
                 '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                 'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                 'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "job has Horizontal Slots OD Spacing each side longer than 1 Inch (length of the tool) "
                           "which is need to check if it is need to add extra pass in horizontal template manually." +
                    '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,4>), Except it's display on |
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if ((horizontal_slots_diameter_depth != 0 and horizontal_slots_diameter_depth is not None) and
                    (i_start_horizontal_slot == "" or j_start_horizontal_slot == "" or horizontal_slot_radius == "")):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                    '[color=ffffff] job has Horizontal Slots Of Pin Bore Size That has NO Horizontal '
                    'Slots Numbers (i,j,Radius) which is need to find and add the Numbers or Make Horizontal '
                    'Slots manually. [/color]' + '\n' + '\n' +
                    '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                    'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                    'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "job has Horizontal Slots Of Pin Bore Size That has NO Horizontal Slots Numbers (i,j,Radius)"
                           " which is need to find and add the Numbers or Make Horizontal Slots manually." + '\n' +
                    '\n' + "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,5>), Except it's display on |
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if (semi_cfren_ID_spacing != 0 and semi_cfren_ID_spacing is not None):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                     '[color=ffffff] job has Semi-C/Fren Grv '
                     'which is need to add the Semi-C/Fren values manually. [/color]' + '\n' + '\n' +
                     '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                     'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                     'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job has Semi-C/Fren Grv "
                           "which is need to add the Semi-C/Fren values manually." + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,6>), Except it's display on |
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if ((lock_ring_cutter_width != cfren_cutter_width) and
                    (cfren_ID_spacing != 0 and cfren_ID_spacing is not None) and
                    (lock_ring_cutter_width != 0 and lock_ring_cutter_width is not None)):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                     '[color=ffffff] job has Lockring width different than C/Fren width '
                     'which is need to customize the horizontal template manually. [/color]' + '\n' + '\n' +
                     '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                     'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                     'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job has Lockring width different than C/Fren width "
                    "which is need to customize the horizontal template manually." + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,7>), Except it's display on |
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if (pin_hole_diameter == 0.0 or pin_hole_diameter is None):
                warning_messages_of_creating_old_horizontal_machine_program.append(
                    '[b][i][u][color=0099ff]' + self.ids["JobNumberForOldHorizontalMachine"].text +
                    '[/color][/u][/i][/b]' + '[color=ffffff] job has Pin Hole Diameter Equal 0 or None, '
                                             'Double check Job Info and try to fix the issue.' + '\n' + '\n' +
                    '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                    'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                    'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job has Pin Hole Diameter Equal 0 or None, Double check Job Info and try to fix the issue. "
                    + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,8>), Except it's display on |
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if (pin_hole_diameter > 1.095):
                warning_messages_of_creating_old_horizontal_machine_program.append(
                    '[b][i][u][color=0099ff]' + self.ids["JobNumberForOldHorizontalMachine"].text +
                    '[/color][/u][/i][/b]' +
                    '[color=ffffff] job has pin hole diameter bigger than 1.095 which is need to customize the '
                    'horizontal template manually to rough the big bore diameter. [/color]' + '\n' + '\n' +
                    '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                    'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                    'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job has pin hole diameter bigger than 1.095 which is need to customize the horizontal "
                           "template manually to rough the big bore diameter." + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            #     STILL NEEDS TO CHECK WITH PROGRAMMING
            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,9>), Except it's display on |
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if (pin_hole_diameter < 0.4710):
                warning_messages_of_creating_old_horizontal_machine_program.append(
                    '[b][i][u][color=0099ff]' + self.ids[
                        "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                    '[color=ffffff] job has small uncommon Pin Hole Diameter Size ' + pin_hole_diameter +
                    ', Double check if we are able to run this Job. [/color]' + '\n' + '\n' +
                    '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                    'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                    'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job has small uncommon Pin Hole Diameter Size " + pin_hole_diameter +
                    ", Double check if we are able to run this Job. " + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,10>), Except it's display on|
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if (rough_bore_tool_number == 0):
                warning_messages_of_creating_old_horizontal_machine_program.append(
                    '[b][i][u][color=0099ff]' + self.ids["JobNumberForOldHorizontalMachine"].text +
                    '[/color][/u][/i][/b]' + '[color=ffffff] job can not find RoughBoreTool that fit the'
                                             ' FinishPinHoleDiameter size, maybe we can use swap tool T60 if applicable'
                    + '\n' + '\n' + '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                    'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                    'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job can not find RoughBoreTool that fit the FinishPinHoleDiameter size, "
                           "maybe we can use swap tool T60 if applicable. "
                    + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,11>), Except it's display on|
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if (piston_overall_length >= 4.300):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                     '[color=ffffff] job overall length longer than 4.300 '
                     'which is need to work on machine rotation manually to avoid crash. [/color]' + '\n' + '\n' +
                     '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                     'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                     'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job overall length longer than 4.300 "
                           "which is need to work on machine rotation manually to avoid crash." + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,12>), Except it's display on|
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if ((ledge_counterbore_diameter != 0 and ledge_counterbore_diameter is not None and
                 ledge_counterbore_diameter == "")):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                    '[color=ffffff] job seems has Ledge Counterbore, '
                    'which is need to add Counterbore Values (VC126 and VC179) manually, '
                    'and ID Spacing If Necessary. [/color]' + '\n' + '\n' +
                    '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                    'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                    'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job seems has Ledge Counterbore, "
                           "which is need to add Counterbore Values (VC126 and VC179) manually,"
                           "and ID Spacing If Necessary." + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,13>), Except it's display on|
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if ((ledge_counterbore_diameter != 0 and ledge_counterbore_diameter is not None and
                 ledge_counterbore_diameter == "") and (lock_ring_ID_spacing == 0 or lock_ring_ID_spacing is None)):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                    '[color=ffffff] job seems has Ledge Counterbore without LockRing which is need to customize '
                    'the horizontal template manually to add Counterbore ID Spacing. [/color]' + '\n' + '\n' +
                    '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                    'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                    'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job seems has Ledge Counterbore without LockRing "
                           "which is need to customize the horizontal template manually to add Counterbore ID Spacing."
                    + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,14>), Except it's display on|
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if ((forging_number != "" and forging_number is not None) and (forging_number in forged_forging_list)):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                    '[color=ffffff] job has Forged Forging which is need to add variable VC85 '
                    'and distance from pilot to top of forging manually to the program. [/color]' + '\n' + '\n' +
                    '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                    'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                    'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job has Forged Forging which is need to add variable VC85 "
                           "and distance from pilot to top of forging manually to the program." + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,15>), Except it's display on|
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if ((pin_hole_diameter != 0 and pin_hole_diameter != "" and pin_hole_diameter is not None and
                 pin_hole_diameter >= 0.901) and
                    (double_oil_hole_slot_ID_spacing != 0 and double_oil_hole_slot_ID_spacing is not None) and
                    (pin_hole_diameter not in double_oil_hole_slots_pin_sizes_list)):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                   '[color=ffffff] job has Pin Size that NOT include in DOHS Logic which is need to figure out '
                   'DOHS Numbers and add them to Template Logic manually. [/color]' + '\n' + '\n' +
                   '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                   'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                   'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job has Pin Size that NOT include in DOHS Logic which is need to figure out "
                           "DOHS Numbers and add them to Template Logic manually." + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,16>), Except it's display on|
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if (pilot_bore_depth == 0 or pilot_bore_depth is None):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                   '[color=ffffff] job has Pilot Bore Depth Equal 0 or None, Double check Job Info and try to fix '
                   'the issue.' + '\n' + '\n' + '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' +
                   '\n' + 'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                   'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job has Pilot Bore Depth Equal 0 or None, Double check Job Info and try to fix the issue. "
                    + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,17>), Except it's display on|
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if (X_distance_from_origin_to_pin_center == 0 or (
                    float(float(X_distance_from_origin_to_pin_center) * (-1)) > (-0.3))):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                   '[color=ffffff] job has xPinCenter Equal 0 or Bigger than -0.3, Double check Job Info and '
                   'try to fix the issue.' + '\n' + '\n' +
                   '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                   'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                   'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job has xPinCenter Equal 0 or Bigger than -0.3, Double check Job Info and "
                           "try to fix the issue. " + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,18>), Except it's display on|
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if (ledge_tool_diameter == 0):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                   '[color=ffffff] job has Ledge Tool Diameter Equal 0, Double check Job Info and try to fix the issue.'
                   + '\n' + '\n' +
                   '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                   'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                   'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job has Ledge Tool Diameter Equal 0, Double check Job Info and try to fix the issue. "
                    + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,19>), Except it's display on|
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if ((lock_ring_cutter_width != 0 and lock_ring_cutter_width is not None) and
                    (lock_ring_diameter <= pin_hole_diameter)):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                   '[color=ffffff] job has Lock Ring Diameter Smaller than Pin Hole Diameter, Double check Job Info '
                   'and try to fix the issue.' + '\n' + '\n' +
                   '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                   'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                   'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job has Lock Ring Diameter Smaller than Pin Hole Diameter, Double check Job Info and try "
                           "to fix the issue. " + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,20>), Except it's display on|
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if ((cfren_cutter_width != 0 and cfren_cutter_width is not None) and
                    (cfren_diameter <= pin_hole_diameter)):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                   '[color=ffffff] job has C/Fren Diameter Smaller than Pin Hole Diameter, Double check Job Info '
                   'and try to fix the issue.' + '\n' + '\n' +
                   '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                   'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                   'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job has C/Fren Diameter Smaller than Pin Hole Diameter, Double check Job Info and try "
                           "to fix the issue. " + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,21>), Except it's display on|
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if ((notch_angle_first_location != 0 and notch_angle_first_location is not None) and
                    (notch_angle_second_location != 0 and notch_angle_second_location is not None) and
                    (lock_ring_cutter_width == 0 or lock_ring_cutter_width is None)):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                   '[color=ffffff] job has CirclipNotch while has NO LockRing, Double check Job Info '
                   'and try to fix the issue.' + '\n' + '\n' +
                   '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                   'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                   'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job has CirclipNotch while has NO LockRing, Double check Job Info and try "
                           "to fix the issue. " + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # ===================================================================================================|
            # [#] Same Description of the above [Exceptional Programs Cases] (Note<2,22>), Except it's display on|
            #     the Screen Message Window of the App.                                                          |
            # ===================================================================================================|
            if ((X_distance_from_origin_to_circlip_notch != 0 and X_distance_from_origin_to_circlip_notch is not None)
                    and (float(float(X_distance_from_origin_to_circlip_notch) * (-1)) > (-0.3))):
                warning_messages_of_creating_old_horizontal_machine_program.append('[b][i][u][color=0099ff]' + self.ids[
                    "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                   '[color=ffffff] job has xCirclipNotch Bigger than -0.3, Double check Job Info '
                   'and try to fix the issue.' + '\n' + '\n' +
                   '[color=ffffff]UNFINISHED program created in ORIGINAL folder only, ' + '\n' +
                   'Needs to modify and save it again in ORIGINAL and RUNNING Folders. ' + '\n' +
                   'Add some Notes if needed. [/color]' + '\n')
                email_messages_of_creating_old_horizontal_machine_program.append(
                    '\n' + "Job has xCirclipNotch Bigger than -0.3, Double check Job Info and try "
                           "to fix the issue. " + '\n' + '\n' +
                    "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                    "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                    "Add some Notes if needed.")
                old_horizontal_exceptional_cases = True

            # (MAYBE) NEEDS TO ADD LOGIC TO **NOT** CREATE HORIZONTAL PROGRAM IN RUNNING FOLDER WHEN WE HAVE Warren Johnson Jobs
            # >>>>>> MAKE NORMAL PROGRAM, AND PROGRAMMER NEED TO CHECK EVERYTHING OR JUST TAKE OLD PROGRAM AND ADJUSTED
            # ---------------------------------------------------------------------------------------------------------

            # ============================================================================================|
            # [#] To Warn the User by display the messages of the exceptions on the Screen Message Window |
            #     of the App, and stop execute the code.                                                  |
            # [#] Call Function (program_needs_attention) to Warn the User.                               |
            # ============================================================================================|
            if (old_horizontal_exceptional_cases is True and
                    dont_create_old_horizontal_machine_program_in_original_folder is False):
                # =================================================================================================|
                # [#] If number of WarningMessages is more than <3>, show the User just general message of (having |
                #     some Exceptional Cases) because there is no enough room to display all of them on App        |
                #     Window, the User still can see all the WarningMessages on the program itself.                |
                # [#] Otherwise, display the messages of the exceptions on the Screen Message Window of the App.   |
                # =================================================================================================|
                if (len(warning_messages_of_creating_old_horizontal_machine_program) >= 3):
                    warning_messages_of_creating_old_horizontal_machine_program = [
                        '[b][i][u][color=0099ff]' + self.ids[
                            "JobNumberForOldHorizontalMachine"].text + '[/color][/u][/i][/b]' +
                        "[color=ffffff] Job has many exceptional cases, look to UNFINISHED program. "
                        + '\n' + '\n' + "UNFINISHED program created in ORIGINAL folder only," + '\n' +
                        "Needs to modify and save it again in ORIGINAL and RUNNING Folders." + '\n' + '\n' +
                        "Add some Notes if needed."]
                old_horizontal_machine_program_needs_attention(self)
                return
            # ==============================================================================================|
            # [#] If the User clicks on 'No_Button' for (Replace Program Question), and the job has some    |
            #     Exceptional Cases, therefore it will warn the User to pay attention for these cases on the|
            #     existing program on the Original folder.                                                  |
            # ==============================================================================================|
            elif (old_horizontal_exceptional_cases is True and
                  dont_create_old_horizontal_machine_program_in_original_folder is True):
                warning_messages_of_creating_old_horizontal_machine_program = [
                    '[b][i][u][color=0099ff]' + self.ids["JobNumberForOldHorizontalMachine"].text +
                    '[/color][/u][/i][/b]' +
                    "[color=ffffff] Job has some exceptional cases, try again and Create (or Replace) program in "
                    "ORIGINAL folder to be able to modify UNFINISHED program that created. " + '\n']

                # ============================================================================================|
                # [#] Set the Status to be "True", Just to make sure is not go through the (while) loop again.|
                # ============================================================================================|
                old_horizontal_exceptional_cases = True
                # ========================================================================================|
                # [#] Reset the Status to be "False", to try always create program on the original folder.|
                # ========================================================================================|
                dont_create_old_horizontal_machine_program_in_original_folder = False
                # =============================================================|
                # [#] Call Function (program_needs_attention) to Warn the User.|
                # =============================================================|
                old_horizontal_machine_program_needs_attention(self)
                return

            else:
                # ============================================================================================|
                # [#] Set the Status to be "True", Just to make sure is not go through the (while) loop again.|
                # ============================================================================================|
                old_horizontal_exceptional_cases = True
                # ========================================================================================|
                # [#] Reset the Status to be "False", to try always create program on the original folder.|
                # ========================================================================================|
                dont_create_old_horizontal_machine_program_in_original_folder = False

        # endregion  <<<<=======================[Exceptional Programs Cases <Messages>]========================>>>>

        # ===================================================================================================|
        # [#] After finishing the try of Creating program on the Original folder, call the                   |
        #     function (Create Program in Running Folder) to try to create the program on the Running folder.|
        # ===================================================================================================|
        create_old_horizontal_machine_program_in_running_folder(self)

    # ============================================================================|
    # [#] Create Function to Close the Message Window that related on the Running |
    #     folder and execute some actions.                                        |
    # ============================================================================|
    def close_old_horizontal_window_of_running_folder(self, obj):
        print("(close_old_horizontal_window_of_running_folder) Function >> Called")

        # =======================================|
        # [#] To Close the Screen Message Window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()
        # ====================================================================================================|
        # [#] Make Variable 'global' again to be able to Reset the list of (Warning Messages) if any Error may|
        #     occur when open the program by CIMCO Editor to avoid duplicate the warning message when close   |
        #     the Screen Message Window (Once 'close_old_horizontal_screen_window' function called).          |
        #  ===================================================================================================|
        global warning_messages_of_creating_old_horizontal_machine_program

        # ==========================================================================================|
        # [#] After Creating the program on the Running folder, it needs to open the program by     |
        #     CIMCO Editor (the software that used to create, edit and view the CNC Programs).      |
        # [#] Steps to open any application (software) on Windows Operative System:                 |
        #   [#] Use subprocess <built_in function in python>.                                       |
        #   [#] (Popen): Used to execute the 'open' functionality.                                  |
        #   [#] (ApplicationName): the kind of APP that need to use to open the file (like Notepad, |
        #       Excel, Word or CIMCO Editor in this case).                                          |
        #   [#] (FileName): the path of the file that's need to open.                               |
        # ==========================================================================================|

        # ================================================================================================|
        # [#] Check Offset direction to open all the programs that relate on the job number.              |
        #     (if the direction is "OFFSET EACH WAY", it will open two programs one for the <0> direction,|
        #      and one for the <180> direction, otherwise it will open one program).                      |
        # [#] Use (try/except) Blocks to Handle any Error may occur when open the program by CIMCO Editor.|
        # ================================================================================================|
        if (offset_direction == "OFFSET EACH WAY" and offset_amount != 0):
            try:
                subprocess.Popen([cimco_editor_path,
                                  new_horizontal_program_To0_direction_for_old_machine_in_running_folder])
                subprocess.Popen([cimco_editor_path,
                                  (new_horizontal_program_To180_direction_for_old_machine_in_running_folder)])
            except Exception as error:
                warning_messages_of_creating_old_horizontal_machine_program.append(
                    "Program has been created successfully, but Failed to Open it by" +
                    '[b][u][color=ffffff] CIMCO Editor. [/color][/u][/b]' + "\n" + "An Error has occurred :" + "\n" +
                    '[color=ff1a1a]' + str(error) + '[/color]' + "\n" +
                    "Double Check Network, and CIMCO Editor Location.")
                email_messages_of_creating_old_horizontal_machine_program.append(
                    "\n" + "But Failed to Open it by CIMCO Editor." + "\n" + "An Error has occurred :" + "\n" + str(
                        error) + "\n" + "Double Check Network, and CIMCO Editor Location.")
                old_horizontal_machine_program_needs_attention(self)
                # ===============================================================================================|
                # [#] Reset the list of (Warning Messages) once Error occur when try to open the program by CIMCO|
                #     Editor to avoid duplicate the warning message when close the Screen Message Window         |
                #     (Once 'close_old_horizontal_screen_window' function called).                               |
                #  ==============================================================================================|
                warning_messages_of_creating_old_horizontal_machine_program = []
                return
        else:
            try:
                subprocess.Popen([cimco_editor_path, new_horizontal_program_for_old_machine_in_running_folder])
            except Exception as error:
                warning_messages_of_creating_old_horizontal_machine_program.append(
                    "Program has been created successfully, but Failed to Open it by" +
                    '[b][u][color=ffffff] CIMCO Editor. [/color][/u][/b]' + "\n" + "An Error has occurred :" + "\n" +
                    '[color=ff1a1a]' + str(error) + '[/color]' + "\n" +
                    "Double Check Network, and CIMCO Editor Location.")
                email_messages_of_creating_old_horizontal_machine_program.append(
                    "\n" + "\n" + "But Failed to Open it by CIMCO Editor." + "\n" + "An Error has occurred :" + "\n" +
                    str(error) + "\n" + "Double Check Network, and CIMCO Editor Location.")
                old_horizontal_machine_program_needs_attention(self)
                # ===============================================================================================|
                # [#] Reset the list of (Warning Messages) once Error occur when try to open the program by CIMCO|
                #     Editor to avoid duplicate the warning message when close the Screen Message Window         |
                #     (Once 'close_old_horizontal_screen_window' function called).                               |
                #  ==============================================================================================|
                warning_messages_of_creating_old_horizontal_machine_program = []
                return

        # =======================================================================================================|
        # [#] After finishing the try of Creating program on the Running folder, call the function               |
        #      (Send Email of Creating Program) to inform the User about all the details that related on the job.|
        # =======================================================================================================|
        send_email_about_create_old_horizontal_machine_program(self)
        print("\n".join(email_messages_of_creating_old_horizontal_machine_program))

        # ========================================================================================|
        # [#] Reset Job Number TextInput Field after finishing creating the program to start over.|
        # ========================================================================================|
        self.ids["JobNumberForOldHorizontalMachine"].text = ""

        # =======================================================================|
        # [#] Reset Variables after finishing creating the program to start over.|
        # =======================================================================|
        four_cycle_pin_bore_variables(self)

    # ===================================================================================|
    # [#] Create Function to Close the Message Window that related on the Old Horizontal |
    #     Screen and execute some actions.                                               |
    # ===================================================================================|
    def close_old_horizontal_screen_window(self, obj):
        print("(close_old_horizontal_screen_window) Function >> called")

        # =======================================|
        # [#] To Close the Screen Message Window.|
        # =======================================|
        self.old_horizontal_screen_message_window.dismiss()

        # ====================================================================================================|
        # [#] Make Variable 'global' again to be able to Reset the list of (Warning Messages) if any Error may|
        #     occur when open the program by CIMCO Editor to avoid duplicate the warning message when close   |
        #     the Screen Message Window (Once 'close_old_horizontal_screen_window' function called).          |
        #  ===================================================================================================|
        global warning_messages_of_creating_old_horizontal_machine_program

        # ======================================================================================================|
        # [#] If job number TextInput field is Empty (User didn't enter a value), then NO need to do any action.|
        # ======================================================================================================|
        if (self.ids["JobNumberForOldHorizontalMachine"].text == ""):
            pass
        # ===============================================================|
        # [#] If the User is 'Guest User', then NO need to do any action.|
        # ===============================================================|
        elif (self.manager.get_screen('LoginScreen').ids["Email"].text in guests_email_address_list):
            pass
        # =======================================================================================================|
        # [#] If it failed to create the program, send email to Trello board to inform the user with all details.|
        # =======================================================================================================|
        elif (fail_messages_of_creating_old_horizontal_machine_program != []):
            print("\n".join(email_messages_of_creating_old_horizontal_machine_program))
            send_email_about_create_old_horizontal_machine_program(self)
        # ====================================================================================================|
        # [#] If it has some warning or confirmation details to create the program, therefore open the program|
        #     that created on Original folder to allow the user to modify whatever it's needed, and send email|
        #     to Trello board to inform the user with all details.                                            |
        # ====================================================================================================|
        elif (warning_messages_of_creating_old_horizontal_machine_program != []):
            print("\n".join(email_messages_of_creating_old_horizontal_machine_program))

            # ================================================================================================|
            # [#] Check Offset direction to open all the programs that relate on the job number.              |
            #     (if the direction is "OFFSET EACH WAY", it will open two programs one for the <0> direction,|
            #      and one for the <180> direction, otherwise it will open one program).                      |
            # [#] Use (try/except) Blocks to Handle any Error may occur when open the program by CIMCO Editor.|
            # ================================================================================================|
            if (offset_direction == "OFFSET EACH WAY" and offset_amount != 0):
                try:
                    subprocess.Popen([cimco_editor_path,
                                      new_horizontal_program_To0_direction_for_old_machine_in_original_folder])
                    subprocess.Popen([cimco_editor_path,
                                      (new_horizontal_program_To180_direction_for_old_machine_in_original_folder)])
                except Exception as error:
                    warning_messages_of_creating_old_horizontal_machine_program.append(
                        "Program has been created successfully, but Failed to Open it by" +
                        '[b][u][color=ffffff] CIMCO Editor. [/color][/u][/b]' + "\n" + "An Error has occurred :" + "\n"
                        + '[color=ff1a1a]' + str(error) + '[/color]' + "\n" +
                        "Double Check Network, and CIMCO Editor Location.")
                    email_messages_of_creating_old_horizontal_machine_program.append(
                        "\n" + "But Failed to Open it by CIMCO Editor." + "\n" + "An Error has occurred :" + "\n" + str(
                            error) + "\n" + "Double Check Network, and CIMCO Editor Location.")
                    old_horizontal_machine_program_needs_attention(self)
                    send_email_about_create_old_horizontal_machine_program(self)
                    # ===============================================================================================|
                    # [#] Reset the list of (Warning Messages) once Error occur when try to open the program by CIMCO|
                    #     Editor to avoid duplicate the warning message when close the Screen Message Window         |
                    #     (Once 'close_old_horizontal_screen_window' function called again).                         |
                    #  ==============================================================================================|
                    warning_messages_of_creating_old_horizontal_machine_program = []
                    return
            else:
                try:
                    subprocess.Popen([cimco_editor_path, new_horizontal_program_for_old_machine_in_original_folder])
                except Exception as error:
                    warning_messages_of_creating_old_horizontal_machine_program.append(
                        "Program has been created successfully, but Failed to Open it by" +
                        '[b][u][color=ffffff] CIMCO Editor. [/color][/u][/b]' + "\n" + "An Error has occurred :" + "\n"
                        + '[color=ff1a1a]' + str(error) + '[/color]' + "\n" +
                        "Double Check Network, and CIMCO Editor Location.")

                    email_messages_of_creating_old_horizontal_machine_program.append(
                        "\n" + "But Failed to Open it by CIMCO Editor." + "\n" + "An Error has occurred :" + "\n" + str(
                            error) + "\n" + "Double Check Network, and CIMCO Editor Location.")
                    old_horizontal_machine_program_needs_attention(self)
                    send_email_about_create_old_horizontal_machine_program(self)
                    # ===============================================================================================|
                    # [#] Reset the list of (Warning Messages) once Error occur when try to open the program by CIMCO|
                    #     Editor to avoid duplicate the warning message when close the Screen Message Window         |
                    #     (Once 'close_old_horizontal_screen_window' function called again).                         |
                    #  ==============================================================================================|
                    warning_messages_of_creating_old_horizontal_machine_program = []
                    return

        # ===================================================|
        # [#] Reset Job Number TextInput Field to start over.|
        # ===================================================|
        self.ids["JobNumberForOldHorizontalMachine"].text = ""

        # ===================================================|
        # [#] Reset the status to be "False" to start over.  |
        # ===================================================|
        global called_need_confirmation_to_create_old_horizontal_machine_program
        called_need_confirmation_to_create_old_horizontal_machine_program = False
        print(
            "called_need_confirmation_to_create_old_horizontal_machine_program on (close_old_horizontal) Function: ",
            called_need_confirmation_to_create_old_horizontal_machine_program)

        # ========================================================================================|
        # [#] Reset Variables to start over.                                                      |
        # [#] Check (forge_spec_id) if it's NOT None to avoid call (four_cycle_pin_bore_variables)|
        #     function when forging ID is not exist to avoid any Error may occurs                 |
        # ========================================================================================|
        if (forge_spec_id is not None):
            four_cycle_pin_bore_variables(self)

    # =========================================================================|
    # [#] Create Function to Reset the Job Number TextInput Field to start over|
    #     once the User clicks on 'Home' or 'Back' buttons.                    |
    # =========================================================================|
    def reset_old_horizontal_screen_fields(self):
        self.ids["JobNumberForOldHorizontalMachine"].text = ""

    # endregion <<<<===========================[Old Horizontal Machines Sub Functions]===========================>>>>


# endregion <<<<===========================[Old Horizontal Machines(28,29,32) Screen]============================>>>>


# region <<<<============================[New Horizontal Machines(127) Items List]===========================>>>>

# ========================================================================================================|
# [#] Create the Class to be able to create List's Items that will be used to let user choose from options|
#     that are needed for some confirmation details that need to finalize by the user.                    |
# ========================================================================================================|
class NewHorizontalMachineItem(OneLineAvatarListItem):
    divider = None

# endregion <<<<==========================[New Horizontal Machines(127) Items List]==========================>>>>


# region <<<<===========================[New Horizontal Machines(127) Functions]============================>>>>
# still need work

# def create_new_horizontal_machine_program_in_original_folder(self):
#     print("create_new_horizontal_machine_program_in_original_folder" + " FUNCTION is called")
#
#     # DEFINE VARIABLES ON THIS FUNCTION AS global TO BE ABLE USE THEM OUT SIDE THE FUNCTION
#     global original_folder_path_of_new_horizontal_machine
#     original_folder_path_of_new_horizontal_machine = self.manager.get_screen('OldHorizontalSettingScreen').ids[
#         "OriginalFolderPathOfOldHorizontalMachine"].text
#
#     global new_horizontal_program_for_new_machine_in_original_folder
#     global new_horizontal_program_To0_direction_for_new_machine_in_original_folder
#     global new_horizontal_program_To180_direction_for_new_machine_in_original_folder
#
#     # WE USE COPY METHOD TO COPY (HorizontalProgramLines) THAT'S CONTAIN ORIGINAL LINES BEFORE
#     # MAKE 2 SEPARATE PROGRAMS To0 AND To180
#     # WE MAKE NEW LISTS TO BE ABLE TO MODIFY THEM WITHOUT ADJUST THE ORIGINAL LIST(HorizontalProgramLines)
#     # BECAUSE WE WANT TO USE IT LATER
#     global horizontal_program_lines_To0_direction_for_new_machine_in_original_folder
#     horizontal_program_lines_To0_direction_for_new_machine_in_original_folder = HorizontalProgramLines.copy()
#
#     global horizontal_program_lines_To180_direction_for_new_machine_in_original_folder
#     horizontal_program_lines_To180_direction_for_new_machine_in_original_folder = HorizontalProgramLines.copy()
#
#     # TO CREATE THE NEW FILE ON THE PATH YOU WANT
#     # open() WITH "x" IT will create a file, returns an error if the file exist(THAT WHY WE USE try/except)
#     # IF IT DOES NOT EXIST IT WILL CREATE THE NEW FILE , IF IT IS EXIST IT WILL GO TO except BLOCK
#     # AND CHECK IF NEED TO SAVE OVER THE EXISTING FILE
#     try:
#         print("try original is called")
#         if (offset_direction == "OFFSET EACH WAY" and offset_amount != 0):
#
#             # To0 PRROGRAM
#
#             # print("ORIGINAL_FOLDER_TO0")
#             # print(HorizontalProgramLines)
#             new_horizontal_program_To0_direction_for_new_machine_in_original_folder = (
#                     original_folder_path_of_new_horizontal_machine + "\\" + "P" +
#                     new_program_number_for_new_horizontal_machine + "TO0.MIN")
#             horizontal_program_lines_To0_direction_for_new_machine_in_original_folder[0] = (
#                     '(PART ' + new_program_number_for_new_horizontal_machine + " TO 0" + ' -- ' +
#                     todaydate + ' <SYS>' + ')')
#
#             try:
#                 create_new_horizontal_program_To0_direction_for_new_machine_in_original_folder = open(
#                     new_horizontal_program_To0_direction_for_new_machine_in_original_folder, "x")
#                 create_new_horizontal_program_To0_direction_for_new_machine_in_original_folder.write(
#                     '\n'.join(horizontal_program_lines_To0_direction_for_new_machine_in_original_folder))
#                 create_new_horizontal_program_To0_direction_for_new_machine_in_original_folder.close()
#             # TO HANDLE ERROR OF NOT FINDING ORIGINAL FOLDER
#             except PermissionError or FileNotFoundError as error:
#                 fail_messages_of_creating_new_horizontal_machine_program.append(
#                     "Failed to Find" + '[b][color=ffff00] ORIGINAL [/color][/b]' +
#                     "Folder location to Save the Program." + "\n" + "An Error has occurred :"
#                     + "\n" + '[color=ff1a1a]'
#                     + str(error) + '[/color]' + "\n" + "Double Check Network, and Folder Location.")
#                 email_messages_of_creating_new_horizontal_machine_program.append(
#                     "Failed to Find Original Folder location to Save the Program." + "\n" +
#                     "Double Check Network, and File Location.")
#                 failed_to_create_new_horizontal_machine_program(self)
#                 return
#
#             # TO 180
#
#             # print("ORIGINAL_FOLDER_TO180")
#             # print(HorizontalProgramLines)
#
#             new_horizontal_program_To180_direction_for_new_machine_in_original_folder = (
#                     original_folder_path_of_new_horizontal_machine + "\\" + "P" +
#                     new_program_number_for_new_horizontal_machine + "TO180.MIN")
#             horizontal_program_lines_To180_direction_for_new_machine_in_original_folder[0] = (
#                     '(PART ' + new_program_number_for_new_horizontal_machine + " TO 180" + ' -- ' +
#                     todaydate + ' <SYS>' + ')')
#             horizontal_program_lines_To180_direction_for_new_machine_in_original_folder[VC155_VARIABLE_INDEX] = (
#                     'VC155=' + format((-1) * offset_amount) + '  (Offset)')
#             # print("OFFSET_AMOUNT for original folder", OFFSET_AMOUNT)
#             horizontal_program_lines_To180_direction_for_new_machine_in_original_folder[VC173_VARIABLE_INDEX] = (
#                     'VC173=' + format(float(Y_distance_from_origin_to_circlip_notch) - (2 * offset_amount),
#                                       '.4f') + '  (yCirclipNotch)')
#
#             try:
#                 create_new_horizontal_program_To180_direction_for_new_machine_in_original_folder = open(
#                     new_horizontal_program_To180_direction_for_new_machine_in_original_folder, "x")
#                 create_new_horizontal_program_To180_direction_for_new_machine_in_original_folder.write(
#                     '\n'.join(horizontal_program_lines_To180_direction_for_new_machine_in_original_folder))
#                 create_new_horizontal_program_To180_direction_for_new_machine_in_original_folder.close()
#             # TO HANDLE ERROR OF NOT FINDING ORIGINAL FOLDER
#             except PermissionError or FileNotFoundError as error:
#                 # print("except LOAD_HORIZONTAL_SHEETS_FOR_AUTOMATION")
#                 fail_messages_of_creating_new_horizontal_machine_program.append(
#                     "Failed to Find" + '[b][color=ffff00] ORIGINAL [/color][/b]' +
#                     "Folder location to Save the Program." + "\n" + "An Error has occurred :" + "\n" +
#                     '[color=ff1a1a]'
#                     + str(error) + '[/color]' + "\n" + "Double Check Network, and Folder Location.")
#                 email_messages_of_creating_new_horizontal_machine_program.append(
#                     "Failed to Find Original Folder location to Save the Program." + "\n" +
#                     "Double Check Network, and File Location.")
#                 failed_to_create_new_horizontal_machine_program(self)
#                 return
#
#         # IF OFFSET DIRECTION IS NOT EACHWAY
#         else:
#
#             new_horizontal_program_for_new_machine_in_original_folder = (
#                     original_folder_path_of_new_horizontal_machine + "\\" + "P" +
#                     new_program_number_for_new_horizontal_machine + ".MIN")
#             try:
#                 create_new_horizontal_program_for_new_machine_in_original_folder = open(
#                     new_horizontal_program_for_new_machine_in_original_folder, "x")
#                 create_new_horizontal_program_for_new_machine_in_original_folder.write(
#                     '\n'.join(HorizontalProgramLines))
#                 create_new_horizontal_program_for_new_machine_in_original_folder.close()
#             # TO HANDLE ERROR OF NOT FINDING ORIGINAL FOLDER
#             except PermissionError or FileNotFoundError as error:
#                 # print("except LOAD_HORIZONTAL_SHEETS_FOR_AUTOMATION")
#                 fail_messages_of_creating_new_horizontal_machine_program.append(
#                     "Failed to Find" + '[b][color=ffff00] ORIGINAL [/color][/b]' +
#                     "Folder location to Save the Program." + "\n" + "An Error has occurred :" +
#                     "\n" + '[color=ff1a1a]'
#                     + str(error) + '[/color]' + "\n" + "Double Check Network, and Folder Location.")
#                 email_messages_of_creating_new_horizontal_machine_program.append(
#                     "Failed to Find Original Folder location to Save the Program." + "\n" +
#                     "Double Check Network, and File Location.")
#                 failed_to_create_new_horizontal_machine_program(self)
#                 return
#
#         #  TO ADD THE CONFIRMATION DETAILS THAT WAS NEEDED TO CREATE THE PROGRAM
#         if (new_horizontal_program_confirmation_email_message_list != []):
#             email_messages_of_creating_new_horizontal_machine_program.append(
#                 "\n".join(new_horizontal_program_confirmation_email_message_list) + "\n")
#         success_messages_of_creating_new_horizontal_machine_program = \
#             ["\n" + "Program has been **CREATED** successfully in **ORIGINAL** Folder." + "\n"]
#
#         email_messages_of_creating_new_horizontal_machine_program.append("\n".join(
#             success_messages_of_creating_new_horizontal_machine_program))
#
#         close_button = MDRaisedButton(text='Close', on_release=self.close_new_horizontal_window_of_original_folder,
#                                       font_size=16)
#         self.new_Horizontal_Message_Dialog = MDDialog(title='[color=248f24]Success Message[/color]', text=(
#                 '[b][i][u][color=0099ff]' + self.ids["JobNumber"].text + '[/color][/u][/i][/b]' +
#                 '[color=ffffff] Program Has been Created Successfully in [/color]' +
#                 '[color=ffff00]ORIGINAL[/color]' +
#                 '[color=ffffff] Folder.[/color]'), size_hint=(0.7, 1.0), buttons=[close_button],auto_dismiss=False)
#         # TO OPEN THE DIALOG WINDOW
#         self.new_Horizontal_Message_Dialog.open()
#         print()
#         # print('\n'.join(HorizontalProgramLines))
#     # IF PROGRAM IS EXIST IT WILL CHECK IF NEED TO SAVE OVER THE EXISTING FILE
#     except(FileExistsError):
#         print("exept original is called")
#         yes_button = MDRaisedButton(text='Yes',
#                                     on_release=self.replace_existing_new_horizontal_machine_program_in_original_folder,
#                                     font_size=16)
#         no_button = MDRaisedButton(text='No', on_release=self.close_new_horizontal_window_of_original_folder,
#                                    font_size=16)
#         self.new_Horizontal_Message_Dialog = MDDialog(title='[color=0066ff]Confirmation Message[/color]',
#                                                       text=('[b][i][u][color=0099ff]' + self.ids[
#                                                           "JobNumber"].text + '[/color][/u][/i][/b]' +
#                                                             '[color=ffffff] Program already Exists in [/color]' +
#                                                             '[color=ffff00]ORIGINAL[/color]' +
#                                                             '[color=ffffff] Folder.[/color]' + '\n' +
#                                                             '[color=ffffff]Do you want to replace it ?[/color]'),
#                                                       size_hint=(0.7, 1.0), buttons=[yes_button, no_button],
#                                                       auto_dismiss=False)
#         # TO OPEN THE DIALOG WINDOW
#         self.new_Horizontal_Message_Dialog.open()
#
#         # MAYBE WE DO NOT NEED IT
#         # TO CLEAR THE LIST TO BE EMPTY IF ALREADY THE PROGRAM EXIST
#         success_messages_of_creating_new_horizontal_machine_program = []
#         # maybe need logic to send this message to trello in case there is warnning make this program
#         email_messages_of_creating_new_horizontal_machine_program.append("\n".join(
#             success_messages_of_creating_new_horizontal_machine_program))
#
#
# def create_new_horizontal_machine_program_in_running_folder(self):
#     print("create_new_horizontal_machine_program_in_running_folder" + " FUNCTION is called")
#     global running_folder_path_of_new_horizontal_machine
#     running_folder_path_of_new_horizontal_machine = self.manager.get_screen('OldHorizontalSettingScreen').ids[
#         "RunningFolderPathOfOldHorizontalMachine"].text
#     global new_horizontal_program_for_new_machine_in_running_folder
#     global new_horizontal_program_To0_direction_for_new_machine_in_running_folder
#     global new_horizontal_program_To180_direction_for_new_machine_in_running_folder
#
#     global horizontal_program_lines_To0_direction_for_new_machine_in_running_folder
#     horizontal_program_lines_To0_direction_for_new_machine_in_running_folder = HorizontalProgramLines.copy()
#
#     global horizontal_program_lines_To180_direction_for_new_machine_in_running_folder
#     horizontal_program_lines_To180_direction_for_new_machine_in_running_folder = HorizontalProgramLines.copy()
#
#     try:
#         print("try running is called")
#         if (offset_direction == "OFFSET EACH WAY" and offset_amount != 0):
#             # TO 0
#             # print("RUNNING_FOLDER_TO0")
#             # print(HorizontalProgramLines)
#             new_horizontal_program_To0_direction_for_new_machine_in_running_folder = (
#                     running_folder_path_of_new_horizontal_machine + "\\" + "P" +
#                     new_program_number_for_new_horizontal_machine + "TO0.MIN")
#             horizontal_program_lines_To0_direction_for_new_machine_in_running_folder[0] = (
#                     '(PART ' + new_program_number_for_new_horizontal_machine + " TO 0" + ' -- ' +
#                     todaydate + ' <SYS>' + ')')
#
#             try:
#                 create_new_horizontal_program_To0_direction_for_new_machine_in_running_folder = open(
#                     new_horizontal_program_To0_direction_for_new_machine_in_running_folder, "x")
#                 create_new_horizontal_program_To0_direction_for_new_machine_in_running_folder.write(
#                     '\n'.join(horizontal_program_lines_To0_direction_for_new_machine_in_running_folder))
#                 create_new_horizontal_program_To0_direction_for_new_machine_in_running_folder.close()
#             except PermissionError or FileNotFoundError as error:
#                 # print("except LOAD_HORIZONTAL_SHEETS_FOR_AUTOMATION")
#                 fail_messages_of_creating_new_horizontal_machine_program.append(
#                     "Failed to Find" + '[b][color=33cc33] RUNNING [/color][/b]' +
#                     "Folder location to Save the Program."
#                     + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]' + str(
#                         error) + '[/color]' + "\n" + "Double Check Network, and Folder Location.")
#                 email_messages_of_creating_new_horizontal_machine_program.append(
#                     "Failed to Find Running Folder location to Save the Program." + "\n" +
#                     "Double Check Network, and File Location.")
#                 failed_to_create_new_horizontal_machine_program(self)
#                 return
#
#             # TO 180
#             # print("RUNNING_FOLDER_TO180")
#             # print(HorizontalProgramLines)
#             new_horizontal_program_To180_direction_for_new_machine_in_running_folder = (
#                     running_folder_path_of_new_horizontal_machine + "\\" + "P" +
#                     new_program_number_for_new_horizontal_machine + "TO180.MIN")
#             horizontal_program_lines_To180_direction_for_new_machine_in_running_folder[0] = (
#                     '(PART ' + new_program_number_for_new_horizontal_machine + " TO 180" + ' -- ' + todaydate +
#                     ' <SYS>' + ')')
#             horizontal_program_lines_To180_direction_for_new_machine_in_running_folder[VC155_VARIABLE_INDEX] = (
#                     'VC155=' + format((-1) * offset_amount) + '  (Offset)')
#             horizontal_program_lines_To180_direction_for_new_machine_in_running_folder[VC173_VARIABLE_INDEX] = (
#                     'VC173=' + format(float(Y_distance_from_origin_to_circlip_notch) - (2 * offset_amount),
#                                       '.4f') + '  (yCirclipNotch)')
#
#             try:
#                 create_new_horizontal_program_To180_direction_for_new_machine_in_running_folder = open(
#                     new_horizontal_program_To180_direction_for_new_machine_in_running_folder, "x")
#                 create_new_horizontal_program_To180_direction_for_new_machine_in_running_folder.write(
#                     '\n'.join(horizontal_program_lines_To180_direction_for_new_machine_in_running_folder))
#                 create_new_horizontal_program_To180_direction_for_new_machine_in_running_folder.close()
#             except PermissionError or FileNotFoundError as error:
#                 # print("except LOAD_HORIZONTAL_SHEETS_FOR_AUTOMATION")
#                 fail_messages_of_creating_new_horizontal_machine_program.append(
#                     "Failed to Find" + '[b][color=33cc33] RUNNING [/color][/b]' +
#                     "Folder location to Save the Program."
#                     + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]' + str(
#                         error) + '[/color]' + "\n" + "Double Check Network, and Folder Location.")
#                 email_messages_of_creating_new_horizontal_machine_program.append(
#                     "Failed to Find Running Folder location to Save the Program." + "\n" +
#                     "Double Check Network, and File Location.")
#                 failed_to_create_new_horizontal_machine_program(self)
#                 return
#
#         else:
#             new_horizontal_program_for_new_machine_in_running_folder = (
#                     running_folder_path_of_new_horizontal_machine + "\\" + "P" +
#                     new_program_number_for_new_horizontal_machine + ".MIN")
#             try:
#                 create_new_horizontal_program_for_new_machine_in_running_folder = open(
#                     new_horizontal_program_for_new_machine_in_running_folder, "x")
#                 create_new_horizontal_program_for_new_machine_in_running_folder.write('\n'.join(HorizontalProgramLines))
#                 create_new_horizontal_program_for_new_machine_in_running_folder.close()
#             except PermissionError or FileNotFoundError as error:
#                 # print("except LOAD_HORIZONTAL_SHEETS_FOR_AUTOMATION")
#                 fail_messages_of_creating_new_horizontal_machine_program.append(
#                     "Failed to Find" + '[b][color=33cc33] RUNNING [/color][/b]' +
#                     "Folder location to Save the Program."
#                     + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]' + str(
#                         error) + '[/color]' + "\n" + "Double Check Network, and Folder Location.")
#                 email_messages_of_creating_new_horizontal_machine_program.append(
#                     "Failed to Find Running Folder location to Save the Program." + "\n" +
#                     "Double Check Network, and File Location.")
#                 failed_to_create_new_horizontal_machine_program(self)
#                 return
#
#         # JUST FOR NOW
#         success_messages_of_creating_new_horizontal_machine_program = \
#             ["\n" + "Program has been **CREATED** successfully in **RUNNING** Folder."]
#         email_messages_of_creating_new_horizontal_machine_program.append("\n".join(
#             success_messages_of_creating_new_horizontal_machine_program))
#
#         close_button = MDRaisedButton(text='Close', on_release=self.close_new_horizontal_window_of_running_folder,
#                                       font_size=16)
#         self.new_Horizontal_Message_Dialog = MDDialog(title='[color=248f24]Success Message[/color]', text=(
#                 '[b][i][u][color=0099ff]' + self.ids["JobNumber"].text + '[/color][/u][/i][/b]' +
#                 '[color=ffffff] Program Has been Created Successfully in [/color]' +
#                 '[color=33cc33]RUNNING[/color]' + '[color=ffffff] Folder.[/color]' + '\n' +
#                 '[color=ffffff]After closing this window, the program will open on CIMCO Editor.[/color]' + '\n' +
#                 '[color=ffffff]Please Double Check and Report any Problem on Trello.[/color]'),
#                                                     size_hint=(0.7, 1.0), buttons=[close_button], auto_dismiss=False)
#         # TO OPEN THE DIALOG WINDOW
#         self.new_Horizontal_Message_Dialog.open()
#         print()
#         # print('\n'.join(HorizontalProgramLines))
#     except(FileExistsError):
#         print("exept running is called")
#         # print(" Program is Exist in Running Folder")
#         yes_button = MDRaisedButton(text='Yes',
#                                     on_release=self.replace_existing_new_horizontal_machine_program_in_running_folder,
#                                     font_size=16)
#         no_button = MDRaisedButton(text='No', on_release=self.close_new_horizontal_screen_window, font_size=16)
#         self.new_Horizontal_Message_Dialog = MDDialog(title='[color=0066ff]Confirmation Message[/color]',
#                                                       text=('[b][i][u][color=0099ff]' + self.ids[
#                                                           "JobNumber"].text + '[/color][/u][/i][/b]' +
#                                                             '[color=ffffff] Program already Exists in [/color]' +
#                                                             '[color=33cc33]Running[/color]' +
#                                                             '[color=ffffff] Folder.[/color]' + '\n' +
#                                                             '[color=ffffff]Do you want to replace it ?[/color]'),
#                                                       size_hint=(0.7, 1.0), buttons=[yes_button, no_button],
#                                                       auto_dismiss=False)
#         # TO OPEN THE DIALOG WINDOW
#         self.new_Horizontal_Message_Dialog.open()
#         # MAYBE WE DO NOT NEED IT
#         # TO CLEAR THE LIST TO BE EMPTY IF ALREADY THE PROGRAM IS EXIST
#         success_messages_of_creating_new_horizontal_machine_program = []
#         email_messages_of_creating_new_horizontal_machine_program.append("\n".join(
#             success_messages_of_creating_new_horizontal_machine_program))
#
#
# def need_confirmation_to_create_new_horizontal_machine_program(self, title, sub_function, dialog_type, content):
#     print("need_confirmation_to_create_new_horizontal_machine_program" + " FUNCTION is called")
#     print('\n'.join(verification_messages_of_creating_new_horizontal_machine_program))
#     # still need to add action to do
#     enter_button = MDRaisedButton(text='Enter', on_press=sub_function,
#                                   on_release=self.create_program_for_new_horizontal_machine,
#                                   font_size=16)
#     close_button = MDRaisedButton(text='Close', on_release=self.close_new_horizontal_screen_window, font_size=16)
#     # IF CONFIRMATION NEED TO ENTER VALUE
#     if (dialog_type == "custom"):
#         self.new_Horizontal_Message_Dialog = MDDialog(title=title, type=dialog_type, content_cls=content,
#                                                       size_hint=(0.7, 1.0), buttons=[enter_button, close_button],
#                                                       auto_dismiss=False)
#     # IF CONFIRMATION NEED TO CHOOSE VALUE FROM OPTION
#     elif (dialog_type == "confirmation"):
#         self.new_Horizontal_Message_Dialog = MDDialog(title=title, type=dialog_type, items=content,
#                                                       size_hint=(0.7, 1.0), buttons=[enter_button, close_button],
#                                                       auto_dismiss=False)
#     # TO OPEN THE DIALOG WINDOW
#     self.new_Horizontal_Message_Dialog.open()
#
#
# def new_horizontal_machine_program_needs_attention(self):
#     print("old_horizontal_machine_program_needs_attention" + " FUNCTION is called")
#     close_button = MDRaisedButton(text='Close', on_release=self.close_new_horizontal_screen_window, font_size=16)
#     self.new_Horizontal_Message_Dialog = MDDialog(title='[color=cccc00]Warning Message[/color]', text=(
#             '[color=ffffff]' + '\n'.join(warning_messages_of_creating_new_horizontal_machine_program) + '[/color]'),
#                                                   size_hint=(0.7, 1.0), buttons=[close_button], auto_dismiss=False)
#     # TO OPEN THE DIALOG WINDOW
#     self.old_horizontal_screen_message_window.open()
#
#
# def failed_to_create_new_horizontal_machine_program(self):
#     print("failed_to_create_old_horizontal_machine_program" + " FUNCTION is called")
#     # print('\n'.join(fail_messages_of_creating_old_horizontal_machine_program))
#     close_button = MDRaisedButton(text='Close', on_release=self.close_new_horizontal_screen_window, font_size=16)
#     self.new_Horizontal_Message_Dialog = MDDialog(title='[color=990000]Warning Message[/color]', text=(
#             '[color=ffffff]' + '\n'.join(fail_messages_of_creating_new_horizontal_machine_program) + '[/color]'),
#                                                   size_hint=(0.7, 1.0), buttons=[close_button], auto_dismiss=False)
#     # TO OPEN THE DIALOG WINDOW
#     self.new_Horizontal_Message_Dialog.open()
#
#
# def send_email_about_create_new_horizontal_machine_program(self):
#     try:
#         service_app = "outlook.office365.com"
#         # SERVER NAME
#         smtp_server = "smtp.office365.com"
#         port = 587
#         # IT WILL USE EMAIL OF USER WHO LOGIN TO THE APP TO SENT THE EMAIL TO THE TRELLO BOARD
#         sender_email = user_email_address
#         # USING KEYRING PACKAGE TO GET PASSWORD FROM Windows Credential Manager WHILE THEY ARE SAVING IN USER COMPUTER
#         sender_password = keyring.get_password(service_app, sender_email)
#         # just for now
#         # recipient_email = "moemenatweh@hotmail.com"
#         trello_board_email = self.manager.get_screen('AppSettingScreen').ids[
#             "TrelloEmailAddress"].text  # moemenalatweh1+sqa4wcni54jz6erwnpbj@boards.trello.com
#         # creates SMTP session
#         email_server = smtplib.SMTP(smtp_server, port)
#         # TLS for security
#         email_server.starttls()
#
#         # authentication
#         # compiler gives an error for wrong credential.
#         email_server.login(sender_email, sender_password)
#
#         #########
#         if (fail_messages_of_creating_new_horizontal_machine_program != []):
#             card_title = "Failed to Create Program for " + new_program_number_for_new_horizontal_machine
#             card_label = "#FAILED"
#             card_member = "@moemenalatweh1 " + "@" + trello_user_name
#         elif (warning_messages_of_creating_new_horizontal_machine_program != []):
#             card_title = new_program_number_for_new_horizontal_machine + " Program Needs Attention."
#             card_label = "#WARNNING"
#             card_member = "@moemenalatweh1 " + "@" + trello_user_name
#         else:
#             card_title = new_program_number_for_new_horizontal_machine
#             card_label = "#SUCCESS"
#             card_member = "@moemenalatweh1 " + "@" + trello_user_name
#
#         #         # message to be sent to personal email
#         #         email_message = f"""From: Alatweh Moemen <malatweh@rwbteam.com>
#         # To: <moemenatweh@hotmail.com>
#         # Subject: Testing Email by python
#         #
#         #
#         # {"".join(email_messages_of_creating_old_horizontal_machine_program)}"""
#         #
#         #         email_Server.sendmail(sender_email, recipient_email, email_message)
#
#         # message to be sent trello email
#         email_message = f"""From: Alatweh Moemen <malatweh@rwbteam.com>
# Subject: {card_title}  {card_label}  {card_member}
#
#
# {"".join(email_messages_of_creating_new_horizontal_machine_program)}"""
#
#         # TO SEND THE EMAIL
#         email_server.sendmail(sender_email, trello_board_email, email_message)
#
#         # terminating the session
#         email_server.quit()
#
#     except Exception as error:
#         fail_messages_of_creating_new_horizontal_machine_program.append(
#             "Failed to send Email to Trello board." + "\n" + "An Error has occurred :" + "\n" + '[color=ff1a1a]' +
#                 str(error) + '[/color]' + "\n" + "Double Check Network and Login Authentication.")
#         # email_messages_of_creating_old_horizontal_machine_program.append(
#         "\n".join(fail_messages_of_creating_new_horizontal_machine_program))
#         failed_to_create_new_horizontal_machine_program(self)


# endregion <<<<=========================[New Horizontal Machines(127) Functions]===========================>>>>


# region <<<<================================[New Horizontal Machine(127) Screen]===============================>>>>
class NewHorizontalScreen(Screen):
    def create_program_for_new_horizontal_machine(self):
        print("create_program_for_new_horizontal_machine" + " FUNCTION is called")

        # global success_messages_of_creating_new_horizontal_machine_program
        # success_messages_of_creating_new_horizontal_machine_program = []
        #
        # global verification_messages_of_creating_new_horizontal_machine_program
        # verification_messages_of_creating_new_horizontal_machine_program = []
        #
        # global fail_messages_of_creating_new_horizontal_machine_program
        # fail_messages_of_creating_new_horizontal_machine_program = []
        #
        # global warning_messages_of_creating_new_horizontal_machine_program
        # warning_messages_of_creating_new_horizontal_machine_program = []
        #
        # global email_messages_of_creating_new_horizontal_machine_program
        # email_messages_of_creating_new_horizontal_machine_program = []
        # email_messages_of_creating_new_horizontal_machine_program.append(self.ids["JobNumber"].text + " Program on " +
        #                                                                  "Old Horizontal Machine" + "\n" +
        #                                                                  "Created by : " + connected_user_name + "\n")
        #
        # global new_program_number_for_new_horizontal_machine
        # new_program_number_for_new_horizontal_machine = self.ids["JobNumber"].text
        # print(new_program_number_for_new_horizontal_machine)
        #
        # if (self.ids["JobNumber"].text == ""):
        #     fail_messages_of_creating_new_horizontal_machine_program.append("Please Enter Job Number.")
        #     failed_to_create_new_horizontal_machine_program(self)
        #     return
        #
        # # NEED LOGIC TO CHECK IF NUMBER THAT ENTERED MATCH WITH SPEC DATABASE
        #
        # try:
        #     # print("try LOAD_HORIZONTAL_SHEETS_FOR_AUTOMATION")
        #     # CALL (LOAD_HORIZONTAL_SHEETS_FOR_AUTOMATION) FUNCTION TO LOAD HORIZONTAL TOOL LIST SHEETS
        #     load_horizontal_machine_tool_list_sheets(self)
        # except Exception as error:
        #     # print("except LOAD_HORIZONTAL_SHEETS_FOR_AUTOMATION")
        #     fail_messages_of_creating_new_horizontal_machine_program.append(
        #         "Failed to Find, Load, or Access" + '[b][u][color=ffffff] Horizontal Tool List [/color][/u][/b]' +
        #         "File to Create the Program." + "\n" + "An Error has occurred " + "\n" + '[color=ff1a1a]' + str(
        #             error) + '[/color]' + "\n" + "Double Check Network, and File Location.")
        #     email_messages_of_creating_new_horizontal_machine_program.append(
        #         "Failed to Find, Load, or Access Horizontal Tool List File to Create the Program." + "\n" +
        #         "Double Check Network, and File Location.")
        #     failed_to_create_new_horizontal_machine_program(self)
        #     return

# endregion <<<<==============================[New Horizontal Machine(127) Screen]=============================>>>>


# region <<<<==========================================[Setting Screen]==========================================>>>>

class SettingScreen(Screen):
    # ==========================================================================================|
    #  Create Function to check Admin info to NOT allowed any one else changing the App Setting.|
    # ==========================================================================================|
    def admin_check(self):
        admin = "malatweh@rwbteam.com"
        if (self.manager.get_screen('LoginScreen').ids["Email"].text == admin):
            self.manager.current = 'AppSettingScreen'
        else:
            close_button = MDRaisedButton(text='Close', on_release=self.close_setting_screen_window, font_size=16)
            self.setting_screen_message_window = MDDialog(title='[color=990000]Warning Message[/color]', text=(
                                '[color=ffffff]Sorry, You are NOT Authorized to access this screen.[/color]'),
                                           size_hint=(0.7, 1.0), buttons=[close_button], auto_dismiss=False)
            self.setting_screen_message_window.open()

    # ====================================================================================|
    #  Create Function to close screen message window when the User clicks on Close Button|
    # ====================================================================================|
    def close_setting_screen_window(self, obj):
        self.setting_screen_message_window.dismiss()


# endregion <<<<========================================[Setting Screen]=========================================>>>>


# region <<<<=====================================[Pin Bore Setting Screen]======================================>>>>
class PinBoreSettingScreen(Screen):
    pass

# endregion <<<<===================================[Pin Bore Setting Screen]=====================================>>>>


# region <<<<==========================[Old Horizontal Machines(28,29,32) Setting Screen]========================>>>>
class OldHorizontalSettingScreen(Screen):
    pass

# endregion <<<<=======================[Old Horizontal Machines(28,29,32) Setting Screen]========================>>>>


# region <<<<====================================[User Setting Screen]====================================>>>>
class UserSettingScreen(Screen):
    pass

# endregion <<<<===================================[User Setting Screen]==================================>>>>


# region <<<<====================================[Application Setting Screen]====================================>>>>
class AppSettingScreen(Screen):
    pass

# endregion <<<<===================================[Application Setting Screen]==================================>>>>


# region <<<<======================================[Add New User Screen]=======================================>>>>

class AddNewUserScreen(Screen):
    # ==========================================================|
    # [#] NOT Ready Yet, Needs Update.                          |
    # [#] Needs to work here if this feature Still will be used.|
    # ==========================================================|
    def add_new_user(self):
        try:
            print("New User Added Successfully, Make Sure You Send Him The Shared Password By Email")
            print("USER NAME:", self.ids["UserName"].text)
            print("RWB EMAIL:", self.ids["NewRWBEmail"].text)
            print("WISECO EMAIL:", self.ids["NewWisecoEmail"].text)
            # ADD NEW USER INFORMATION FOR EXCEL FILE(USER_NAME,WISECO_EMAIL, AND RWB_EMAIL)
            new_user_name = email_address_list_file['Email'].loc[self.ids["UserName"].text, 'Users'] = self.ids[
                "UserName"].text
            new_wiseco_email = email_address_list_file['Email'].loc[
                self.ids["NewWisecoEmail"].text, 'Wiseco Email Address'] = self.ids["NewWisecoEmail"].text
            new_rwb_email = email_address_list_file['Email'].loc[self.ids["NewRWBEmail"].text, 'RWB Email Address'] = \
                self.ids["NewRWBEmail"].text
            # CREATE DATA FRAME WITH THE NEW USER DATA TO ADD THEM TO THE EXCEL SHEET
            new_user_data = pd.DataFrame(
                data={'Users': [new_user_name], 'Wiseco Email Address': [new_wiseco_email],
                      'RWB Email Address': [new_rwb_email]})
            # LOAD THE EXCEL SHEET TO BE ABLE TO ADD NEW DATA
            user_information_workbook = openpyxl.load_workbook(email_address_list_file_path)
            # ACCESS THE EXCEL SHEET AND USE (mode= 'a') TO ADD THE NEW DATA
            user_information_sheet_update = pd.ExcelWriter(email_address_list_file_path, engine='openpyxl', mode='a')
            # SET (EMAIL_SHEET_UPDATE) AS CURRENT EXCEL BOOK (EXCEL FILE IN ANOTHER WORD)
            user_information_sheet_update.book = user_information_workbook
            # LOOP THROUGH THE EXCEL FILE TO SCAN ALL THE SHEETS
            # (MUST PUT THIS LINE OF CODE WHEN WE USE THIS MODE (mode= 'a') )
            user_information_sheet_update.sheets = dict((ws.title, ws) for ws in user_information_workbook.worksheets)
            # print("EMAIL_SHEET_UPDATE.sheets:", EMAIL_SHEET_UPDATE.sheets)
            # UPDATE THE EMAIL SHEET WITH ADDING THE NEW USER DATA
            new_user_data.to_excel(user_information_sheet_update, sheet_name='Email',
                                   startrow=user_information_workbook['Email'].max_row,
                                   startcol=0, header=False, index=False)
            # SAVE CHANGES
            user_information_sheet_update.save()
            # CLOSE SHEET
            user_information_sheet_update.close()

            # SHOW MESSAGE OF New User Added Successfully
            close_button = MDRaisedButton(text='Close', on_release=self.close_new_user_screen_window, font_size=16)
            self.new_user_screen_message_window = MDDialog(title='', text=(
                "New User Added Successfully, Make Sure You Send to Him The Shared Password By Email"),
                                           size_hint=(0.7, 1.0), buttons=[close_button], auto_dismiss=False)
            # TO OPEN THE DIALOG WINDOW
            self.new_user_screen_message_window.open()

        # ADD USER WORK FINE ONLY FOR ADDING ONE USER ,
        # THE PROBLEM IT WILL CORRUPT THE EXCEL WHICH MAKE IT INACCESSIBLE FILE, WE NEED TO FIGURE OUT THAT LATER***
        except (PermissionError):
            # SHOW MESSAGE OF New User Added Successfully
            close_button = MDRaisedButton(text='Close', on_release=self.close_new_user_screen_window, font_size=16)
            self.new_user_screen_message_window = MDDialog(title='', text=(
                "PermissionError: Something Get Wrong to Access the File, Email malatweh@rwbteam.com "),
                                           size_hint=(0.7, 1.0), buttons=[close_button], auto_dismiss=False)
            # TO OPEN THE DIALOG WINDOW
            self.new_user_screen_message_window.open()

    def close_new_user_screen_window(self, obj):
        self.new_user_screen_message_window.dismiss()
        # TO RESET ADD USER FIELDS TO START OVER
        self.ids["UserName"].text = ""
        self.ids["NewRWBEmail"].text = ""
        self.ids["NewWisecoEmail"].text = ""

    def reset_new_user_screen_fields(self):
        # TO RESET ADD USER FIELDS TO START OVER
        self.ids["UserName"].text = ""
        self.ids["NewRWBEmail"].text = ""
        self.ids["NewWisecoEmail"].text = ""


# endregion <<<<====================================[Add New User Screen]=====================================>>>>


# region <<<<==========================================[Screen Manager]==========================================>>>>

# ==============================|
# [#] Create the screen manager.|
# ==============================|
sm = ScreenManager()
sm.add_widget(LoginScreen(name='LoginScreen'))
sm.add_widget(HomeScreen(name='HomeScreen'))
sm.add_widget(PinBoreScreen(name='PinBoreScreen'))
sm.add_widget(OldHorizontalScreen(name='OldHorizontalScreen'))
sm.add_widget(NewHorizontalScreen(name='NewHorizontalScreen'))
sm.add_widget(SettingScreen(name='SettingScreen'))
sm.add_widget(PinBoreSettingScreen(name='PinBoreSettingScreen'))
sm.add_widget(OldHorizontalSettingScreen(name='OldHorizontalSettingScreen'))
sm.add_widget(UserSettingScreen(name='UserSettingScreen'))
sm.add_widget(AppSettingScreen(name='AppSettingScreen'))
sm.add_widget(AddNewUserScreen(name='AddNewUserScreen'))

# endregion <<<<=========================================[Screen Manager]========================================>>>>


# region <<<<=======================================[Application Builder]========================================>>>>

# ================================================|
# [#] Create Class with App name to Build the App.|
# ================================================|
class WisecoProgramsMaker(MDApp):
    # ====================================================|
    # [#] Use (build) method as function to build the app.|
    # ====================================================|
    def build(self):
        # ==========================================================================|
        # [#] To control the size of the App Screen (Window.size = (Width, Height)).|
        # ==========================================================================|
        Window.size = (900, 650)
        # =======================================================================|
        # [#] To choose the background mode of the App whether 'Dark' or 'Light'.|
        # =======================================================================|
        self.theme_cls.theme_style = "Dark"
        # =========================================================================|
        # [#] To set the default color of the App Elements (Labels, Buttons...Etc).|
        # =========================================================================|
        self.theme_cls.primary_palette = "Red"
        # ===========================================================================|
        # [#] To set the default color concentration (darkness and brightness) of the|
        #     App Elements (Labels, Bttons...etc).                                   |
        # ===========================================================================|
        self.theme_cls.primary_hue = "900"
        # =============================================================|
        # [#] Load (Screens_Builder) that contains all the App Screens.|
        # =============================================================|
        builder_screen = Builder.load_string(Screens_Builder)
        # ===========================================================|
        # [#] Define Variable to have All Screens and their Elements.|
        # ===========================================================|
        app_screen = Screen()
        # ===================================================|
        # [#] Create BoxLayout that's contain the Entire APP.|
        # ===================================================|
        app_box_layout = MDBoxLayout(
            orientation='vertical', spacing=20, padding=15,
            md_bg_color=[32 / 255.0, 32 / 255.0, 32 / 255.0, 1])
        # =====================================================================|
        # [#] Create Variable to Set the AppPicture from Online Website Source.|
        # =====================================================================|
        app_image = AsyncImage(
            source=r'https://images.squarespace-cdn.com/content/v1/5cf6a7664ba6460001928b8b/'
                   r'1559864161158-H9K9FU00BDGENWMNCLX7/Wiseco_Black.gif', size_hint_y=None,
            height=70, allow_stretch=True, pos_hint={'center_x': 0.5, 'center_y': 0.10},
            color=[150 / 255.0, 0 / 255.0, 0 / 255.0, 1])
        # ===============================================|
        # [#] Add AppPicture to the BoxLayout of the App.|
        # ===============================================|
        app_box_layout.add_widget(app_image)
        # ====================================================|
        # [#] Add Screens_Builder to the BoxLayout of the App.|
        # ====================================================|
        app_box_layout.add_widget(builder_screen)
        # ==============================================|
        # [#] Add BoxLayout of the App to the AppScreen.|
        # [#] Return the AppScreen to Display the App.  |
        # ==============================================|
        app_screen.add_widget(app_box_layout)
        return app_screen


# ==========================|
# [#] Run and Start the App.|
# ==========================|
WisecoProgramsMaker().run()

# endregion <<<<=====================================[Application Builder]=======================================>>>>
