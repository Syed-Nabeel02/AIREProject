from kivymd.app import MDApp
from kivy.lang.builder import Builder
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.clock import Clock
from kivy.core.window import Window

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

import earlyengagement
import preagp0
import agp0
import plancomparison

class Menu_Screen(Screen):
    pass

class Early_Engagement_Screen(Screen):
    pass

class Pre_AGP0_Screen(Screen):
    pass

class AGP0_Screen(Screen):
    pass

class Loading_Screen(Screen):
    pass

class Comparison_Screen(Screen):
    pass

class Recess_Screen(Screen):
    pass

class Info_Screen(Screen):
    pass

# Window.size = (600,300)

sm = ScreenManager()
sm.add_widget(Menu_Screen(name='menu'))
sm.add_widget(Early_Engagement_Screen(name='profile'))
sm.add_widget(Pre_AGP0_Screen(name='upload'))
sm.add_widget(AGP0_Screen(name='upload'))
sm.add_widget(Loading_Screen(name='loading'))
sm.add_widget(Comparison_Screen(name='comparison'))
sm.add_widget(Recess_Screen(name='recess'))
sm.add_widget(Info_Screen(name='info'))

class DemoApp(MDApp):

    # Model *************************
    directory = ""
    sheet_name = ""
    assessment_file_directory = ""
    agp0_supplements_directory = ""
    old_operational_plan_directory = ""
    new_operational_plan_directory = ""

    def build(self):
        self.directory = ""
        self.sheet_name = ""
        self.assessment_file_directory = ""
        self.agp0_supplements_directory = ""
        self.old_operational_plan_directory = ""
        self.new_operational_plan_directory = ""
        screen = Builder.load_file('AIRE.kv')
        return screen

    # Controller ************** 
    
    def select_file(self): # ***********************************
        filetypes = (
            ('Excel files', '*.xlsx'),
        )

        #    ('text files', '*.txt'),
        #    ('Excel files', '*.xlsx'),
        #    ('All files', '*.*')

        directory = fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)

        self.directory = directory
        
        current_screen = DemoApp.get_running_app().root.get_screen('early_engagement')
        current_screen.ids.directory.text =  directory

    def select_file_2(self):
        filetypes = (
            ('Excel files', '*.xlsx'),
        )

        directory = fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)

        self.assessment_file_directory = directory
        current_screen = DemoApp.get_running_app().root.get_screen('pre_agp0')
        current_screen.ids.assessment_file_directory.text =  directory

    def select_file_3(self):
        filetypes = (
            ('All files', '*.*'),
        )

        directory = fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)

        self.agp0_supplements_directory = directory
        current_screen = DemoApp.get_running_app().root.get_screen('agp0')
        current_screen.ids.agp0_supplements_directory.text =  directory

    def select_file_4(self):
        filetypes = (
            ('Excel files', '*.xlsx'),
        )

        directory = fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)

        self.old_operationa_plan_directory = directory
        current_screen = DemoApp.get_running_app().root.get_screen('comparison')
        current_screen.ids.old_operationa_plan_directory.text =  directory

    def select_file_5(self):
        filetypes = (
            ('Excel files', '*.xlsx'),
        )

        directory = fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)

        self.new_operationa_plan_directory = directory
        current_screen = DemoApp.get_running_app().root.get_screen('comparison')
        current_screen.ids.new_operationa_plan_directory.text =  directory
    
    def select_run(self):
        self.sheet_name = "RUN"
        current_screen = DemoApp.get_running_app().root.get_screen('early_engagement')
        current_screen.ids.sheet_name.text =  "Run"
    
    def select_grow(self):
        self.sheet_name = "GROW"
        current_screen = DemoApp.get_running_app().root.get_screen('early_engagement')
        current_screen.ids.sheet_name.text =  "Grow"

    def select_transform(self):
        self.sheet_name = "TRANSFORM"
        current_screen = DemoApp.get_running_app().root.get_screen('early_engagement')
        current_screen.ids.sheet_name.text = "Transform"

    def ea_generate(self):
        earlyengagement.generate_templates(self.directory, self.sheet_name)

    def pre_agp0_generate(self):
        current_screen = DemoApp.get_running_app().root.get_screen('pre_agp0')
        initiative_name = current_screen.ids.initiative_name.text 

        preagp0.generate_pre_agp_0_report(initiative_name, self.assessment_file_directory)

    def agp0_generate(self):
        agp0.generate_agp_0_report(self.agp0_supplements_directory)

    def compare_old_new_operational_plan(self):
        plancomparison.compare_opertional_plan_files(self.old_operationa_plan_directory, self.new_operationa_plan_directory)

DemoApp().run()