from kivymd.app import MDApp
from kivy.lang.builder import Builder
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.clock import Clock
from kivy.core.window import Window

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

import AIREV1
import AIREV2

import time

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


# Window.size = (600,300)

sm = ScreenManager()
sm.add_widget(Menu_Screen(name='menu'))
sm.add_widget(Early_Engagement_Screen(name='profile'))
sm.add_widget(Pre_AGP0_Screen(name='upload'))
sm.add_widget(AGP0_Screen(name='upload'))
sm.add_widget(Loading_Screen(name='loading'))

class DemoApp(MDApp):

    directory = ""
    sheet_name = ""
    assessment_file_directory = ""

    def build(self):
        self.directory = ""
        self.sheet_name = ""
        self.assessment_file_directory = ""
        screen = Builder.load_file('kivy_sample.kv')
        return screen

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
        AIREV1.generate_templates(self.sheet_name)

    def pre_agp0_generate(self):
        current_screen = DemoApp.get_running_app().root.get_screen('pre_agp0')
        initiative_name = current_screen.ids.initiative_name.text 

        AIREV2.generate_pre_agp_0_report(initiative_name, self.assessment_file_directory)

DemoApp().run()