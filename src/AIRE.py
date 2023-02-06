from kivymd.app import MDApp
from kivy.lang.builder import Builder
from kivy.uix.screenmanager import ScreenManager, Screen
from kivymd.uix.dialog import MDDialog

from tkinter import filedialog as fd, Tk

from early_engagement.early_engagement import early_engagement 
from pre_agp0 import pre_agp0
from agp0 import agp0

import os
from pathlib import Path

class Main_Screen(Screen):
    pass

class EA_Screen(Screen):
    pass

class Pre_AGP0_Screen(Screen):
    pass

class AGP0_Screen(Screen):
    pass

class Info_Screen(Screen):
    pass

sm = ScreenManager()
sm.add_widget(Main_Screen(name="main"))
sm.add_widget(Main_Screen(name="early_engagement"))
sm.add_widget(Pre_AGP0_Screen(name="pre_agp0"))
sm.add_widget(AGP0_Screen(name="agp0"))
sm.add_widget(Info_Screen(name="info"))

class AIREApp(MDApp):

    def build(self):
        self.path_to_current_op = ""
        self.path_to_risk_assessment = ""
        self.path_to_agp0_zip = ""

        Tk().withdraw()

        screen = Builder.load_file("view.kv")
        return screen

    def run_early_engagement_module(self):
        early_engagement(self.path_to_current_op).check_first_run().add_previous_op_to_output_folder().compare_current_previous_op().generate_comparison_report().generate_comparison_tables().generate_intake_forms().archive_files().clear_output_folder()

    def prompt_path_to_current_op(self):
        filetypes = (
            ('Excel files', '*.xlsx'),
        )

        self.path_to_current_op = fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)

    def prompt_path_to_risk_assessment(self):
        filetypes = (
            ('Excel files', '*.xlsx'),
        )

        self.path_to_risk_assessment = fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)

    def prompt_path_to_agp0_zip(self):
        filetypes = (
            ('All files', '*.*'),
        )

        self.path_to_agp0_zip = fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)

    def show_dialog(self):
        dialog = MDDialog(text="Intake form generation started. Wait about 30 seconds")
        dialog.open()
        
    def open_early_engagement_archive(self):
        os.startfile(str(Path().absolute() / 'data' / 'archive' / 'early_engagement'))

    def run_pre_agp0_module(self):
        pre_agp0.run("sample", self.path_to_risk_assessment)

    def open_pre_agp0_output(self):
        os.startfile(str(Path().absolute() / 'data' / 'output' / 'pre_agp0')) 

    def run_agp0_module(self):
        agp0.run(self.path_to_agp0_zip)

AIREApp().run()