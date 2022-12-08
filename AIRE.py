from kivymd.app import MDApp
from kivy.lang.builder import Builder
from kivy.uix.screenmanager import ScreenManager, Screen

from modules import earlyengagement, preagp0, agp0

from tkinter import filedialog as fd

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

class DemoApp(MDApp):

    operational_plan_path = ""
    sheet_name = ""
    assessment_file_path = ""
    agp0_file_path = ""

    def build(self):
        self.operational_plan_path = ""
        self.sheet_name = "RUN"
        self.assessment_file_path = ""
        agp0_file_path = ""

        screen = Builder.load_file("view.kv")
        return screen

    def get_operational_plan_path(self):
        filetypes = (
            ('Excel files', '*.xlsx'),
        )

        self.operational_plan_path = fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)

        print(self.operational_plan_path)

    def get_assessment_file_path(self):
        filetypes = (
            ('Excel files', '*.xlsx'),
        )

        self.assessment_file_path = fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)

    def get_agp0_file_path(self):
        filetypes = (
            ('All files', '*.*'),
        )

        self.agp0_file_path = fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)

    def set_sheet_name_run(self):
        self.sheet_name = "RUN"
        print(self.sheet_name)

    def set_sheet_name_grow(self):
        self.sheet_name = "GROW"
        print(self.sheet_name)

    def set_sheet_name_transform(self):
        self.sheet_name = "TRANSFORM"
        print(self.sheet_name)

    def run_early_engagement_module(self):
        earlyengagement.run(self.operational_plan_path, self.sheet_name)

    def run_pre_agp0_module(self):
        preagp0.run("sample", self.assessment_file_path)
        
    def run_agp0_module(self):
        agp0.run(self.agp0_file_path)

DemoApp().run()