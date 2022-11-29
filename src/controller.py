from kivymd.app import MDApp
from kivy.lang.builder import Builder
from kivy.uix.screenmanager import ScreenManager, Screen

import model

class Main_Screen(Screen):
    pass

class EA_Screen(Screen):
    pass

sm = ScreenManager()
sm.add_widget(Main_Screen(name="main"))
sm.add_widget(Main_Screen(name="ea"))

class DemoApp(MDApp):
    def build(self):
        screen = Builder.load_file("view.kv")
        return screen

    def click(self):
        model.early_engagement()
        
DemoApp().run()