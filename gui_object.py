import openpyxl
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *

def MakeAction(name:str, parent:QObject, shortcut:str = "", tip:str = ""):
    action = QAction(name, parent)
    action.setShortcut(shortcut)
    action.setStatusTip(tip)
    return action

def AddMenu(menu_bar:QMenuBar, menu_name:str, action:QAction):
    menu = menu_bar.addMenu(menu_name)
    menu.addAction(action)
    return menu

class ProgressBar(QProgressBar):
    def __init__(self):
        super().__init__()
        self.setRange(0, 1)
        self.setAlignment(Qt.AlignCenter)
        self.isWorking = False

    def Processing(self):
        self.repaint()

    def Start(self):
        self.setRange(0, 0)
        self.isWorking = True

    def End(self):
        self.setRange(0, 1)
        self.isWorking = False
        return True