import sys
from Functions import *
from PyQt5.QtWidgets import (QWidget, QPushButton, QVBoxLayout,
                             QApplication,
                             QHBoxLayout, QLabel)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setGeometry(0, 0, 400, 200)
        self.move(400, 600)

        self.title = QLabel('Before starting please close all outlook windows, except main one')
        self.btn = QPushButton('Continue')

        self.initui()

    def initui(self):
        hbox = QHBoxLayout()
        hbox.addStretch()
        hbox.addWidget(self.title)
        hbox.addStretch()

        hbox2 = QHBoxLayout()
        hbox2.addStretch()
        hbox2.addWidget(self.btn)
        hbox2.addStretch()

        vbox = QVBoxLayout()
        vbox.addStretch()
        vbox.addLayout(hbox)
        vbox.addLayout(hbox2)
        vbox.addStretch()


        self.setLayout(vbox)
        self.show()


class SecondWindow():
    pass