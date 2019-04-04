import sys
from Functions import *
from PyQt5.QtWidgets import (QWidget, QPushButton, QVBoxLayout,
                             QApplication,
                             QHBoxLayout, QLabel)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Email Generator')
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

        self.btn.clicked.connect(self.preparation)
        self.btn.clicked.connect(self.windowclose)

        self.setLayout(vbox)
        self.show()

    def preparation(self):

        self.preparation = DataPreparation(xl_rows, data_list)
        self.preparation.datamixer()

        self.preparation = SecondWindow()
        self.preparation.initui()

    def windowclose(self):
        self.close()


class SecondWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.setWindowTitle('Email Generator')
        self.setGeometry(0, 0, 400, 200)
        self.move(400, 600)

    def initui(self):
        self.label = QLabel(start)
        self.btn1 = QPushButton('Yes')
        self.btn2 = QPushButton('No')

        hbox = QHBoxLayout()
        hbox.addStretch()
        hbox.addWidget(self.label)
        hbox.addStretch()

        hbox2 = QHBoxLayout()
        hbox2.addStretch()
        hbox2.addWidget(self.btn1)
        hbox2.addWidget(self.btn2)
        hbox2.addStretch()

        vbox = QVBoxLayout()

        vbox.addStretch()
        vbox.addLayout(hbox)
        vbox.addLayout(hbox2)
        vbox.addStretch()

        self.setLayout(vbox)
        self.show()

    def windowclose(self):
        self.close()


class ThirdWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Email Generator')
        self.setGeometry(0, 0, 400, 200)
        self.move(400, 600)

    def initui(self):
        self.btn = QPushButton()



        self.show()