from Functions import *
from PyQt5.QtWidgets import (QWidget, QPushButton, QVBoxLayout,
                             QApplication,
                             QHBoxLayout, QLabel)
from PyQt5.QtCore import QThreadPool, QRunnable


class Worker(QRunnable):
    def __init__(self):
        super().__init__()

        self.processStart = EmailBuilder(xl_rows)

    def run(self):
        pythoncom.CoInitialize()
        self.processStart.create()

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Email Generator')
        self.setGeometry(0, 0, 400, 200)
        self.move(400, 600)

        self.title = QLabel('Before starting please close all outlook windows, except main one')
        self.btn = QPushButton('Continue')

        self.hbox = QHBoxLayout()
        self.hbox2 = QHBoxLayout()
        self.vbox = QVBoxLayout()
        self.secondWindow = SecondWindow()
        self.dataPrep = DataPreparation(xl_rows, data_list)

        self.initui()

    def initui(self):
        self.hbox.addStretch()
        self.hbox.addWidget(self.title)
        self.hbox.addStretch()

        self.hbox2.addStretch()
        self.hbox2.addWidget(self.btn)
        self.hbox2.addStretch()

        self.vbox.addStretch()
        self.vbox.addLayout(self.hbox)
        self.vbox.addLayout(self.hbox2)
        self.vbox.addStretch()

        self.btn.clicked.connect(self.preparation)
        self.btn.clicked.connect(self.windowclose)

        self.setLayout(self.vbox)
        self.show()

    def preparation(self):
        self.dataPrep.datamixer()
        self.secondWindow.initui()

    def windowclose(self):
        self.close()


class SecondWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.setWindowTitle('Email Generator')
        self.setGeometry(0, 0, 400, 200)
        self.move(400, 600)

        self.label = QLabel(start)
        self.btn1 = QPushButton('Yes')
        self.btn2 = QPushButton('No')
        self.hbox = QHBoxLayout()
        self.hbox2 = QHBoxLayout()
        self.vbox = QVBoxLayout()


    def initui(self):
        self.hbox.addStretch()
        self.hbox.addWidget(self.label)
        self.hbox.addStretch()

        self.hbox2.addStretch()
        self.hbox2.addWidget(self.btn1)
        self.hbox2.addWidget(self.btn2)
        self.hbox2.addStretch()

        self.vbox.addStretch()
        self.vbox.addLayout(self.hbox)
        self.vbox.addLayout(self.hbox2)
        self.vbox.addStretch()
        self.btn1.clicked.connect(self.createemails)
        self.btn2.clicked.connect(self.windowclose)
        self.process = Worker()
        self.threadpool = QThreadPool()

        self.setLayout(self.vbox)
        self.show()

    def windowclose(self):
        self.close()

    def createemails(self):
        self.threadpool.start(self.process)

class ThirdWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Email Generator')
        self.setGeometry(0, 0, 400, 200)
        self.move(400, 600)

        self.btn = QPushButton('Stop Generating Emails')
        self.vbox = QVBoxLayout()

    def initui(self):
        self.vbox.addStretch()
        self.vbox.addWidget(self.btn)
        self.vbox.addStretch()

        self.setLayout(self.vbox)
        self.show()

    def closeapp(self):
        pass



