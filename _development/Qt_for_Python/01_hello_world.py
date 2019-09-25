import sys
from PySide2.QtWidgets import QApplication, QLabel

# https://wiki.qt.io/Qt_for_Python_Tutorial_HelloWorld

app = QApplication(sys.argv)
#label = QLabel("Hello World")
label = QLabel("<font color=red size=40>Hello World!</font>")
label.show()
app.exec_()