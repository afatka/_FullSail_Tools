import sys
import random
from PySide2 import QtCore, QtWidgets, QtGui

class MyWidget(QtWidgets.QWidget):
	def __init__(self):
		QtWidgets.QWidget.__init__(self)

		self.hello = ["Hallo Welt", "你好，世界", "Hei maailma",
			"Hola Mundo", "Привет мир"]

		self.button = QtWidgets.QPushButton("Click Me!")
		self.text = QtWidgets.QLabel("Hello World!")
		self.text.setAlignment(QtCore.Qt.AlignCenter)

		self.layout = QtWidgets.QVBoxLayout()
		self.layout.addWidget(self.text)
		self.layout.addWidget(self.button)
		self.setLayout(self.layout)

		self.button.clicked.connect(self.magic)


	def magic(self):
		self.text.setText(random.choice(self.hello))
		self.text.repaint()

if __name__ == "__main__":
	app = QtWidgets.QApplication(sys.argv)

	widget = MyWidget()
	widget.show()

	sys.exit(app.exec_())