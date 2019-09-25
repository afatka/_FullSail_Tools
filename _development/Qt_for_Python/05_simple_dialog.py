import sys
from PySide2.QtWidgets import (QApplication, QDialog, QLineEdit, 
	QPushButton, QVBoxLayout)

class Form(QDialog):

	def __init__(self, parent = None):
		super(Form, self).__init__(parent)
		self.setWindowTitle("My Form")

		# Create Widgets
		self.edit = QLineEdit("Write my name here...")
		self.button = QPushButton("Show Greeting")

		# Create the layout and add the widgets
		layout = QVBoxLayout()
		layout.addWidget(self.edit)
		layout.addWidget(self.button)
		# Set dialog layout
		self.setLayout(layout)

		# Add button signal to greeting slot
		self.button.clicked.connect(self.greetings)

	def greetings(self):
		print("Hello {}".format(self.edit.text()))


if __name__ == "__main__":
	# Create Qt Application
	app = QApplication(sys.argv)
	# Create and show the form
	form = Form()
	form.show()
	# run the main Qt Loop
	sys.exit(app.exec_())