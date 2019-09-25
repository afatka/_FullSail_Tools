import sys
from PySide2.QtWidgets import QApplication, QPushButton

# Greetings
def say_hello():
	print("Button clicked, Hello!")

# Create the Qt Application
app = QApplication(sys.argv)

# Create a button
# button = QPushButton("click me!")
# # Connect the button to the function
# button.clicked.connect(say_hello)
# # Show the button
# button.show()

button_exit = QPushButton("exit")
button_exit.clicked.connect(app.exit)
button_exit.show()


# Run the main Qt Loop
app.exec_()