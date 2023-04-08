import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton
from PyQt5.QtGui import QIcon

from stats import run_daily_stats

class App(QWidget):

    def __init__(self):
        super().__init__()
        self.title = 'Stats Organizer'
        self.left = 100
        self.top = 100
        self.width = 320
        self.height = 200
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        # Create a button and style it
        button = QPushButton('Create Stats', self)
        button.setToolTip('Generate daily statistics')
        button.setGeometry(100, 70, 120, 50)
        button.setStyleSheet('background-color: #4CAF50; color: white; border-radius: 10px; font-size: 16px;')

        # Add an icon to the button
        icon = QIcon('icon.png')
        button.setIcon(icon)
        button.setIconSize(button.sizeHint())

        # Connect the button to the function that generates the statistics
        button.clicked.connect(self.on_click)

        self.show()

    def on_click(self):
        run_daily_stats()
        print('Stats created')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
