# import sys
# from PyQt5.QtWidgets import QApplication, QWidget, QPushButton
# from PyQt5.QtGui import QIcon, QCursor
# from PyQt5.QtCore import Qt, QThreadPool, QRunnable
# from stats import run_daily_stats

# class App(QWidget):

#     def __init__(self):
#         super().__init__()
#         self.title = 'Stats Orginizer'
#         self.left = 100
#         self.top = 100
#         self.width = 320
#         self.height = 200
#         self.initUI()

#     def initUI(self):
#         self.setWindowTitle(self.title)
#         self.setGeometry(self.left, self.top, self.width, self.height)

#         # Create a button and style it
#         button = QPushButton('Create Stats', self)
#         button.setToolTip('Generate daily statistics')
#         button.setGeometry(100, 70, 120, 50)
#         button.setStyleSheet('''
#             QPushButton {
#                 background-color: #4CAF50;
#                 color: white;
#                 border-radius: 10px;
#                 font-size: 16px;
#             }
#             QPushButton:hover {
#                 background-color: #66BB6A;
#             }
#             QPushButton:pressed {
#                 background-color: #2E7D32;
#             }
#         ''')

#         button.setCursor(Qt.PointingHandCursor)

#         # Connect the button to the function that generates the statistics
#         button.clicked.connect(self.on_click)

#         self.show()

#     def on_click(self):
#         # Set the cursor to WaitCursor
#         QApplication.setOverrideCursor(Qt.WaitCursor)

#         # Create a runnable object and run it in a separate thread
#         stats_runnable = StatsRunnable()
#         QThreadPool.globalInstance().start(stats_runnable)

#         # Set the cursor back to ArrowCursor
#         QApplication.restoreOverrideCursor()

# class StatsRunnable(QRunnable):
#     def run(self):
#         run_daily_stats()
#         print('Stats created')

# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     ex = App()
#     sys.exit(app.exec_())


# import sys
# from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QTextEdit
# from PyQt5.QtGui import QIcon, QCursor
# from PyQt5.QtCore import Qt, QThreadPool, QRunnable
# from stats import run_daily_stats

# class App(QWidget):

#     def __init__(self):
#         super().__init__()
#         self.title = 'Stats Orginizer'
#         self.left = 100
#         self.top = 100
#         self.width = 320
#         self.height = 300
#         self.initUI()

#     def initUI(self):
#         self.setWindowTitle(self.title)
#         self.setGeometry(self.left, self.top, self.width, self.height)

#         # Create a layout
#         layout = QVBoxLayout()

#         # Create a button and style it
#         button = QPushButton('Create Stats', self)
#         button.setToolTip('Generate daily statistics')
#         button.setStyleSheet('''
#             QPushButton {
#                 background-color: #4CAF50;
#                 color: white;
#                 border-radius: 10px;
#                 font-size: 16px;
#             }
#             QPushButton:hover {
#                 background-color: #66BB6A;
#             }
#             QPushButton:pressed {
#                 background-color: #2E7D32;
#             }
#         ''')

#         # Set the cursor to a pointer when hovering the button
#         button.setCursor(Qt.PointingHandCursor)

#         # Connect the button to the function that generates the statistics
#         button.clicked.connect(self.on_click)

#         # Create a QTextEdit for output
#         self.output_text = QTextEdit()
#         self.output_text.setReadOnly(True)

#         # Add button and QTextEdit to the layout
#         layout.addWidget(button)
#         layout.addWidget(self.output_text)

#         self.setLayout(layout)
#         self.show()

#     def on_click(self):
#         # Set the cursor to WaitCursor
#         QApplication.setOverrideCursor(Qt.WaitCursor)

#         # Create a runnable object and run it in a separate thread
#         stats_runnable = StatsRunnable(self.append_output)
#         QThreadPool.globalInstance().start(stats_runnable)

#         # Set the cursor back to ArrowCursor
#         QApplication.restoreOverrideCursor()

#     def append_output(self, text):
#         self.output_text.append(text)

# class StatsRunnable(QRunnable):
#     def __init__(self, append_output_func):
#         super().__init__()
#         self.append_output_func = append_output_func

#     def run(self):
#         run_daily_stats()
#         self.append_output_func('Stats created')

# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     ex = App()
#     sys.exit(app.exec_())


import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QTextEdit
from PyQt5.QtGui import QIcon, QCursor
from PyQt5.QtCore import Qt, QThreadPool, QRunnable,  QDateTime
from stats import run_daily_stats
import resources


class App(QWidget):

    def __init__(self):
        super().__init__()
        self.title = 'Stats Orginizer'
        self.left = 100
        self.top = 100
        self.width = 320
        self.height = 300
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        # Set the window icon
        self.setWindowIcon(QIcon(':/icon.png'))

        # Create a layout
        layout = QVBoxLayout()

        # Create a button and style it
        button = QPushButton('Generate Stats', self)
        button.setToolTip('Generate statistics')
        button.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')

        # Set the cursor to a pointer when hovering the button
        button.setCursor(Qt.PointingHandCursor)

        # Connect the button to the function that generates the statistics
        button.clicked.connect(self.on_click)

        # Create a QTextEdit for output
        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)

        # Add the current date and time to the output text
        current_datetime = QDateTime.currentDateTime()
        self.output_text.append(
            f"Date & Time: {current_datetime.toString('yyyy-MM-dd hh:mm:ss')}")
        self.output_text.append(
        "The newly generated stats downloaded from ICBM must be saved "
        "in the same folder as the applications. The filename must always "
        "be kept as 'User Productivity Summary' and should not be changed."
        )

        # Add button and QTextEdit to the layout
        layout.addWidget(button)
        layout.addWidget(self.output_text)

        self.setLayout(layout)
        self.show()

    def on_click(self):
        # Set the cursor to WaitCursor
        QApplication.setOverrideCursor(Qt.WaitCursor)

        # Create a runnable object and run it in a separate thread
        stats_runnable = StatsRunnable(self.append_output)
        QThreadPool.globalInstance().start(stats_runnable)

        # Set the cursor back to ArrowCursor
        QApplication.restoreOverrideCursor()

    def append_output(self, text):
        self.output_text.append(text)


class StatsRunnable(QRunnable):
    def __init__(self, append_output_func):
        super().__init__()
        self.append_output_func = append_output_func

    def run(self):
        errors = run_daily_stats()
        if (errors):
            self.append_output_func(str(errors))
        else:
            self.append_output_func('Stats created')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())

