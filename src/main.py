import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QPixmap, QImage
from loading_screen import LoadingScreen
from main_window import MainWindow

class MainApp(QApplication):
    def __init__(self, argv):
        super().__init__(argv)
        self.loading_screen = LoadingScreen()
        self.loading_screen.show()

        # Create a timer to simulate loading progress
        self.timer = QTimer()
        self.timer.timeout.connect(self.simulate_loading)
        self.timer.start(20)  # Simulate loading progress every 50 ms

        self.progress = 0

    def simulate_loading(self):
        self.progress += 1
        self.loading_screen.set_progress(self.progress)

        if self.progress >= 100:
            self.timer.stop()
            self.loading_screen.close()
            self.open_main_window()

    def open_main_window(self):
        self.main_window = MainWindow()
        self.main_window.show()

if __name__ == "__main__":
    app = MainApp(sys.argv)
    sys.exit(app.exec())
