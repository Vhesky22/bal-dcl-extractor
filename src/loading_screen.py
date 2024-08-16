from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QProgressBar
from PySide6.QtGui import QPixmap
from PySide6.QtCore import Qt, QRect
import os

class LoadingScreen(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Loading...")
        self.setGeometry(100, 100, 765, 505)

        # Create a layout for the main window
        self.layout = QVBoxLayout(self)

        # Create a QLabel for the background image
        self.background_label = QLabel(self)
        self.background_label.setGeometry(QRect(0, 0, self.width(), self.height()))

        # Calculate the path to the image
        base_path = os.path.dirname(__file__)  # Directory of this file
        image_path = os.path.join(base_path, '..', 'resources', 'wallpapers', 'loading_screen.jpg')
        image_path = os.path.abspath(image_path)  # Convert to absolute path

        # Load the pixmap and set the background image
        self.background_pixmap = QPixmap(image_path)
        self.update_background_image()

        # Create the widgets to be stacked on top of the background
        self.label = QLabel("Loading, please wait...")
        self.layout.addWidget(self.label)

        self.progress_bar = QProgressBar()
        self.layout.addWidget(self.progress_bar)

        # Set the layout for the main window
        self.setLayout(self.layout)

        # Initialize progress bar value
        self.progress_bar.setValue(0)

    def update_background_image(self):
        # Center the image and ensure it fits within the window
        pixmap_scaled = self.background_pixmap.scaled(self.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        self.background_label.setPixmap(pixmap_scaled)
        self.background_label.setGeometry(QRect(0, 0, self.width(), self.height()))

    def resizeEvent(self, event):
        # Update the background image size and position when the window is resized
        self.update_background_image()
        super().resizeEvent(event)

    def set_progress(self, value):
        self.progress_bar.setValue(value)
        if value == 100:
            # Placeholder for additional functionality
            pass
