import sys
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QTextEdit,  # Import QTextEdit for text display
    QMessageBox,  # Import QMessageBox for popup notifications
    QWidget,
)
from PyQt5.QtGui import QPixmap, QFont  # QPainter, QPen, QColor
from PyQt5.QtCore import Qt  # QRect

# import mss
# import mss.tools
import pyautogui

# import pyperclip
from PIL import ImageGrab
import io
import win32clipboard
import time
import pytesseract
import pyperclip  # this library function is to copy the text to clipboard


class ScreenshotApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Screenshot Capture")
        self.setGeometry(100, 100, 800, 600)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)

        self.capture_button = QPushButton("Capture Region")
        self.capture_button.clicked.connect(self.start_capture)
        self.layout.addWidget(self.capture_button)

        self.image_label = QLabel()
        self.layout.addWidget(self.image_label)

        # Add a QTextEdit widget for displaying the extracted text
        self.text_box = QTextEdit()
        self.text_box.setPlaceholderText("Extracted text will appear here...")
        # Set a larger font size for the QTextEdit widget
        font = QFont()
        font.setPointSize(14)  # Set the font size to 14 (or any size you prefer)
        self.text_box.setFont(font)
        self.layout.addWidget(self.text_box)

        self.capturing = False
        self.start_point = None
        self.end_point = None
        self.capture_rect = None

    def start_capture(self):
        self.showMinimized()
        # self.capturing = True
        # self.setCursor(Qt.CrossCursor)
        pyautogui.hotkey("win", "shift", "s")  # open snipping tool native to Windows

        # Get the initial clipboard sequence number
        initial_sequence = win32clipboard.GetClipboardSequenceNumber()

        # Wait for the clipboard content to change
        while True:
            try:
                current_sequence = win32clipboard.GetClipboardSequenceNumber()

                # Check if the clipboard sequence number has changed
                if current_sequence != initial_sequence:
                    # Proceed to process the new image
                    print("New capture detected in clipboard.")
                    self.capture_clipboard_image()
                    break  # Exit the loop once the new capture is detected
            except Exception as e:
                print(f"Error checking clipboard: {e}")
                break
            time.sleep(0.1)  # Wait for a short time before checking again

    def capture_clipboard_image(self):
        try:
            # Open the clipboard
            win32clipboard.OpenClipboard()

            # Use ImageGrab from Pillow to get the image from the clipboard
            clipboard_image = ImageGrab.grabclipboard()

            if clipboard_image:
                # Convert the image to a format compatible with Tesseract
                clipboard_image = clipboard_image.convert("L")  # Convert to grayscale

                # Perform OCR on the image
                pytesseract.pytesseract.tesseract_cmd = (
                    r"C:\Program Files\Tesseract-OCR\tesseract.exe"
                )
                text = pytesseract.image_to_string(clipboard_image, lang="eng+vie")

                # Display the extracted text in the QTextEdit widget
                self.text_box.setPlainText(
                    text
                )  # Set the extracted text in the QTextEdit widget

                # Copy the extracted text to the clipboard
                pyperclip.copy(text)
                # print("Extracted text has been copied to the clipboard.")

                # Show a popup notification
                msg_box = QMessageBox()
                msg_box.setIcon(QMessageBox.Information)
                msg_box.setWindowTitle("Text Extraction Complete")
                msg_box.setText(
                    "The texts has been extracted and copied to the clipboard."
                )
                msg_box.setStandardButtons(QMessageBox.Ok)
                msg_box.exec_()  # Display the popup

                # Save the image to a byte array for display
                img_byte_arr = io.BytesIO()
                clipboard_image.save(img_byte_arr, format="PNG")
                img_byte_arr = img_byte_arr.getvalue()

                # Display the image in the QLabel
                pixmap = QPixmap()
                pixmap.loadFromData(img_byte_arr)

                # Check if the image width exceeds the QLabel width
                if pixmap.width() > self.image_label.width():
                    # Scale the image to fit within the QLabel width while maintaining aspect ratio
                    pixmap = pixmap.scaled(
                        self.image_label.width(),
                        pixmap.height(),
                        Qt.KeepAspectRatio,
                        Qt.SmoothTransformation,
                    )

                self.image_label.setPixmap(pixmap)
            else:
                print("No image found in clipboard.")
        except Exception as e:
            print(f"Error getting image from clipboard: {e}")

        # win32clipboard.CloseClipboard()
        self.showNormal()  # Restore the window.


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ScreenshotApp()
    window.show()
    sys.exit(app.exec_())
