import base64
import json
import sys
from PyQt5.QtWidgets import QDesktopWidget, QApplication, QMainWindow, QFileDialog, QLabel, QVBoxLayout, QWidget, QScrollArea, QMessageBox, QPushButton
from PyQt5.QtGui import QPixmap, QImage, QPainter
from PyQt5.QtCore import Qt
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
import fitz  # PyMuPDF for PDFs
from PIL import Image
import os

class CustomFileViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Custom File Viewer")
        self.setGeometry(100, 100, 800, 600)
        self.center_window()

        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.setCentralWidget(self.scroll_area)

        self.container = QWidget()
        self.layout = QVBoxLayout(self.container)
        self.layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)
        self.scroll_area.setWidget(self.container)

        # Add a print button
        self.print_button = QPushButton("Print", self)
        self.print_button.clicked.connect(self.print_document)
        self.layout.addWidget(self.print_button)

    def open_file(self):
        try:
            file_path, _ = QFileDialog.getOpenFileName(self, "Select a Custom File", "", "Custom Files (*.myext);;All Files (*)")
            if file_path:
                self.open_viewer_window(file_path)
                self.show()  # Show the main window only if a file is selected
            else:
                QMessageBox.information(self, "No File Selected", "No file was selected. Exiting the application.")
                QApplication.instance().quit()  # Ensure the application exits completely
                sys.exit()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {e}")
            QApplication.instance().quit()  # Ensure the application exits completely
            sys.exit()

    def open_viewer_window(self, custom_file_path):
        try:
            # Read and decode the .myext file
            with open(custom_file_path, "r") as custom_file:
                data = json.load(custom_file)

            # Extract metadata and content
            original_type = data["metadata"]["original_type"]
            file_name = data["metadata"]["file_name"]
            decoded_content = base64.b64decode(data["content"])

            # Create a folder named "extracted_files" if it doesn't exist
            extracted_folder = "extracted_files"
            if not os.path.exists(extracted_folder):
                os.makedirs(extracted_folder)

            # Save the extracted file locally in the "extracted_files" folder
            output_file = os.path.join(extracted_folder, f"extracted_{file_name}")
            with open(output_file, "wb") as original_file:
                original_file.write(decoded_content)

            # Display content based on file type
            if original_type == "pdf":
                self.display_pdf(output_file)
            else:
                QMessageBox.information(self, "Unsupported File", "File type is not supported for viewing.")
                QApplication.instance().quit()  # Ensure the application exits completely
                sys.exit()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {e}")
            QApplication.instance().quit()  # Ensure the application exits completely
            sys.exit()

    def display_pdf(self, pdf_path):
        """Render a PDF document."""
        try:
            self.pages = []  # Clear any previous pages
            self.page_dimensions = []  # Store original dimensions
            pdf_document = fitz.open(pdf_path)

            for page_num in range(len(pdf_document)):
                page = pdf_document[page_num]
                pix = page.get_pixmap()
                image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                self.pages.append(image)

                # Get original dimensions in points (1 point = 1/72 inch)
                width_points = page.mediabox.width
                height_points = page.mediabox.height

                # Convert points to centimeters (1 point = 0.0352778 cm)
                width_cm = width_points * 0.0352778
                height_cm = height_points * 0.0352778

                self.page_dimensions.append((width_cm, height_cm))

            self.render_all_pages()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to render PDF: {e}")

    def render_all_pages(self):
        """Render all pages on the canvas."""
        max_width_portrait = 800
        max_height_portrait = 1000
        max_width_landscape = 1000
        max_height_landscape = 800

        tolerance = 0.1  # Tolerance for dimension comparison

        for page_num, (page_image, (width_cm, height_cm)) in enumerate(zip(self.pages, self.page_dimensions)):
            # Determine the maximum dimensions based on orientation
            if page_image.width > page_image.height:
                max_width = max_width_landscape
                max_height = max_height_landscape
            else:
                max_width = max_width_portrait
                max_height = max_height_portrait

            # Calculate the scaling factor to fit the image within the maximum dimensions
            scale_factor = min(max_width / page_image.width, max_height / page_image.height, 1)
            new_width = int(page_image.width * scale_factor)
            new_height = int(page_image.height * scale_factor)

            # Resize the image
            resized_image = page_image.resize((new_width, new_height), Image.LANCZOS)

            # Convert PIL image to QImage
            image_qt = QImage(resized_image.tobytes(), resized_image.width, resized_image.height, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(image_qt)

            # Create a QLabel to display the image
            label = QLabel(self)
            label.setPixmap(pixmap)
            label.setAlignment(Qt.AlignCenter)
            self.layout.addWidget(label)

            # Display the page number
            page_label = QLabel(f"Page {page_num + 1}", self)
            page_label.setAlignment(Qt.AlignCenter)
            self.layout.addWidget(page_label)

            # Display the original size of the file in centimeters
            if abs(width_cm - 21.59) < tolerance and abs(height_cm - 27.94) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (Letter: 8.5' x 11')", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)
            
            elif abs(width_cm - 21.59) < tolerance and abs(height_cm - 35.56) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (Legal: 8.5' x 14')", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)
            
            elif abs(width_cm - 14.8) < tolerance and abs(height_cm - 21) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (A5)", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)
            
            elif abs(width_cm - 21) < tolerance and abs(height_cm - 29.7) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (A4)", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)
            
            elif abs(width_cm - 29.7) < tolerance and abs(height_cm - 42) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (A3)", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)

            elif abs(width_cm - 42) < tolerance and abs(height_cm - 59.4) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (A2)", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)
            
            elif abs(width_cm - 59.4) < tolerance and abs(height_cm - 84.1) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (A1)", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)

            elif abs(width_cm - 84.1) < tolerance and abs(height_cm - 118.9) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (A0)", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)
            
            elif abs(width_cm - 10.5) < tolerance and abs(height_cm - 14.8) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (A6)", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)
            
            elif abs(width_cm - 12.7) < tolerance and abs(height_cm - 17.8) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (5' x 7')", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)

            elif abs(width_cm - 15.2) < tolerance and abs(height_cm - 20.3) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (6' x 8')", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)
            
            elif abs(width_cm - 20.3) < tolerance and abs(height_cm - 25.4) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (8' x 10')", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)
            
            elif abs(width_cm - 25.4) < tolerance and abs(height_cm - 30.5) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (10' x 12')", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)
            
            elif abs(width_cm - 30.5) < tolerance and abs(height_cm - 40.6) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (12' x 16')", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)
            
            elif abs(width_cm - 40.6) < tolerance and abs(height_cm - 50.8) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (16' x 20')", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)
            
            elif abs(width_cm - 50.8) < tolerance and abs(height_cm - 61) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (20' x 24')", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)
            
            elif abs(width_cm - 61) < tolerance and abs(height_cm - 76.2) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (24' x 30')", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)
            
            elif abs(width_cm - 76.2) < tolerance and abs(height_cm - 101.6) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (30' x 40')", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)
            
            elif abs(width_cm - 101.6) < tolerance and abs(height_cm - 127) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (40' x 50')", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)
            
            elif abs(width_cm - 127) < tolerance and abs(height_cm - 152.4) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (50' x 60')", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)
            
            else:
                size_label = QLabel(f"Custom Size: {width_cm:.2f} cm x {height_cm:.2f} cm", self)
                size_label.setAlignment(Qt.AlignCenter)
                self.layout.addWidget(size_label)

    def print_document(self):
        """Print the displayed document."""
        printer = QPrinter(QPrinter.HighResolution)
        print_dialog = QPrintDialog(printer, self)
        if print_dialog.exec_() == QPrintDialog.Accepted:
            painter = QPainter(printer)
            for page_image in self.pages:
                image_qt = QImage(page_image.tobytes(), page_image.width, page_image.height, QImage.Format_RGB888)
                pixmap = QPixmap.fromImage(image_qt)
                rect = painter.viewport()
                size = pixmap.size()
                size.scale(rect.size(), Qt.KeepAspectRatio)
                painter.setViewport(rect.x(), rect.y(), size.width(), size.height())
                painter.setWindow(pixmap.rect())
                painter.drawPixmap(0, 0, pixmap)
                printer.newPage()
            painter.end()

    def center_window(self):
        """Centers the window on the screen."""
        screen_geometry = QDesktopWidget().screenGeometry()
        window_geometry = self.frameGeometry()
        window_geometry.moveCenter(screen_geometry.center())  # Move to center
        self.move(window_geometry.topLeft())  # Apply new position

if __name__ == "__main__":
    app = QApplication([])
    viewer = CustomFileViewer()
    viewer.open_file()  # Open the file dialog before showing the main window
    app.exec_()