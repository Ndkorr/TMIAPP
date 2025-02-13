import base64
import json
import sys
from PyQt5.QtWidgets import QDesktopWidget, QToolButton, QPushButton, QLineEdit, QApplication, QMainWindow, QFileDialog, QLabel, QVBoxLayout, QWidget, QScrollArea, QMessageBox, QMenuBar, QMenu, QAction, QTabWidget
from PyQt5.QtGui import QPixmap, QImage, QPainter, QIcon, QColor, QFont, QBrush
from PyQt5.QtCore import Qt
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewDialog
import fitz  # PyMuPDF for PDFs
from PIL import Image
import os
import logging
from PyQt5.QtCore import QPointF


# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')


class WatermarkedWidget(QWidget):
    def __init__(self, parent=None, watermark_text="QFS COPY"):
        super().__init__(parent)
        self.watermark_text = watermark_text
    
    def generate_watermark_pixmap(self, width=200, height=200):
        pixmap = QPixmap(width, height)
        pixmap.fill(Qt.transparent)
        
        painter = QPainter(pixmap)
        painter.setPen(QColor(200, 200, 200, 50))  # Semi-transparent gray
        font = QFont("Arial", 20)
        painter.setFont(font)
        
        center_x = int(width / 2)
        center_y = int(height / 2)
        
        painter.translate(center_x, center_y)
        painter.rotate(-45)
        painter.drawText(-center_x , 0, self.watermark_text)
        painter.end()
        return pixmap
        
    
    def paintEvent(self, event):
        painter = QPainter(self)
        watermark_pixmap = self.generate_watermark_pixmap()
        brush = QBrush(watermark_pixmap)
        painter.fillRect(self.rect(), brush)
        
        super().paintEvent(event)
        
        
class CustomPrintDialog(QPrintDialog):
    def __init__(self, printer, parent=None):
        super().__init__(printer, parent)
        self.setOption(QPrintDialog.PrintPageRange)
        self.buttonBox = self.findChild(QPushButton, "qt_msgbox_buttonbox")
        self.buttonBox.button(QPrintDialog.Accepted).setEnabled(False)
        self.fromPageLineEdit = self.findChild(QLineEdit, "qt_spinbox_lineedit")
        self.toPageLineEdit = self.findChild(QLineEdit, "qt_spinbox_lineedit")
        self.fromPageLineEdit.textChanged.connect(self.validate_page_range)
        self.toPageLineEdit.textChanged.connect(self.validate_page_range)

    def validate_page_range(self):
        try:
            from_page = int(self.fromPageLineEdit.text()) - 1
            to_page = int(self.toPageLineEdit.text()) - 1
            if 0 <= from_page <= to_page < len(self.parent().pages):
                self.buttonBox.button(QPrintDialog.Accepted).setEnabled(True)
            else:
                self.buttonBox.button(QPrintDialog.Accepted).setEnabled(False)
        except ValueError:
            self.buttonBox.button(QPrintDialog.Accepted).setEnabled(False)

    def accept(self):
        self.validate_page_range()
        if self.buttonBox.button(QPrintDialog.Accepted).isEnabled():
            super().accept()


class CustomFileViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Custom File Viewer")
        self.setGeometry(100, 100, 800, 600)
        self.center_window()

        # Initialize pages attribute
        self.pages = []

        # Change the window icon
        # self.setWindowIcon(QIcon('path/to/your/icon.png'))  # Change the icon
        self.setWindowIcon(QIcon())  # Remove the icon
        
        # Add a tab widget
        self.tab_widget = QTabWidget(self)
        self.tab_widget.setTabsClosable(True)
        self.tab_widget.tabCloseRequested.connect(self.close_tab)
        self.setCentralWidget(self.tab_widget)
        
        # Add a menu bar
        self.menu_bar = QMenuBar(self)
        self.setMenuBar(self.menu_bar)

        # Add a "File" menu
        self.file_menu = QMenu("File", self)
        self.menu_bar.addMenu(self.file_menu)

        # Add a "Print" action to the "File" menu
        self.print_action = QAction("Print", self)
        self.print_action.triggered.connect(self.print_document)
        self.file_menu.addAction(self.print_action)
        
        # Add a "Print Preview" action to the "File" menu
        self.print_preview_action = QAction("Print Preview", self)
        self.print_preview_action.triggered.connect(self.print_preview)
        self.file_menu.addAction(self.print_preview_action)
        
        # Add a "Open" action to the "File" menu
        self.open_action = QAction("Open", self)
        self.open_action.triggered.connect(self.open_file)
        self.file_menu.addAction(self.open_action)
        
        # Add a "About" menu
        self.about_menu = QMenu("About", self)
        self.menu_bar.addMenu(self.about_menu)

        self.about_action = QAction("Info", self)
        self.about_action.triggered.connect(self.about)
        self.about_menu.addAction(self.about_action)
        
        # Add a "Toggle Theme" action to the "File" menu
        self.dark_mode = False
        self.toggle_theme_action = QAction("Dark Mode", self)
        self.toggle_theme_action.triggered.connect(self.toggle_theme)
        self.file_menu.addAction(self.toggle_theme_action)
        
        # Add an "Exit" action to the "File" menu
        self.exit_action = QAction("Exit", self)
        self.exit_action.triggered.connect(self.close)
        self.file_menu.addAction(self.exit_action)
        
    
    def print_preview(self):
        """Show a print preview dialog."""
        self.printer = QPrinter(QPrinter.HighResolution)
        self.preview = QPrintPreviewDialog(self.printer, self)
        self.preview.paintRequested.connect(self.print_preview_document)
        self.preview.exec()
        
        
    def open_file(self):
        try:
            file_paths, _ = QFileDialog.getOpenFileNames(self, "Select Custom Files", "", "Custom Files (*.myext);;All Files (*)")
            if file_paths:
                for file_path in file_paths:
                    self.open_viewer_window(file_path)
                self.show()  # Show the main window only if files are selected
            
            elif self.isVisible():
                self.show()
            
            else:
                QMessageBox.information(self, "No File Selected", "No file was selected. Exiting the application.")
                QApplication.instance().quit()  # Ensure the application exits completely
                sys.exit()
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {e}")
        

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

            # Create a new tab for the file
            tab = QWidget()
            layout = QVBoxLayout(tab)
            layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)
            scroll_area = QScrollArea(tab)
            scroll_area.setWidgetResizable(True)
            container = QWidget()
            
            container = WatermarkedWidget(watermark_text="QFS COPY")
            container_layout = QVBoxLayout(container)
            container_layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)
            scroll_area.setWidget(container)
            
            
            layout.addWidget(scroll_area)
            self.tab_widget.addTab(tab, file_name.replace(".pdf", ""))

            # Display content based on file type
            if original_type == "pdf":
                self.display_pdf(output_file, container_layout)
            
            elif self.isVisible():
                QMessageBox.information(self, "Unsupported File", "File type is not supported for viewing.")
                
            else:
                logging.error(f"Unsupported file type: {original_type}")
                QMessageBox.information(self, "Unsupported File", "File type is not supported for viewing.")
                
        except Exception as e:
            logging.error(f"Failed to open viewer window for {custom_file_path}: {e}")
            QMessageBox.critical(self, "Error", f"An error occurred: {e}")


    def draw_tile_watermark(self, painter, rect, watermark_text = "QFS COPY"):
        painter.save()
        # Set the opacity
        painter.setPen(QColor(200, 200, 200, 50))
        font = painter.font()
        font.setPointSize(40)
        painter.setFont(font)
        
        
        # Determine steps for tiling
        step_x = 300
        step_y = 150
        
        # Loop over the area to tile the watermark
        for x in range(rect.left(), rect.right(), step_x):
            for y in range(rect.top(), rect.bottom(), step_y):
                painter.save()
                painter.translate(x, y)
                painter.rotate(-45)
                painter.drawText(0, 0, watermark_text)
                painter.restore()
                
        painter.restore()
        
    def display_pdf(self, pdf_path, layout):
        """Render a PDF document."""
        try:
            self.pages = []  # Clear any previous pages
            page_dimensions = []  # Store original dimensions
            pdf_document = fitz.open(pdf_path)
            self.pdf_path = pdf_path

            for page_num in range(len(pdf_document)):
                page = pdf_document[page_num]
                pix = page.get_pixmap()  # Render at 300 DPI
                image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                self.pages.append(image)

                # Get original dimensions in points (1 point = 1/72 inch)
                width_points = page.mediabox.width
                height_points = page.mediabox.height

                # Convert points to centimeters (1 point = 0.0352778 cm)
                width_cm = width_points * 0.0352778
                height_cm = height_points * 0.0352778

                page_dimensions.append((width_cm, height_cm))

            self.render_all_pages(self.pages, page_dimensions, layout)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to render PDF: {e}")
            
    
    
    def render_all_pages(self, pages, page_dimensions, layout):
        """Render all pages on the canvas.""" 
        max_width_portrait = 800
        max_height_portrait = 1000
        max_width_landscape = 1000
        max_height_landscape = 800

        tolerance = 0.1  # Tolerance for dimension comparison

        for page_num, (page_image, (width_cm, height_cm)) in enumerate(zip(pages, page_dimensions)):
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

            # Add a Watermark
            watermark_painter = QPainter(pixmap)
            watermark_painter.setRenderHint(QPainter.Antialiasing)
            watermark_painter.setRenderHint(QPainter.SmoothPixmapTransform)
            
            # Set the Opacity
            watermark_painter.setOpacity(0.2)
            
            # Rotate Diagonally
            watermark_painter.translate(pixmap.width() / 2, pixmap.height() / 2)
            watermark_painter.rotate(-45)
            
            # Set the Font, size and color
            font = watermark_painter.font()
            font.setFamily("Arial")
            font.setPointSize(72)
            font.setBold(True)
            watermark_painter.setFont(font)
            
            watermark_painter.setPen(Qt.red)
            
            # Draw the Watermark
            watermark_text = "Confidential"
            text_width = watermark_painter.fontMetrics().horizontalAdvance(watermark_text)
            
            # Position the text in the center
            x = int(-text_width // 2)
            watermark_painter.drawText(x, 0, watermark_text)
            
            watermark_painter.end()
            
            
            # Create a QLabel to display the image
            label = QLabel(self)
            label.setPixmap(pixmap)
            label.setAlignment(Qt.AlignCenter)
            layout.addWidget(label)

            # Display the page number
            page_label = QLabel(f"Page {page_num + 1}", self)
            page_label.setAlignment(Qt.AlignCenter)
            layout.addWidget(page_label)

            # Display the original size of the file in centimeters
            if abs(width_cm - 21.59) < tolerance and abs(height_cm - 27.94) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (Letter: 8.5' x 11')", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)
            
            elif abs(width_cm - 21.59) < tolerance and abs(height_cm - 35.56) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (Legal: 8.5' x 14')", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)
            
            elif abs(width_cm - 14.8) < tolerance and abs(height_cm - 21) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (A5)", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)
            
            elif abs(width_cm - 21) < tolerance and abs(height_cm - 29.7) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (A4)", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)
            
            elif abs(width_cm - 29.7) < tolerance and abs(height_cm - 42) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (A3)", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)

            elif abs(width_cm - 42) < tolerance and abs(height_cm - 59.4) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (A2)", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)
            
            elif abs(width_cm - 59.4) < tolerance and abs(height_cm - 84.1) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (A1)", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)

            elif abs(width_cm - 84.1) < tolerance and abs(height_cm - 118.9) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (A0)", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)
            
            elif abs(width_cm - 10.5) < tolerance and abs(height_cm - 14.8) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (A6)", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)
            
            elif abs(width_cm - 12.7) < tolerance and abs(height_cm - 17.8) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (5' x 7')", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)

            elif abs(width_cm - 15.2) < tolerance and abs(height_cm - 20.3) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (6' x 8')", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)
            
            elif abs(width_cm - 20.3) < tolerance and abs(height_cm - 25.4) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (8' x 10')", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)
            
            elif abs(width_cm - 25.4) < tolerance and abs(height_cm - 30.5) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (10' x 12')", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)
            
            elif abs(width_cm - 30.5) < tolerance and abs(height_cm - 40.6) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (12' x 16')", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)
            
            elif abs(width_cm - 40.6) < tolerance and abs(height_cm - 50.8) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (16' x 20')", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)
            
            elif abs(width_cm - 50.8) < tolerance and abs(height_cm - 61) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (20' x 24')", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)
            
            elif abs(width_cm - 61) < tolerance and abs(height_cm - 76.2) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (24' x 30')", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)
            
            elif abs(width_cm - 76.2) < tolerance and abs(height_cm - 101.6) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (30' x 40')", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)
            
            elif abs(width_cm - 101.6) < tolerance and abs(height_cm - 127) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (40' x 50')", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)
            
            elif abs(width_cm - 127) < tolerance and abs(height_cm - 152.4) < tolerance:
                size_label = QLabel(f"Size: {width_cm:.2f} cm x {height_cm:.2f} cm (50' x 60')", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)
            
            else:
                size_label = QLabel(f"Custom Size: {width_cm:.2f} cm x {height_cm:.2f} cm", self)
                size_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(size_label)

    def print_preview_document(self, printer):
        """Render the document for print preview."""
        self.override_print_button()
        painter = QPainter(printer)
        
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.SmoothPixmapTransform)
        
        for i, page_image in enumerate(self.pages):
            if i > 0:
                printer.newPage()
            image_qt = QImage(page_image.tobytes(), page_image.width, page_image.height, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(image_qt)
            rect = painter.viewport()
            size = pixmap.size()
            size.scale(rect.size(), Qt.KeepAspectRatio)
            painter.setViewport(rect.x(), rect.y(), size.width(), size.height())
            painter.setWindow(pixmap.rect())
            painter.drawPixmap(0, 0, pixmap)

        painter.end()
    
    def render_pages_for_print(self):
        "re-render PDF pages at a higher DPI for printing"
        pages_print = []
        pdf_document = fitz.open(self.pdf_path)
        for page_number, page in enumerate(pdf_document):
            try:
                #Attempt to render at 600 dpi
                pix = page.get_pixmap(dpi = 600)
            except Exception as e:
                logging.warning(f"Rendering page {page_number + 1} at 600 DPI failed: {e}. Trying 300 DPI.")
                try:
                    #Fallback to 300 DPI
                    pix = page.get_pixmap(dpi = 300)
                except Exception as e:
                    logging.error(f"Failed to render page {page_number + 1} at 300 DPI: {e}")
                    continue
            image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            pages_print.append(image)
        return pages_print
    
    def print_document(self):
        printer = QPrinter(QPrinter.HighResolution)
        printer.setResolution(600)  # Set high DPI for better quality
        print_dialog = QPrintDialog(printer, self)
        print_dialog.setOption(QPrintDialog.PrintPageRange)
        
        if print_dialog.exec_() == QPrintDialog.Accepted:
            # Get the selected page range
            from_page = print_dialog.fromPage() - 1  # Convert to zero-based index
            to_page = print_dialog.toPage() - 1  # Convert to zero-based index
            
            # Re-render pages at high DPI for printing
            pages_to_print = self.render_pages_for_print()
            
            if from_page == -1 and to_page == -1:
                from_page = 0
                to_page = len(pages_to_print) - 1
            
            
            self.print_selected_pages(printer, pages_to_print, from_page, to_page)
            
            
    def print_selected_pages(self, printer, pages, from_page, to_page):
        """Print the selected pages."""
        painter = QPainter(printer)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.SmoothPixmapTransform)
        
        for i in range(from_page, to_page + 1):
            if i > from_page:
                printer.newPage()
            page_image = pages[i]
            image_qt = QImage(page_image.tobytes(), page_image.width, page_image.height, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(image_qt)
            rect = painter.viewport()
            size = pixmap.size()
            size.scale(rect.size(), Qt.KeepAspectRatio)
            painter.setViewport(rect.x(), rect.y(), size.width(), size.height())
            painter.setWindow(pixmap.rect())
            painter.drawPixmap(0, 0, pixmap)
        
        painter.end()
        
    def override_print_button(self):
        """Override the print button action inside QPrintPreviewDialog."""
        for widget in self.findChildren(QToolButton):
            if widget.text() == "Print":
                widget.setEnabled(False)  # Connect to custom print action
    
    def close_tab(self, index):
        """Close the tab8.10 at the given index."""
        self.tab_widget.removeTab(index)

    def center_window(self):
        """Centers the window on the screen."""
        screen_geometry = QDesktopWidget().screenGeometry()
        window_geometry = self.frameGeometry()
        window_geometry.moveCenter(screen_geometry.center())  # Move to center
        self.move(window_geometry.topLeft())  # Apply new position

    def about(self):
        QMessageBox.about(self, "About this Application",
                          "This application/system is monitored by QFS DEPARTMENT and ADMINISTRATOR. ")

    def toggle_theme(self):
        """Toggles between Dark Mode and Light Mode."""
        if self.dark_mode:
            # Light Mode
            self.setStyleSheet("")
            self.dark_mode = False
        else:
            # Dark Mode
            dark_stylesheet = """
                QMainWindow { background-color: #121212; color: #ffffff; }
                QWidget { background-color: #121212; color: #ffffff; }
                QLabel { color: #ffffff; }
                QPushButton { background-color: #333333; color: #ffffff; border: 1px solid #555555; padding: 5px; }
                QPushButton:hover { background-color: #444444; }
                QPushButton:pressed { background-color: #555555; }
                QTabWidget::pane { border: 1px solid #555555; }
                QTabBar::tab { background-color: #333333; color: #ffffff; padding: 8px; }
                QTabBar::tab:selected { background-color: #555555; }
                QScrollArea { background-color: #121212; }
            """
            self.setStyleSheet(dark_stylesheet)
            self.dark_mode = True


if __name__ == "__main__":
    app = QApplication([])
    
    viewer = CustomFileViewer()
    viewer.open_file()  # Open the file dialog before showing the main window
    app.exec_()
