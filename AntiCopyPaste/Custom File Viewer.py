import base64
import json
import sys
from PyQt5.QtWidgets import QDesktopWidget, QToolButton, QPushButton, QLineEdit, QApplication, QMainWindow, QFileDialog, QLabel, QVBoxLayout, QWidget, QScrollArea, QMessageBox, QMenuBar, QMenu, QAction, QTabWidget, QHBoxLayout, QPushButton, QDialog
from PyQt5.QtGui import QPixmap, QImage, QPainter, QIcon, QColor, QFont, QBrush
from PyQt5.QtCore import Qt, QRect
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewDialog
import fitz  # PyMuPDF for PDFs
from PIL import Image
import os
import logging
from PyQt5.QtCore import QPointF
import BNExtensionAsAModule


# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')


class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.setWindowTitle("How to Use This Application")
        
        screen_geometry = QDesktopWidget().screenGeometry()
        # Get the current screen resolution in terms of width and height
        screen_width = screen_geometry.width()
        screen_height = screen_geometry.height()
        
        # Calculate desired window size
        reduce_width = int(screen_width * 0.8)
        reduce_height = int(screen_height * 0.8)
        
        self.resize(reduce_width, reduce_height)
        
        window_geometry = self.frameGeometry()
        window_geometry.moveCenter(screen_geometry.center())
        self.move(window_geometry.topLeft())
        
        
        
        #self.setGeometry(100, 100, 600, 400)
        
        # List of help images (file paths or QPixmaps)
        self.help_images = [
            "1.png",
            "2.png",
            "3.png",
            "4.png",
            "5.png",
            "6.png"
        ]
        self.current_index = 0
        
        # Create label to display image
        self.image_label = QLabel(self)
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setMinimumSize(int(reduce_width * 0.8), int(reduce_height * 0.8))
        self.update_image()
        
        # Create navigation buttons
        self.prev_button = QPushButton("Previous", self)
        self.next_button = QPushButton("Next", self)
        self.prev_button.clicked.connect(self.show_previous)
        self.next_button.clicked.connect(self.show_next)
        
        # Layout
        btn_layout = QHBoxLayout()
        btn_layout.addWidget(self.prev_button)
        btn_layout.addWidget(self.next_button)
        
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.image_label)
        main_layout.addLayout(btn_layout)
        self.setLayout(main_layout)
        
        self.update_buttons()
    
    def update_image(self):
        if self.help_images:
            pixmap = QPixmap(self.help_images[self.current_index])
            self.image_label.setPixmap(pixmap.scaled(self.image_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation))
    
    def show_previous(self):
        if self.current_index > 0:
            self.current_index -= 1
            self.update_image()
            self.update_buttons()
    
    def show_next(self):
        if self.current_index < len(self.help_images) - 1:
            self.current_index += 1
            self.update_image()
            self.update_buttons()
    
    def update_buttons(self):
        self.prev_button.setEnabled(self.current_index > 0)
        self.next_button.setEnabled(self.current_index < len(self.help_images) - 1)
    
    def resizeEvent(self, event):
        # Update the image scaling on resize
        self.update_image()
        super().resizeEvent(event)


class WatermarkedWidget(QWidget):
    def __init__(self, parent=None, watermark_text="QFS COPY"):
        super().__init__(parent)
        self.watermark_text = watermark_text
    
    def generate_watermark_pixmap(self, width=200, height=200):
        pixmap = QPixmap(width, height)
        pixmap.fill(Qt.transparent)
        
        painter = QPainter(pixmap)
        painter.setPen(QColor(200, 200, 200, 90))  # Semi-transparent gray
        font = QFont("Arial", 30)
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

class LoginDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Login")
        self.setup_ui()
        
    def setup_ui(self):
        # Create labels and line edits
        self.user_label = QLabel("Username:")
        self.user_edit = QLineEdit()
        
        self.pass_label = QLabel("Password:")
        self.pass_edit = QLineEdit()
        self.pass_edit.setEchoMode(QLineEdit.Password)
        
        # Create buttons
        self.ok_button = QPushButton("OK")
        self.cancel_button = QPushButton("Cancel")
        
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)
        
        # Layouts for organization
        form_layout = QVBoxLayout()
        form_layout.addWidget(self.user_label)
        form_layout.addWidget(self.user_edit)
        form_layout.addWidget(self.pass_label)
        form_layout.addWidget(self.pass_edit)
        
        button_layout = QHBoxLayout()
        button_layout.addStretch(1)
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        
        main_layout = QVBoxLayout()
        main_layout.addLayout(form_layout)
        main_layout.addLayout(button_layout)
        
        self.setLayout(main_layout)
    
    def get_credentials(self):
        return self.user_edit.text(), self.pass_edit.text()
    
    def validate_credentials(self):
        # Insert your credential validation logic here.
        username, password = self.get_credentials()
        # For example, using simple static check:
        if username == "admin" and password == "QFSADMIN":
            return True
        return False
    
    def accept(self):
        if self.validate_credentials():
            super().accept()
        else:
            QMessageBox.warning(self, "Login Failed", "Invalid username or password")

class CustomFileViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Custom File Viewer")
        screen_geometry = QDesktopWidget().screenGeometry()
        # Get the current screen resolution in terms of width and height
        screen_width = screen_geometry.width()
        screen_height = screen_geometry.height()
        
        # Calculate desired window size
        reduce_width = int(screen_width * 0.6)
        reduce_height = int(screen_height * 0.6)
        
        self.resize(reduce_width, reduce_height)
        
        window_geometry = self.frameGeometry()
        window_geometry.moveCenter(screen_geometry.center())
        self.move(window_geometry.topLeft())

        # Initialize pages attribute
        self.pages = []
        
        # Add a drag & drop function
        self.setAcceptDrops(True)
        
        

        # Change the window icon
        # self.setWindowIcon(QIcon('path/to/your/icon.png'))  # Change the icon
        self.setWindowIcon(QIcon("iconFileViewer512.ico"))  # Change the icon
    
        # Add a tab widget
        self.tab_widget = QTabWidget(self)
        self.tab_widget.setTabsClosable(True)
        self.tab_widget.tabCloseRequested.connect(self.close_tab)
        self.setCentralWidget(self.tab_widget)
        
        # Add an introductory instructions on viewer
        self.add_instructions_tab()
        
        # Add a menu bar
        self.menu_bar = QMenuBar(self)
        self.setMenuBar(self.menu_bar)

        # Add a "File" menu
        self.file_menu = QMenu("File", self)
        self.menu_bar.addMenu(self.file_menu)

        # Add a "Open" action to the "File" menu
        self.open_action = QAction("Open", self)
        self.open_action.triggered.connect(self.open_file)
        self.file_menu.addAction(self.open_action)
        
        # Add a "QFS" action to the "File" menu
        self.qfs_action = QAction("QFS", self)
        self.qfs_action.triggered.connect(self.qfs_open)
        self.file_menu.addAction(self.qfs_action)
        
        # Add a "Print" action to the "File" menu
        self.print_action = QAction("Print", self)
        self.print_action.triggered.connect(self.print_document)
        self.file_menu.addAction(self.print_action)
        
        # Add a "Print Preview" action to the "File" menu
        self.print_preview_action = QAction("Print Preview", self)
        self.print_preview_action.triggered.connect(self.print_preview)
        self.file_menu.addAction(self.print_preview_action)
        
        # Add a "About" menu
        self.about_menu = QMenu("About", self)
        self.menu_bar.addMenu(self.about_menu)

        self.about_action = QAction("Info", self)
        self.about_action.triggered.connect(self.about)
        self.about_menu.addAction(self.about_action)
        
        self.help_action = QAction("Help", self)
        self.help_action.triggered.connect(self.help)
        self.about_menu.addAction(self.help_action)
        
        # Add a "Toggle Theme" action to the "File" menu
        self.dark_mode = False
        self.toggle_theme_action = QAction("Dark Mode", self)
        self.toggle_theme_action.triggered.connect(self.toggle_theme)
        self.file_menu.addAction(self.toggle_theme_action)
        
        # Add an "Exit" action to the "File" menu
        self.exit_action = QAction("Exit", self)
        self.exit_action.triggered.connect(self.close)
        self.file_menu.addAction(self.exit_action)
        
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.endswith(".QFS"):
                self.open_viewer_window(file_path)
            else:
                QMessageBox.information(self, "Invalid File", "File Not Supported")
        event.acceptProposedAction()
    
        
    def print_preview(self):
        """Show a print preview dialog."""
        self.printer = QPrinter(QPrinter.HighResolution)
        self.preview = QPrintPreviewDialog(self.printer, self)
        self.preview.paintRequested.connect(self.print_preview_document)
        self.preview.exec()
        
        
    def open_file(self):
        if self.tab_widget.count() == 1 and self.tab_widget.tabText(0) == "Instructions":
            self.tab_widget.removeTab(0)
        try:
            file_paths, _ = QFileDialog.getOpenFileNames(self, "Select Custom Files", "", "Custom Files (*.QFS);;All Files (*)")
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
            
            
    def qfs_open(self):
        login = LoginDialog(self)
        if login.exec_() == QDialog.Accepted:
            BNExtensionAsAModule.main()
        
    def add_instructions_tab(self):
        """Add a tab with initial instructions."""
        instructions_tab = QWidget()
        layout = QVBoxLayout(instructions_tab)
        instructions_label = QLabel("To get started, click File then Open.", self)
        instructions_label.setAlignment(Qt.AlignCenter)
        instructions_label.setText("To get started, click File then Open.<br><br>"
                               "<a href='#' style='color:#0000FF;'>or show me how to use this application</a>")
        instructions_label.setOpenExternalLinks(False)  # We'll handle link click ourselves
        instructions_label.linkActivated.connect(self.show_usage_instructions)
       
        instructions_label.setFont(QFont("Arial", 16))
        layout.addWidget(instructions_label)
        self.tab_widget.addTab(instructions_tab, "Instructions")
    
    def show_usage_instructions(self):
        # This function can open another dialog or tab explaining the application usage.
        help_dialog = HelpDialog(self)
        help_dialog.exec_()
    
    
    
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
            # STORE IN %APPDATA%
            appdata_folder = os.getenv("APPDATA")
            
            extracted_folder = os.path.join(appdata_folder, "CFV")
            
            subfolder_name = "extracted files"
            subfolder_path = os.path.join(extracted_folder, subfolder_name)
            
            os.makedirs(subfolder_path, exist_ok=True)

            # Save the extracted file locally in the "extracted_files" folder
            output_file = os.path.join(subfolder_path, f"extracted_{file_name}")
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
    
    def render_pages_for_print(self, start_page, end_page):
        "Re-render PDF pages at various DPI settings for printing"
        pages_print = []
        pdf_document = fitz.open(self.pdf_path)
        dpi_options = [600, 300, None]  # None for default DPI
        
        # Ensure we stay within bounds
        end_page = min(end_page, len(pdf_document) - 1)
        start_page = max(start_page, 0)

        for page_number in range(start_page, end_page + 1):
            page = pdf_document[page_number]
            pix = None
            for dpi in dpi_options:
                try:
                    if dpi is None:
                        pix = page.get_pixmap()  # default DPI
                    else:
                        pix = page.get_pixmap(dpi=dpi)
                    # If rendering succeeds, break out of the loop
                    logging.info(f"Rendered page {page_number+1} at {'default' if dpi is None else dpi} DPI.")
                    break
                except Exception as e:
                    logging.warning(f"Failed to render page {page_number+1} at {'default' if dpi is None else dpi} DPI: {e}")
            if pix is None:
                logging.error(f"All rendering attempts failed for page {page_number+1}. Skipping this page.")
                continue

            image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            pages_print.append(image)
        
        pdf_document.close()
        return pages_print

    
    def print_document(self):
        printer = QPrinter(QPrinter.HighResolution)
        printer.setResolution(600)  # Set high DPI for better quality
        # set page margins
        printer.setPageMargins(0, 0, 30, 0, QPrinter.Millimeter)
        print_dialog = QPrintDialog(printer, self)
        print_dialog.setOption(QPrintDialog.PrintPageRange)
        
        if print_dialog.exec_() == QPrintDialog.Accepted:
            # Get the selected page range
            from_page = print_dialog.fromPage() - 1  # Convert to zero-based index
            to_page = print_dialog.toPage() - 1  # Convert to zero-based index
            
            # if the user did not specify a range, print all
            if from_page == -1 and to_page == -1:
                from_page = 0
                # Know how many page are there
                total_pages = self.get_pdf_page_count(self.pdf_path)
                to_page = total_pages - 1
                
            # Re-render pages at high DPI for printing in selected range
            pages_to_print = self.render_pages_for_print(from_page, to_page)
            
            self.print_selected_pages(printer, pages_to_print, 0, (to_page - from_page))
            
            
    def print_selected_pages(self, printer, pages, from_page, to_page):
        """Print the selected pages."""
        painter = QPainter(printer)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.SmoothPixmapTransform)
        
        for i in range(from_page, to_page + 1):
            if i > from_page:
                printer.newPage()
            page_image = pages[i]
            # Ensure image is in RGB Mode
            if page_image.mode != "RGB":
                page_image = page_image.convert("RGB")
            bytes_per_line = page_image.width * 3    
            image_qt = QImage(page_image.tobytes(), page_image.width, page_image.height, bytes_per_line, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(image_qt)
            
            # Calculate target rect based on printer page react
            page_rect = printer.pageRect()
            scaled_size = pixmap.size()
            scaled_size.scale(page_rect.size(), Qt.KeepAspectRatio)
            pixmap_size = pixmap.size()
            
            x = (page_rect.x() + (page_rect.width() - pixmap_size.width()) - 200) // 2
            y = (page_rect.y() + (page_rect.height() - pixmap_size.height())) // 2
            
            if x >= 0 and x <= 1000:
                target_rect = QRect(-175, y, pixmap_size.width(), pixmap_size.height())
            else:
                target_rect = QRect(x, y, pixmap_size.width(), pixmap_size.height())
            
            logging.debug(f"Drawing page {i+1}: target rect {target_rect}")
            
            painter.drawPixmap(target_rect, pixmap)
        
        painter.end()
    
    def get_pdf_page_count(self, pdf_path):
        doc = fitz.open(pdf_path)
        count = len(doc)
        doc.close()
        return count
        
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

    def help(self):
        dialog = HelpDialog(self)
        dialog.exec_()
        
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
                QTabBar::tab { background-color: #333333; color:rgb(255, 255, 255); padding: 8px; }
                QTabBar::tab:selected { background-color: #555555; }
                QScrollArea { background-color: #121212; }
            """
            self.setStyleSheet(dark_stylesheet)
            self.dark_mode = True


if __name__ == "__main__":
    app = QApplication([])
    app.setWindowIcon(QIcon("iconFileViewer512.ico"))
    viewer = CustomFileViewer()
    viewer.show()  # Open the file dialog before showing the main window
    app.exec_()