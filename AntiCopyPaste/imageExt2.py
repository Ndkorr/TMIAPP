import os
import base64
import json
from tkinter import Tk, filedialog, Canvas, Frame, Button, Scrollbar, messagebox, NW
from PIL import Image, ImageTk
from docx import Document
import fitz  # PyMuPDF


class CustomFileViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Custom File Viewer")
        self.root.geometry("800x600")

        # Frame for Navigation and Content
        self.nav_frame = Frame(root)
        self.nav_frame.pack(pady=10)

        self.prev_button = Button(self.nav_frame, text="Previous Page", command=self.show_previous_page, state="disabled")
        self.prev_button.pack(side="left", padx=5)

        self.next_button = Button(self.nav_frame, text="Next Page", command=self.show_next_page, state="disabled")
        self.next_button.pack(side="left", padx=5)

        self.open_button = Button(self.nav_frame, text="Open File", command=self.open_file)
        self.open_button.pack(side="left", padx=5)

        # Canvas for rendering pages
        self.canvas = Canvas(root, bg="white")
        self.canvas.pack(fill="both", expand=True)

        # Scrollbar
        self.scrollbar = Scrollbar(self.canvas, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Internal Variables
        self.pages = []  # Store rendered pages (images)
        self.current_page = 0

    def open_file(self):
        file_path = filedialog.askopenfilename(
            title="Select a Custom File",
            filetypes=(("Custom Files", "*.myext"), ("All Files", "*.*"))
        )
        if file_path:
            self.process_file(file_path)

    def process_file(self, custom_file_path):
        try:
            # Read and decode the .myext file
            with open(custom_file_path, "r") as custom_file:
                data = json.load(custom_file)

            # Extract metadata and content
            original_type = data["metadata"]["original_type"]
            file_name = data["metadata"]["file_name"]
            decoded_content = base64.b64decode(data["content"])

            # Save the extracted file locally
            extracted_file = f"extracted_{file_name}"
            with open(extracted_file, "wb") as original_file:
                original_file.write(decoded_content)

            # Display the file based on type
            if original_type == "docx":
                self.render_docx(extracted_file)
            elif original_type == "pdf":
                self.render_pdf(extracted_file)
            else:
                messagebox.showinfo("Unsupported File", f"File type '{original_type}' is not supported.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def render_docx(self, docx_path):
        """Render a Word document."""
        try:
            self.pages = []  # Clear any previous pages
            document = Document(docx_path)

            for paragraph in document.paragraphs:
                # For simplicity, render paragraphs as lines on an image
                page_image = Image.new("RGB", (800, 100), "white")
                draw = ImageDraw.Draw(page_image)
                draw.text((10, 10), paragraph.text, fill="black")
                self.pages.append(page_image)

            self.current_page = 0
            self.display_page()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to render Word document: {e}")

    def render_pdf(self, pdf_path):
        """Render a PDF document."""
        try:
            self.pages = []  # Clear any previous pages
            pdf_document = fitz.open(pdf_path)

            for page_num in range(len(pdf_document)):
                page = pdf_document[page_num]
                pix = page.get_pixmap()
                image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                self.pages.append(image)

            self.current_page = 0
            self.display_page()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to render PDF: {e}")

    def display_page(self):
        """Display the current page on the canvas."""
        if 0 <= self.current_page < len(self.pages):
            page_image = self.pages[self.current_page]
            page_tk = ImageTk.PhotoImage(page_image)

            self.canvas.delete("all")
            self.canvas.create_image(10, 10, anchor=NW, image=page_tk)
            self.canvas.image = page_tk

            # Enable/Disable navigation buttons
            self.prev_button.config(state="normal" if self.current_page > 0 else "disabled")
            self.next_button.config(state="normal" if self.current_page < len(self.pages) - 1 else "disabled")

    def show_previous_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.display_page()

    def show_next_page(self):
        if self.current_page < len(self.pages) - 1:
            self.current_page += 1
            self.display_page()


# Run the GUI application
if __name__ == "__main__":
    root = Tk()
    app = CustomFileViewer(root)
    root.mainloop()
