import base64
import json
from tkinter import Tk, filedialog, Button, Text, Scrollbar, messagebox, Toplevel, Canvas, Frame, NW, Label
import fitz  # PyMuPDF for PDFs
from PIL import Image, ImageTk

# Main GUI Viewer Class
class CustomFileViewer:
    def __init__(self, root=None):
        if root:
            self.root = root
            self.root.title("Custom File Viewer")
            self.root.withdraw()
        else:
            self.root = Tk()
            self.root.withdraw()  # Hide the main Tkinter window

        # Directly open file dialog and process the file
        self.open_file()

    def open_file(self):
        file_path = filedialog.askopenfilename(
            title="Select a Custom File",
            filetypes=(("Custom Files", "*.myext"), ("All Files", "*.*"))
        )
        if file_path:
            self.open_viewer_window(file_path)

    def open_viewer_window(self, custom_file_path):
        viewer_window = Toplevel(self.root)
        viewer_window.title("File Viewer")
        viewer_window.geometry("800x600")
        viewer_window.update_idletasks()
        width = viewer_window.winfo_width()
        height = viewer_window.winfo_height()
        x = (viewer_window.winfo_screenwidth() // 2) - (width // 2)
        y = (viewer_window.winfo_screenheight() // 2) - (height // 2)
        viewer_window.geometry(f'{width}x{height}+{x}+{y}')

        # Frame for Navigation and Content
        self.nav_frame = Frame(viewer_window)
        self.nav_frame.pack(pady=10)

        self.prev_button = Button(self.nav_frame, text="Previous Page", command=self.show_previous_page, state="disabled")
        self.prev_button.pack(side="left", padx=5)

        self.page_label = Label(self.nav_frame, text="Page 0 of 0")
        self.page_label.pack(side="left", padx=5)

        self.next_button = Button(self.nav_frame, text="Next Page", command=self.show_next_page, state="disabled")
        self.next_button.pack(side="left", padx=5)

        # Canvas for rendering pages
        self.canvas = Canvas(viewer_window, bg="white")
        self.canvas.pack(fill="both", expand=True)

        # Scrollbar
        self.scrollbar = Scrollbar(self.canvas, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        try:
            # Handle window close event
            def on_closing():
                self.root.quit()
                self.root.destroy()

            # Read and decode the .myext file
            with open(custom_file_path, "r") as custom_file:
                data = json.load(custom_file)

            # Extract metadata and content
            original_type = data["metadata"]["original_type"]
            file_name = data["metadata"]["file_name"]
            decoded_content = base64.b64decode(data["content"])

            # Save the extracted file locally
            output_file = f"extracted_{file_name}"
            with open(output_file, "wb") as original_file:
                original_file.write(decoded_content)

            # Display content based on file type
            if original_type == "pdf":
                self.render_pdf(output_file)
            else:
                messagebox.showinfo("Unsupported File", f"File type is not supported for viewing.")
                self.root.quit()
                self.root.destroy()
                return

            viewer_window.protocol("WM_DELETE_WINDOW", on_closing)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            self.root.quit()
            self.root.destroy()

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

            # Update page label
            self.page_label.config(text=f"Page {self.current_page + 1} of {len(self.pages)}")

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