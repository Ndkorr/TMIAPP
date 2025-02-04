import base64
import json
import os
from tkinter import Tk, filedialog, Scrollbar, messagebox, Toplevel, Canvas, Frame, NW
from PIL import Image, ImageTk
import fitz
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

        # Canvas for rendering pages
        self.canvas = Canvas(viewer_window, bg="white")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Scrollbar
        self.scrollbar = Scrollbar(viewer_window, orient="vertical", command=self.canvas.yview)
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
                self.display_pdf(output_file)
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

    def display_pdf(self, pdf_path):
        """Render a PDF document."""
        try:
            self.pages = []  # Clear any previous pages
            self.page_images = []  # Store references to PhotoImage objects
            pdf_document = fitz.open(pdf_path)

            for page_num in range(len(pdf_document)):
                page = pdf_document[page_num]
                pix = page.get_pixmap()
                image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                self.pages.append(image)

            self.render_all_pages()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to render PDF: {e}")

    def render_all_pages(self):
        """Render all pages on the canvas."""
        y_position = 10  # Initial y position for the first page

        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()

        for page_image in self.pages:
            # Calculate the scaling factor to fit the image within the canvas
            scale_factor = min(canvas_width / page_image.width, canvas_height / page_image.height, 1)
            new_width = int(page_image.width * scale_factor)
            new_height = int(page_image.height * scale_factor)
            resized_image = page_image.resize((new_width, new_height), Image.LANCZOS)

            page_tk = ImageTk.PhotoImage(resized_image)
            self.page_images.append(page_tk)  # Keep a reference to avoid garbage collection

            # Calculate the x position to center the page
            x_position = (canvas_width - new_width) // 2

            self.canvas.create_image(x_position, y_position, anchor=NW, image=page_tk)
            y_position += new_height + 10  # Update y position for the next page

        # Update the scroll region to include all pages
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

        # Bind the configure event to update the canvas size dynamically
        self.canvas.bind("<Configure>", self.on_canvas_configure)

    def on_canvas_configure(self, event):
        """Update the canvas when the window is resized."""
        self.canvas.delete("all")
        self.render_all_pages()

# Run the GUI application
if __name__ == "__main__":
    root = Tk()
    app = CustomFileViewer(root)
    root.mainloop()