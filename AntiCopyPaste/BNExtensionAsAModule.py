import base64
import json
import os
from tkinter import Tk, filedialog
import tkinter.messagebox as tkmb
from tkinter import ttk
import time
import tkinter as tk

# PDF CONVERSION
def convert_pdf(file_path, output_path, progress_callback=None):
    # Import the required module
    import fitz

    # To get the metadata for its orientation and size (cm)
    doc = fitz.open(file_path)
    first_page = doc[0]
    rect = first_page.rect  # Page dimensions in points
    width_cm = rect.width * 0.0352778  # Convert points to cm
    height_cm = rect.height * 0.0352778
    orientation = "Landscape" if width_cm > height_cm else "Portrait"

    with open(file_path, "rb") as f_in:
        content = f_in.read()
        
        total_steps = 100
        for step in range(total_steps):
            time.sleep(0.1)
            if progress_callback:
                progress_callback(step + 1)

        encoded_content = base64.b64encode(content).decode("utf-8")

        file_metadata = {
            "metadata": {
                "original_type": file_path.split(".")[-1],
                "file_name": os.path.basename(file_path),
                "orientation": orientation,
                "height": f"{height_cm:.2f}cm",
                "width": f"{width_cm:.2f}cm",
                "author": doc.metadata.get("author", ""),
                "creation_date": doc.metadata.get("creationDate", "")
            },
            "content": encoded_content
        }

        with open(output_path, "w") as f_out:
            json.dump(file_metadata, f_out)

def main():
    root = Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(
        # LIMITING THE FILE TYPES TO PDF ONLY
        title="Select a PDF file to convert",
        filetypes=(("PDF Files", "*.pdf"),)
    )
    if not file_path:
        tkmb.showinfo("Cancelled", "File selection was cancelled")
        return
        
    output_path = filedialog.asksaveasfilename(
        title="New extension name",
        defaultextension=".QFS",
        filetypes=(("QFS", "*.QFS"), ("All Files", "*.*"))
    )
    if not output_path:
        tkmb.showinfo("Cancelled", "Save operation was cancelled")
        return
    
    # Create a progress dialog
    progress_win = tk.Toplevel()
    progress_win.title("Creating...")
    progress_win.resizable(False, False)
    
    progress_label = tk.Label(progress_win, text="Creating Custom file, please wait...")
    progress_label.pack(padx=20, pady=(10, 5))
    
    # Progress bar widget
    progress_bar = ttk.Progressbar(progress_win, orient="horizontal", length=300, mode="determinate")
    progress_bar.pack(padx=20, pady=5)
    progress_bar["value"] = 0
    progress_bar["maximum"] = 100

    # Label to show percentage
    percentage_label = tk.Label(progress_win, text="0%")
    percentage_label.pack(padx=20, pady=(5, 10))
    
    # Center the progress window on screen
    progress_win.update_idletasks()
    x = (progress_win.winfo_screenwidth() // 2) - (progress_win.winfo_width() // 2)
    y = (progress_win.winfo_screenheight() // 2) - (progress_win.winfo_height() // 2)
    progress_win.geometry(f"+{x}+{y}")
    
    def update_progress(percent):
        progress_bar["value"] = percent
        percentage_label.config(text=f"{percent}%")    
        progress_win.update()
    
    try:
        convert_pdf(file_path, output_path, progress_callback=update_progress)
        progress_win.destroy()
        tkmb.showinfo("Success", f"File converted and saved to {output_path}")
    except Exception as e:
        progress_win.destroy()
        tkmb.showerror("Error", f"Conversion failed: {e}")

if __name__ == "__main__":
    main()
