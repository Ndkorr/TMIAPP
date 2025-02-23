import base64
import json
import os
from tkinter import Tk, filedialog


# PDF CONVERSION
def convert_pdf(file_path, output_path):
    # Import the required module
    import fitz

    # To get the metadata for its orientation and size (cm)
    doc = fitz.open(file_path)
    first_page = doc[0]
    rect = first_page.rect  # Page dimensions in points
    width_cm = rect.width * 0.0352778  # Convert points to cm
    height_cm = rect.height * 0.0352778
    orientation = "Landscape" if width_cm > height_cm else "Portrait"

    with open(file_path, "rb") as file:
        content = file.read()
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

        with open(output_path, "w") as output_file:
            json.dump(file_metadata, output_file)


# EXCEL CONVERSION
##def convert_excel(file_path, output_path):
    # Import the required module
    ##from openpyxl import load_workbook

    # To get the metadata for its orientation and size (cm)
    ##wb = load_workbook(file_path)
    ##ws = wb.active
    # Get paper size and orientation
    ##paper_size = ws.page_setup.paperSize  # Numeric code for paper size
    ##orientation = ws.page_setup.orientation
    ##orientation = "Landscape" if orientation == "landscape" else "Portrait"
    
    # Paper size mapping (assuming default sizes)
    ##paper_sizes = {
        ##1: {"width_cm": 21.59, "height_cm": 27.94},  # Letter
        ##9: {"width_cm": 21.0, "height_cm": 29.7},    # A4
        ##11: {"width_cm": 14.8, "height_cm": 21.0},   # A5
        ##5: {"width_cm": 21.59, "height_cm": 35.56},  # Legal
        ##3: {"width_cm": 27.94, "height_cm": 43.18},  # Tabloid
        ##4: {"width_cm": 43.18, "height_cm": 27.94},  # Ledger
        ##6: {"width_cm": 18.41, "height_cm": 26.67},  # Executive
        ##7: {"width_cm": 19.05, "height_cm": 25.4},   # Folio
        ##8: {"width_cm": 25.4, "height_cm": 33.02},   # Quarto
        ##13: {"width_cm": 18.2, "height_cm": 25.7},   # B5
    ##}
    ##size = paper_sizes.get(paper_size, {"width_cm": None, "height_cm": None})
    ##size["orientation"] = orientation

    ##with open(file_path, "rb") as excel_file:
        ##binary_content = excel_file.read()
        ##encoded_content = base64.b64encode(binary_content).decode("utf-8")
        ##file_size = os.path.getsize(file_path)
    
    # Create metadata
    ##custom_file = {
        ##"metadata": {
            ##"original_type": "xlsx",
            ##"file_name": os.path.basename(file_path),
            ##"file_size": file_size,
            ##"orientation": size["orientation"],
            ##"height": f"{size['height_cm']:.2f}cm" if size["height_cm"] else "Unknown",
            ##"width": f"{size['width_cm']:.2f}cm" if size["width_cm"] else "Unknown",
            ##"author": wb.properties.creator,
            ##"creation_date": wb.properties.created.strftime("%Y-%m-%d %H:%M:%S")
        ##},
        ##"content": encoded_content
    ##}
    
    # Save to custom format
    ##with open(output_path, "w") as output_file:
        ##json.dump(custom_file, output_file)


# POWERPOINT CONVERSION
##def convert_ppt(file_path, output_path):

    # Import the required module
    ##from pptx import Presentation
    
    # To get the metadata for its orientation
    ##prs = Presentation(file_path)
    ##slide_width = prs.slide_width.cm  # Width in cm
    ##slide_height = prs.slide_height.cm  # Height in cm
    ##orientation = "Landscape" if slide_width > slide_height else "Portrait"

    ##with open(file_path, "rb") as ppt_file:
        ##binary_content = ppt_file.read()
        ##encoded_content = base64.b64encode(binary_content).decode("utf-8")
    
    # Create metadata
    ##custom_file = {
        ##"metadata": {
            ##"original_type": "pptx",
            ##"file_name": os.path.basename(file_path),
            ##"orientation": orientation,
            ##"author": prs.core_properties.author,
            ##"creation_date": prs.core_properties.created.strftime("%Y-%m-%d %H:%M:%S")
        ##},
        ##"content": encoded_content
    ##}
    
    # Save to custom format
    ##with open(output_path, "w") as output_file:
        ##json.dump(custom_file, output_file)


# WORD DOCUMENT CONVERSION
#def convert_docx(file_path, output_path):
    # Import the required module
    ##from docx import Document

    # To get the metadata (DOCX) for its orientation and size 
    ##doc = Document(file_path)
    ##section = doc.sections[0]
    ##page_width = section.page_width.cm  # Width in cm
    ##page_height = section.page_height.cm  # Height in cm
    ##orientation = "Landscape" if page_width > page_height else "Portrait"

    ##with open(file_path, "rb") as file:
        ##binary_content = file.read()
        ##encoded_content = base64.b64encode(binary_content).decode("utf-8")
    
    # Get file metadata
    ##custom_file = {
        ##"metadata": {
            ##"original_type": "docx",
            ##"file_name": os.path.basename(file_path),
            ##"orientation": orientation,
            ##"height": f"{page_height:.2f}cm",
            ##"width": f"{page_width:.2f}cm",
            ##"author": doc.core_properties.author,
            ##"creation_date": doc.core_properties.created.strftime("%Y-%m-%d %H:%M:%S")
        ##},
        ##"content": encoded_content
    ##}
    
    # Save to custom format
    ##with open(output_path, "w") as output_file:
        ##json.dump(custom_file, output_file)

# ANY FORMAT CONVERSION
# This function will determine the file type and call the appropriate conversion function
# IF AND ELSE STATEMENTS
def convert_to_custom_format(file_path, output_path):
    # Get the file extension
    file_extension = file_path.split(".")[-1]

    # Check the file extension and call the appropriate conversion function
    if file_extension == "pdf":
        convert_pdf(file_path, output_path)
    ##elif file_extension == "docx":
    ##    convert_docx(file_path, output_path)
    ##elif file_extension == "pptx":
    ##    convert_ppt(file_path, output_path)
    ##elif file_extension == "xlsx":
    ##    convert_excel(file_path, output_path)
    else:
        print(f"Unsupported file type: {file_extension}")




#convert_to_custom_format(r"C:\Users\HR-IT-MATTHEW-PC\Desktop\Projects\Applications\AntiCopyPaste\try.docx", "docx.myext")

def main():
    root = Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(
        # LIMITING THE FILE TYPES TO PDF ONLY
        title="Select a PDF file to convert",
        filetypes=(("PDF Files", "*.pdf"),)
    )
    if file_path:
        output_path = filedialog.asksaveasfilename(
            title="New extension name",
            defaultextension=".myext",
            filetypes=(("myext", "*.myext"), ("All Files", "*.*"))
        )
        if output_path:
            convert_to_custom_format(file_path, output_path)
            print(f"File converted and saved to {output_path}")
        else:
            print("Save operation cancelled.")
    else:
        print("File selection cancelled.")

if __name__ == "__main__":
    main()


    ## NOT3 ALL THE FUNCTIONS ARE COMMENTED OUT BECAUSE THEY ARE NOT USED IN THE MAIN FUNCTION AND WILL CAUSE ERRORS ON
    ## IMPLEMENTATIONS, THUS THEY ARE COMMENTED OUT. PDF CONVERSION IS THE ONLY FUNCTION USED IN THE MAIN FUNCTION BECAUSE
    ## IT RETAINS ALL FORMATTING AND IS THE MOST RELIABLE FORMAT TO CONVERT TO. THE REST OF THE FUNCTIONS ARE COMMENTED OUT


print("Current working directory: ", os.getcwd())