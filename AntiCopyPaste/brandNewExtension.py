import base64
import json
import os


# PDF CONVERSION
def convert_to_custom_format(file_path, output_path):
    with open(file_path, "rb") as file:
        content = file.read()
        encoded_content = base64.b64encode(content).decode("utf-8")

        file_metadata = {
            "metadata": {
                "original_type": file_path.split(".")[-1],
                "file_name": os.path.basename(file_path),
            },
            "content": encoded_content
        }

        with open(output_path, "w") as output_file:
            json.dump(file_metadata, output_file)


#  EXCEL CONVERSION
def convert_excel(file_path, output_path):
    with open(file_path, "rb") as excel_file:
        binary_content = excel_file.read()
        encoded_content = base64.b64encode(binary_content).decode("utf-8")
    
    # Create metadata
    custom_file = {
        "metadata": {
            "original_type": "xlsx",
            "file_name": os.path.basename(file_path)
        },
        "content": encoded_content
    }
    
    # Save to custom format
    with open(output_path, "w") as output_file:
        json.dump(custom_file, output_file)


# POWERPOINT CONVERSION
def convert_ppt(file_path, output_path):
    with open(file_path, "rb") as ppt_file:
        binary_content = ppt_file.read()
        encoded_content = base64.b64encode(binary_content).decode("utf-8")
    
    # Create metadata
    custom_file = {
        "metadata": {
            "original_type": "pptx",
            "file_name": os.path.basename(file_path)
        },
        "content": encoded_content
    }
    
    # Save to custom format
    with open(output_path, "w") as output_file:
        json.dump(custom_file, output_file)


# ANY FORMAT CONVERSION
def convert_to_custom_format(file_path, output_path):
    with open(file_path, "rb") as file:
        binary_content = file.read()
        encoded_content = base64.b64encode(binary_content).decode("utf-8")
    
    # Get file extension and metadata
    file_extension = file_path.split(".")[-1]
    custom_file = {
        "metadata": {
            "original_type": file_extension,
            "file_name": os.path.basename(file_path)
        },
        "content": encoded_content
    }
    
    # Save to custom format
    with open(output_path, "w") as output_file:
        json.dump(custom_file, output_file)




#convert_ppt("example.pptx", "output.myext")
#convert_excel("example.xlsx", "output.myext")
convert_to_custom_format(r"C:\Users\HR-IT-MATTHEW-PC\Desktop\Projects\Applications\AntiCopyPaste\try.pptx", "try.p")

print("Current working directory: ", os.getcwd())