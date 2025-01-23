import os
import tkinter as tk
from tkinter import filedialog

def open_file():
  """Opens a file dialog and prints the selected file path."""
  filepath = filedialog.askopenfilename(
    initialdir="/",
    title="Select a file",
    filetypes=(("all files", "*.*"), ("text files", "*.txt"))
  )
  if filepath:
    with open(filepath, 'r') as f:
      file_contents = f.read()
      print(f"Selected file: {filepath}")
      print(f"File contents: {file_contents}")

root = tk.Tk()
root.title("File Selector")

button = tk.Button(root, text="Select File", command=open_file)
button.pack(pady=20)

root.mainloop()