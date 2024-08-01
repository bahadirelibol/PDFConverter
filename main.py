from pdf2docx import Converter

import tkinter as tk
from tkinter import filedialog, messagebox

import os


pdf_file = None
#pdf_file is initially defined as None, i.e. no file is initially selected.


def select_file():
    global pdf_file
    pdf_file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf_file:
        selected_label.config(text=f"Selected: {os.path.basename(pdf_file)}", fg="#98FF98")
        status_label.config(text="")


def convert_pdf_to_docx():
    if pdf_file is None:
        messagebox.showwarning("Warning", "No PDF file selected. Please select a PDF file first.")
        return

    docx_file = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("DOCX files", "*.docx")])
    if not docx_file:
        return

    status_label.config(text="Converting...", fg="#FFA500")
    root.update()

    try:
        cv = Converter(pdf_file)
        cv.convert(docx_file)
        cv.close()
        status_label.config(text=f"Converted: {os.path.basename(docx_file)}", fg="#98FF98")
    except Exception as e:
        status_label.config(text=f"Error: {e}", fg="red")
    finally:
        root.update()


def close_program():
    if messagebox.askyesno("Exit", "Are you sure you want to exit?"):
        root.destroy()


# Tkinter Window, Widgets, Creating Buttons and Labels
root = tk.Tk()
root.config(background="#CD5C5C")
root.title("PDF to DOCX Converter")
root.geometry("800x600")

button_style = {
    "bg": "#5dade2",
    "fg": "white",
    "font": ("Arial", 8, "bold"),
    "padx": 20,
    "pady": 10
}

select_button = tk.Button(root, text="Select File", command=select_file, **button_style)
select_button.place(x=50, y=50)

selected_label = tk.Label(root, text="", bg="#CD5C5C")
selected_label.place(x=50, y=150)

convert_button = tk.Button(root, text="Convert PDF to DOCX and Save", command=convert_pdf_to_docx, **button_style)
convert_button.place(x=500, y=50)

status_label = tk.Label(root, text="", bg="#CD5C5C")
status_label.place(x=500, y=150)

exit_button = tk.Button(root, text="Exit", command=close_program, **button_style)
exit_button.pack(padx=20, pady=20, side='right', anchor='s')

root.mainloop()
