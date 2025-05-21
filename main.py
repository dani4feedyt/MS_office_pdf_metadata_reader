import os
from datetime import datetime
from openpyxl import load_workbook
from docx import Document
from tkinter import filedialog
import tkinter as tk

file_path = None

def get_excel_metadata(file_path):
    workbook = load_workbook(file_path, keep_vba=False)
    metadata = {
        'SheetNames': workbook.sheetnames,
        'SheetCount': len(workbook.sheetnames),
        'Author': workbook.properties.creator,
        'Title': workbook.properties.title,
        'Created': workbook.properties.created,
        'Tags': workbook.properties.keywords,
        'Category': workbook.properties.category
    }
    return metadata

def get_word_metadata(file_path):
    doc = Document(file_path)
    props = doc.core_properties
    metadata = {
        'SheetNames': None,
        'SheetCount': None,
        'Author': props.author,
        'Title': props.title,
        'Created': props.created,
        'Tags': props.keywords,
        'Category': props.category
    }
    return metadata

def get_file_metadata(file_path):
    try:
        file_stats = os.stat(file_path)
        file_size = file_stats.st_size
        last_modified = datetime.fromtimestamp(file_stats.st_mtime)

        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.xlsx':
            metadata = get_excel_metadata(file_path)
        elif ext in ['.docx', '.doc']:
            metadata = get_word_metadata(file_path)
        else:
            return f"Unsupported file type: {ext}"

        metadata['FileSize'] = file_size
        metadata['LastModifiedTime'] = last_modified

        return metadata

    except Exception as e:
        return f"Error: {str(e)}"


def change_pic_down(event):
    button = event.widget
    if str(button.winfo_name()) == "!button":
        activated_img = tk.PhotoImage(file='assets/button_choose_A.png')
        button.configure(image=activated_img)
        button.image = activated_img
        open_file()

    elif str(button.winfo_name()) == "!button2":
        activated_img = tk.PhotoImage(file='assets/button_extract_A.png')
        button.configure(image=activated_img)
        button.image = activated_img
        extract_metadata(file_path)


def change_pic_up(event):
    button = event.widget
    if str(button.winfo_name()) == "!button":
        button = event.widget
        deactivated_img = tk.PhotoImage(file='assets/button_choose.png')
        button.configure(image=deactivated_img)
        button.image = deactivated_img
    elif str(button.winfo_name()) == "!button2":
        activated_img = tk.PhotoImage(file='assets/button_extract.png')
        button.configure(image=activated_img)
        button.image = activated_img


def open_file():

    local_file_path = filedialog.askopenfilename(filetypes=[
        ("Supported Files", "*.xlsx *.docx *.doc"),
        ("Excel Files", "*.xlsx"),
        ("Word Files", "*.docx *.doc")
    ])
    global file_path
    file_path = local_file_path
    if file_path:
        text_box.insert(tk.END, f'File <<{file_path}>> loaded successfully!')
    else:
        text_box.insert(tk.END, 'Error, please select a file!')


def extract_metadata(file_path):
    if file_path:
        metadata = get_file_metadata(file_path)
        if isinstance(metadata, dict):
            result = f"File Size: {metadata.get('FileSize')} bytes\n"
            result += f"Last Modified: {metadata.get('LastModifiedTime')}\n"
            result += f"Sheet Names: {metadata.get('SheetNames')}\n"
            result += f"Sheet Count: {metadata.get('SheetCount')}\n"
            result += f"Author: {metadata.get('Author')}\n"
            result += f"Title: {metadata.get('Title')}\n"
            result += f"Created: {metadata.get('Created')}\n"
            result += f"Tags: {metadata.get('Tags')}\n"
            result += f"Category: {metadata.get('Category')}\n"
        else:
            result = metadata

        text_box.delete(1.0, tk.END)
        text_box.insert(tk.END, result)


# GUI Setup
root = tk.Tk()
root.title("Metadata Extraction Tool")
root.geometry("600x500")
root.resizable(False, False)
true_bg = "#222831"
root.config(bg=true_bg)

b_img = tk.PhotoImage(file='assets/button_choose.png')
open_button = tk.Button(root, image=b_img, bg=true_bg, borderwidth=0, relief="solid", activebackground=true_bg)
open_button.image = b_img

open_button.bind("<ButtonPress-1>", change_pic_down)
open_button.bind("<ButtonRelease-1>", change_pic_up)


open_button.pack(pady=10)

text_box = tk.Text(root, height=20, width=70, bg="#393E46", fg="white",
                   font=("Consolas", 10), wrap=tk.WORD)
text_box.pack(pady=10)

b2_img = tk.PhotoImage(file='assets/button_extract.png')

b_extract = tk.Button(root, image=b2_img, bg=true_bg, borderwidth=0, relief="solid", activebackground=true_bg)

b_extract.bind("<ButtonPress-1>", change_pic_down)
b_extract.bind("<ButtonRelease-1>", change_pic_up)

b_extract.image = b2_img
b_extract.pack(pady=10)

root.mainloop()
