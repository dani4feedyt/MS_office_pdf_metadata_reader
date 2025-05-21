import os
from datetime import datetime
from openpyxl import load_workbook
from docx import Document
from tkinter import filedialog
import tkinter as tk

b_down = False
jobid = None

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
    global b_down
    if not b_down:
        button = event.widget
        activated_img = tk.PhotoImage(file='assets/button_choose_A.png')
        button.configure(image=activated_img)
        button.image = activated_img

        b_down = True
        print("downnn")
        open_file()


def change_pic_up(event):
    global b_down
    if b_down:

        print("upnn")
        button = event.widget
        deactivated_img = tk.PhotoImage(file='assets/button_choose.png')
        button.configure(image=deactivated_img)
        button.image = deactivated_img
        b_down = False


def open_file():

    file_path = filedialog.askopenfilename(filetypes=[
        ("Supported Files", "*.xlsx *.docx *.doc"),
        ("Excel Files", "*.xlsx"),
        ("Word Files", "*.docx *.doc")
    ])
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
b_img_a = tk.PhotoImage(file='assets/button_choose_A.png')
open_button = tk.Button(root, image=b_img, bg=true_bg, borderwidth=0, relief="solid", activebackground=true_bg)
open_button.image = b_img
#open_button.configure(command=lambda: change_pic_down(open_button))
open_button.bind("<ButtonPress-1>", change_pic_down)
open_button.bind("<ButtonRelease-1>", change_pic_up)


open_button.pack(pady=10)

text_box = tk.Text(root, height=20, width=70, bg="#393E46", fg="white",
                   font=("Consolas", 10), wrap=tk.WORD)
text_box.pack(pady=10)

exit_button = tk.Button(root, text="Exit", command=root.quit,
                        bg="#DFD0B8", fg="#222831", font=("Arial", 12, "bold"), activebackground="#948979")
exit_button.pack(pady=10)

root.mainloop()
