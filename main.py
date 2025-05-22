import os
from pypdf import PdfReader
from openpyxl import load_workbook
from docx import Document
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import tkinter as tk
import re

file_path = None

def get_pdf_metadata(file_path):
    reader = PdfReader(file_path)
    content = reader.metadata
    print(content)
    metadata = {
        'FileType': 'PDF document',
        'Created': content.creation_date,
        'Author': content.creator,
        'LastModified': content.modification_date,
        'LastModifiedBy': content.producer_raw,#############
        'Title': content.title,
        'Category': content.subject,
        'Tags': content.keywords
    }
    return metadata



def get_excel_metadata(file_path):
    workbook = load_workbook(file_path, keep_vba=False)
    print(workbook.properties)
    metadata = {
        'FileType': 'Excel spreadsheet',
        'SheetCount': len(workbook.sheetnames),
        'SheetNames': workbook.sheetnames,
        'Created': workbook.properties.created,
        'Author': workbook.properties.creator,
        'LastModified': workbook.properties.modified,
        'LastModifiedBy': workbook.properties.last_modified_by,
        'Title': workbook.properties.title,
        'Category': workbook.properties.category,
        'Tags': workbook.properties.keywords
    }
    return metadata

def get_word_metadata(file_path):
    doc = Document(file_path)
    props = doc.core_properties
    print(props)
    metadata = {
        'FileType': 'Word document',
        'Created': props.created,
        'Author': props.author,
        'LastModified': props.modified,
        'LastModifiedBy': props.last_modified_by,
        'Title': props.title,
        'Category': props.category,
        'Tags': props.keywords
    }
    return metadata

def get_file_metadata(file_path):
    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.pdf':
            metadata = get_pdf_metadata(file_path)
        elif ext in ['.xlsx', '.xls']:
            metadata = get_excel_metadata(file_path)
        elif ext in ['.docx', '.doc']:
            metadata = get_word_metadata(file_path)
        else:
            return f'Unsupported file type: {ext}'

        metadata['FileSize'] = os.stat(file_path).st_size
        metadata['FileName'] = os.path.basename(file_path)
        metadata['FilePath'] = file_path
        metadata['FileExtension'] = ext

        print(metadata)

        return metadata

    except Exception as e:
        return f'Error: {str(e)}'


def change_pic_down(event):
    button = event.widget
    if str(button.winfo_name()) == '!button':
        activated_img = tk.PhotoImage(file='assets/button_choose_A.png')
        button.configure(image=activated_img)
        button.image = activated_img
        open_file()

    elif str(button.winfo_name()) == '!button2':
        activated_img = tk.PhotoImage(file='assets/button_extract_A.png')
        button.configure(image=activated_img)
        button.image = activated_img
        extract_metadata(file_path)


def change_pic_up(event):
    button = event.widget
    if str(button.winfo_name()) == '!button':
        button = event.widget
        deactivated_img = tk.PhotoImage(file='assets/button_choose.png')
        button.configure(image=deactivated_img)
        button.image = deactivated_img
    elif str(button.winfo_name()) == '!button2':
        activated_img = tk.PhotoImage(file='assets/button_extract.png')
        button.configure(image=activated_img)
        button.image = activated_img


def open_file(called=False):
    global file_path
    local_file_path = ''
    text_box.config(state='normal')

    if called:
        if not text_box.get(0.0, tk.END).strip("\n").endswith(('.pdf', '.doc', '.docx', '.xls', '.xlsx')):
            text_box.insert(tk.END, 'Error, unsupported file type. Supported file types: .pdf,.xlsx,.docx')
            text_box.config(state='disabled')
            file_path = ''
            return

        if text_box.get(0.0, tk.END).startswith(("C:", "D:")):
            local_file_path = text_box.get(0.0, tk.END).strip('\n')
    else:
        text_box.delete(0.0, tk.END)
        local_file_path = filedialog.askopenfilename(filetypes=[
            ('Supported Files', '*.xlsx *.xls *.docx *.doc *.pdf'),
            ('Excel Files', '*.xlsx *.xls'),
            ('Word Files', '*.docx *.doc'),
            ('PDF Files', '*.pdf')
        ])

    file_path = local_file_path
    text_box.delete(0.0, tk.END)
    if file_path:
        text_box.insert(tk.END, f'File <{os.path.basename(file_path)}> loaded successfully!')
        print(file_path)
    else:
        text_box.insert(tk.END, 'Error, please select a file!')
    text_box.config(state='disabled')


def handle_drop(event):
    text_box.config(state='normal')
    text_box.delete(0.0, tk.END)
    text_box.insert(tk.END, re.sub(r'[{}]', '', f'{event.data}\n'))
    text_box.config(state='disabled')
    open_file(True)
    return event.action

def extract_metadata(file_path):
    text_box.config(state='normal')
    if file_path:
        result = 'Error reading file metadata'
        metadata = get_file_metadata(file_path)
        f_ext = metadata.get('FileExtension')
        if isinstance(metadata, dict):
            result = f"File Type: {metadata.get('FileType')}\n"
            result += f"File Name: {metadata.get('FileName')}\n"
            result += f"File Extension: {f_ext}\n"
            result += f"File Size: {metadata.get('FileSize')} bytes\n"
            result += f"File Path: {metadata.get('FilePath')}\n"
            if f_ext == '.xlsx':
                result += f"Sheet Count: {metadata.get('SheetCount')}\n"
                result += f"Sheet Names: {metadata.get('SheetNames')}\n"
            result += f"Created: {metadata.get('Created')}\n"
            result += f"Author: {metadata.get('Author')}\n"
            result += f"Last Modified: {metadata.get('LastModified')}\n"
            result += f"Last Modified By: {metadata.get('LastModifiedBy')}\n"
            result += f"Title: {metadata.get('Title')}\n"
            result += f"Category: {metadata.get('Category')}\n"
            result += f"Tags: {metadata.get('Tags')}\n"

        result_f = ''
        for line in result.splitlines():
            if line.endswith(' '):
                line += 'None'
            result_f += f'{line}\n'

        text_box.delete(0.0, tk.END)
        text_box.insert(tk.END, result_f)
        text_box.config(state='disabled')
    else:
        text_box.delete(0.0, tk.END)
        text_box.insert(tk.END, 'Error, filepath corrupted!')
        text_box.config(state='disabled')


# GUI Setup
root = TkinterDnD.Tk()
root.title('Metadata Extraction Tool')
root.geometry('600x500')
root.resizable(False, False)
true_bg = '#222831'
root.config(bg=true_bg)

b_img = tk.PhotoImage(file='assets/button_choose.png')
open_button = tk.Button(root, image=b_img, bg=true_bg, borderwidth=0, relief='solid', activebackground=true_bg)
open_button.image = b_img

open_button.bind('<ButtonPress-1>', change_pic_down)
open_button.bind('<ButtonRelease-1>', change_pic_up)

canvas = tk.Canvas(root, width=600, height=200, borderwidth=0, highlightthickness=0)
open_button.pack(pady=10)
text_box = tk.Text(canvas, height=20, width=70, bg='#393E46', fg='white', font=('Consolas', 10), wrap=tk.WORD, state='disabled')
canvas.create_window(0, 0, window=text_box)
canvas.pack()
canvas.bind("<Button-1>", root.focus())
text_box.pack()

canvas.drop_target_register(DND_FILES)
canvas.dnd_bind('<<Drop>>', handle_drop)

b2_img = tk.PhotoImage(file='assets/button_extract.png')

b_extract = tk.Button(root, image=b2_img, bg=true_bg, borderwidth=0, relief='solid', activebackground=true_bg)

b_extract.bind('<ButtonPress-1>', change_pic_down)
b_extract.bind('<ButtonRelease-1>', change_pic_up)

b_extract.image = b2_img
b_extract.pack(pady=10)

root.mainloop()
