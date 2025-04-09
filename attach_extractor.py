import os
import extract_msg
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

dir1 = "Emails"
dir2 = "Attachments"

def cancel_action():
    app.destroy()

def extract_msg_from_msg(msg_file_path, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    with extract_msg.openMsg(msg_file_path) as msg:
        emails = msg.attachments
        for email in emails:
            email.save(customPath = output_dir, extractEmbedded = True)

def extract_statements(input_dir, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    for filename in os.listdir(input_dir):
        if filename.endswith('.msg'):
            path = os.path.join(input_dir, filename)
            with extract_msg.openMsg(path) as msg:
                msg.saveAttachments(customPath=output_dir)
        print(f"Extracted {filename}")

# Error handling
def checkbox_check():
    if emails_var.get() and attachments_var.get():
        messagebox.showerror("Error", "Please select only one option.")
        return 1
    elif not(emails_var.get() or attachments_var.get()):
        messagebox.showerror("Error", "Please select an option.")
        return 1
    else:
        return 0

def extract_action():
    if checkbox_check():
        return 0
    path = file_name.get()
    if not(path.endswith(".msg") or path.endswith(".MSG")):
        path = path + ".msg"
    
    if os.path.exists(path):
        if emails_var.get():
            extract_msg_from_msg(path, dir1)
            extract_statements(dir1, dir2)
            messagebox.showinfo("Action", f"Extracted {path}")
        elif attachments_var.get():
            if not os.path.exists(dir2):
                os.makedirs(dir2)
            with extract_msg.openMsg(path) as msg:
                msg.saveAttachments(customPath=dir2)
            messagebox.showinfo("Action", f"Extracted {path}")
    else:
        messagebox.showerror("Error","File not found")
    
app = tk.Tk()
app.title("Attachment Extractor")
app.geometry('300x170')

label = tk.Label(app, text="Enter source file name:")
label.pack()

file_name = tk.StringVar()
source_file = tk.Entry(app, width=30, textvariable=file_name)
source_file.pack()

label2 = tk.Label(app, text="My file contains:")
label2.pack(pady=(10,0))

chkbx_frame = tk.Frame(app)
chkbx_frame.pack()

emails_var = tk.IntVar()
attachments_var = tk.IntVar()

emails_chbx = tk.Checkbutton(chkbx_frame, text='Emails', variable=emails_var)
emails_chbx.pack(side=tk.LEFT, padx=6)
attachments_chbx = tk.Checkbutton(chkbx_frame, text='Attachments', variable=attachments_var)
attachments_chbx.pack(side=tk.LEFT, padx=6)

cancel_button = tk.Button(app, text="Close", command=cancel_action)
cancel_button.pack(side=tk.BOTTOM, fill='x', pady=3)

extract_button = tk.Button(app, text="Extract", command=extract_action)
extract_button.pack(side=tk.BOTTOM, fill='x', pady=3)

app.mainloop()