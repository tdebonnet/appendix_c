#!/usr/bin/env python
# coding: utf-8

# In[59]:


import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pickle
import os
import subprocess
from tkinter import ttk
from PIL import Image, ImageTk

def browse_file():
    file_path = filedialog.askopenfilename()
    with open('file_path.pickle', 'wb') as file:
        pickle.dump(file_path, file)
        
    file_var.set(file_path)

def select_save_path():
    save_path = filedialog.askdirectory()
    with open('save_path.pickle', 'wb') as file:
        pickle.dump(save_path, file)

    save_var.set(save_path)

def create_file():
    selected_option = option_var.get()
    save_path = save_var.get()
    # Ajoutez votre logique ici pour créer le fichier en utilisant les informations fournies
    
def on_checkbox_click(checkbox_var):
    option_a_var.set(0)
    option_b_var.set(0)
    option_c_var.set(0)
    checkbox_var.set(1)
    
def center_window(window):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    window.update_idletasks()

    window_width = window.winfo_reqwidth()
    window_height = window.winfo_reqheight()

    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2-20

    window.geometry(f"{window_width}x{window_height}+{x}+{y}")

root = tk.Tk()
root.title("Creation of Appendix C")
root.geometry("400x550")

#root.iconbitmap("C:/Users/u119708/Documents/Untitled Folder/icone.ico")  # Remplacez par le chemin de votre logo au format .ico
###MENU help
def show_instructions(option):
    instructions = {
        "SimaPro": {"text":"SimaPro export file must be in .csv extension and the export must be done with these following parameters :","image_path": r"C:\Users\u119708\Documents\annex_C_dvpt\image.png"},
        "openLCA": {
            "text": "OpenLCA export file must be in .zip extension."
            #"image_path": None  # Pas d'image pour cette option
        },
        "Brightway": {
            "text": "Brightway export file must be in .xlsx extension."
            #"image_path": None  # Pas d'image pour cette option
        }
    }
    
    if option in instructions:
        instruction_info = instructions[option]
        instruction_text = instruction_info["text"]
        
        if "image_path" in instruction_info:
            
            image_path = instruction_info["image_path"]
            image = Image.open(image_path)
            image = image.resize((500, 500), resample=Image.LANCZOS)  # Redimensionner l'image
            photo = ImageTk.PhotoImage(image)
            
            # Créer une fenêtre Tkinter pour afficher les instructions avec l'image
            window = tk.Toplevel()
            window.title(option)
            
            text_label = tk.Label(window, text=instruction_text)
            image_label = tk.Label(window, image=photo)
            
            text_label.pack(padx=10, pady=10)
            image_label.pack(padx=10, pady=10)
            
            # Nécessaire pour éviter la suppression de la référence à la photo
            image_label.photo = photo
            
        else:
            messagebox.showinfo("Instructions", instruction_text)
    else:
        messagebox.showwarning("Erreur", "Option non reconnue.")
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

help_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Help", menu=help_menu)

options = ["SimaPro", "openLCA", "Brightway"]
for option in options:
    help_menu.add_command(label=option, command=lambda o=option: show_instructions(o))

############### menu help
# Couleurs
bg_color = "white"
text_color = "black"
accent_color = "#00A1C0"
button_color = "#0BAC43"
label_bg_color = "#B91E32"

root.configure(bg=bg_color)

content_frame = tk.Frame(root, bg=bg_color)
content_frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)

# Utilisation de la police moderne "Segoe UI"
title_font = ("Segoe UI", 18, "bold")

heading_label = tk.Label(content_frame, text="Creation of Appendix C", font=title_font, fg=text_color, bg="white")
heading_label.pack(pady=10)

file_var = tk.StringVar()
file_label = tk.Label(content_frame, text="Select a File:", font=("Segoe UI", 14), fg=text_color, bg=bg_color)
file_label.pack(anchor="w", pady=(10, 0))
#selected_file_label = tk.Label(content_frame, textvariable=file_var, width=30, bg="white")#, fg=text_color, bg="white")
#selected_file_label.pack(anchor="w", pady=(5, 10))
save_entry = tk.Entry(content_frame, textvariable=file_var, width=35)
save_entry.pack(anchor="w", pady=(5, 10))
browse_button = tk.Button(content_frame, text="Browse", command=browse_file, font=("Segoe UI", 10),bg=accent_color, fg="white", width=15)
browse_button.pack(anchor="w")

option_var = tk.StringVar()

options_label = tk.Label(content_frame, text="Select an app:", font=("Segoe UI", 14), fg=text_color, bg=bg_color)
options_label.pack(anchor="w", pady=(0, 0))

style = ttk.Style()
style.configure("Custom.TRadiobutton", background=bg_color, foreground=text_color, font=("Segoe UI", 12))

option_a_radiobutton = ttk.Radiobutton(content_frame, text="SimaPro", variable=option_var, value="SimaPro", command=lambda: on_checkbox_click(option_var), style="Custom.TRadiobutton")
option_b_radiobutton = ttk.Radiobutton(content_frame, text="OpenLCA", variable=option_var, value="OpenLCA", command=lambda: on_checkbox_click(option_var), style="Custom.TRadiobutton")
option_c_radiobutton = ttk.Radiobutton(content_frame, text="Brightway", variable=option_var, value="Brightway", command=lambda: on_checkbox_click(option_var), style="Custom.TRadiobutton")

option_a_radiobutton.pack(pady=5, anchor=tk.W)
option_b_radiobutton.pack(pady=5, anchor=tk.W)
option_c_radiobutton.pack(pady=5, anchor=tk.W)



label = tk.Label(content_frame, text="Enter file name:", font=("Segoe UI", 14), fg=text_color, bg=bg_color)
label.pack(anchor="w")
entry = tk.Entry(content_frame, font=("Segoe UI", 12), fg=text_color, bg=bg_color)
entry.pack(anchor="w")

file_name_saved_var = tk.StringVar()

save_var = tk.StringVar()
save_label = tk.Label(content_frame, text="Save Path:", font=("Segoe UI", 14), fg=text_color, bg=bg_color)
save_label.pack(anchor="w", pady=(10, 0))
save_entry = tk.Entry(content_frame, textvariable=save_var, width=35)
save_entry.pack(anchor="w", pady=(5, 10))

select_save_button = tk.Button(content_frame, text="Select Folder", command=select_save_path,  font=("Segoe UI", 10),bg=accent_color, fg="white", width=15)
select_save_button.pack(anchor="w", pady=(5, 0))

def run_annex_c_script():
    try:
        selected_option = option_var.get()
        if selected_option == "SimaPro":
            subprocess.run(["python", "annex_c_simapro.py", file_name_saved], shell=True, check=True)
        elif selected_option == "OpenLCA":
            subprocess.run(["python", "annex_c_olca.py", file_name_saved], shell=True, check=True)
        elif selected_option == "Brightway":
            subprocess.run(["python", "annex_c_brightway.py", file_name_saved], shell=True, check=True)
        messagebox.showinfo("Script over", "Appendix C successfully created.")
    except subprocess.CalledProcessError:
        messagebox.showerror("Error", "An error occurred while executing the script.")

def create_and_run():
    global file_name_saved
    file_name_saved = entry.get()
    if file_name_saved:
        if not file_name_saved.endswith('.xlsx'):
            file_name_saved += '.xlsx'
        with open('file_name_saved.pickle', 'wb') as file:
            pickle.dump(file_name_saved, file)
        file_name_saved_var.set(file_name_saved)
        run_annex_c_script()
create_button = tk.Button(root, text="Create", command=create_and_run, font=("Segoe UI", 11), bg=accent_color, fg="white", width=20)
create_button.pack(anchor="n", pady=(0, 1), padx=(0, 10))  # Ajuste la marge inférieure et ajoute un peu d'espace horizontal
center_window(root)
root.mainloop()


# In[ ]:




