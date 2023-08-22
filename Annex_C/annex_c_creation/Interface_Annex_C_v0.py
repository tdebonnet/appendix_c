#!/usr/bin/env python
# coding: utf-8

# In[8]:


import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pickle
import os
import subprocess
from tkinter import ttk

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

root = tk.Tk()
root.title("Creation of Annex C")
root.geometry("400x550")


root.iconbitmap("C:/Users/u119708/Documents/Untitled Folder/icone.ico")  # Remplacez par le chemin de votre logo au format .ico

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

heading_label = tk.Label(content_frame, text="Creation of Annex C", font=title_font, fg=text_color, bg="white")
heading_label.pack(pady=10)

file_var = tk.StringVar()
file_label = tk.Label(content_frame, text="Select a File:", font=("Segoe UI", 14), fg=text_color, bg=bg_color)
file_label.pack(anchor="w", pady=(10, 0))
#selected_file_label = tk.Label(content_frame, textvariable=file_var, width=30, bg="white")#, fg=text_color, bg="white")
#selected_file_label.pack(anchor="w", pady=(5, 10))
save_entry = tk.Entry(content_frame, textvariable=file_var, width=35)
save_entry.pack(anchor="w", pady=(5, 10))
browse_button = tk.Button(content_frame, text="Browse", command=browse_file, bg=accent_color, fg=text_color)
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

select_save_button = tk.Button(content_frame, text="Select Folder", command=select_save_path, bg=accent_color, fg=text_color)
select_save_button.pack(anchor="w", pady=(5, 0))

def run_annex_c_script():
    try:
        selected_option = option_var.get()
        if selected_option == "SimaPro":
            subprocess.run(["python", "annex_c_simapro.py", file_name_saved], shell=True, check=True)
        elif selected_option == "OpenLCA":
            subprocess.run(["python", "annex_c_openlca.py", file_name_saved], shell=True, check=True)
        elif selected_option == "Brightway":
            subprocess.run(["python", "annex_c_brightway.py", file_name_saved], shell=True, check=True)
        messagebox.showinfo("Script over", "Annex C successfully created.")
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
create_button = tk.Button(root, text="Create", command=create_and_run, bg=accent_color, fg=text_color, width=20)
create_button.pack(anchor="n", pady=(0, 1), padx=(0, 10))  # Ajuste la marge inférieure et ajoute un peu d'espace horizontal
root.mainloop()


# import tkinter as tk
# from tkinter import filedialog
# from tkinter import messagebox
# import pickle
# import os
# import subprocess
# from tkinter import ttk
# from ttkthemes import ThemedTk
# 
# def browse_file():
#     file_path = filedialog.askopenfilename()
#     with open('file_path.pickle', 'wb') as file:
#         pickle.dump(file_path, file)
#         
#     file_var.set(file_path)
# 
# def select_save_path():
#     save_path = filedialog.askdirectory()
#     with open('save_path.pickle', 'wb') as file:
#         pickle.dump(save_path, file)
# 
#     save_var.set(save_path)
# 
# def create_file():
#     selected_option = option_var.get()
#     save_path = save_var.get()
#     # Add your logic here to create the file using the provided information
#     
# def on_checkbox_click(checkbox_var):
#     option_a_var.set(0)
#     option_b_var.set(0)
#     option_c_var.set(0)
#     checkbox_var.set(1)
# 
# # Create the main window
# root = tk.Tk()
# root.title("Creation of Annex C")
# root.geometry("400x550")  # Set the window size
# root.configure(bg="#34495E")  # Background color of the window
# 
# # Colors for the elements
# button_color = "#3498DB"
# text_color = "#ECF0F1"
# label_bg_color = "#2C3E50"
# 
# # Create a frame for the content
# content_frame = tk.Frame(root, bg="#34495E")
# content_frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)
# 
# # Heading
# heading_label = tk.Label(content_frame, text="Creation of Annex C", font=("Helvetica", 18, "bold"), fg=text_color, bg=label_bg_color)
# heading_label.pack(pady=10)
# 
# # File Selection
# file_var = tk.StringVar()
# file_label = tk.Label(content_frame, text="Select a File:", font=("Helvetica", 14), fg=text_color, bg="#34495E")
# file_label.pack(anchor="w", pady=(10, 0))
# selected_file_label = tk.Label(content_frame, textvariable=file_var, width=30)#font=("Helvetica", 10), fg=text_color, bg="#34495E")
# selected_file_label.pack(anchor="w", pady=(5, 10))
# browse_button = tk.Button(content_frame, text="Browse", command=browse_file, bg=button_color, fg=text_color)
# browse_button.pack(anchor="w")
# 
# 
# 
# # Options
# options_label = tk.Label(content_frame, text="Select an app:", font=("Helvetica", 14), fg=text_color, bg="#34495E")
# options_label.pack(anchor="w", pady=(0, 0))
# """option_a_var = tk.IntVar()
# option_b_var = tk.IntVar()
# option_c_var = tk.IntVar()
# 
# options_label = tk.Label(content_frame, text="Select an app:", font=("Helvetica", 12), fg=text_color, bg="#34495E")
# options_label.pack(anchor="w", pady=(0, 0))
# option_a_checkbox = tk.Checkbutton(content_frame, text="SimaPro", variable=option_a_var, command=lambda: on_checkbox_click(option_a_var), font=("Helvetica", 12), fg=text_color, bg="#34495E")
# option_b_checkbox = tk.Checkbutton(content_frame, text="OpenLCA", variable=option_b_var, command=lambda: on_checkbox_click(option_b_var), font=("Helvetica", 12), fg=text_color, bg="#34495E")
# option_c_checkbox = tk.Checkbutton(content_frame, text="Brightway", variable=option_c_var, command=lambda: on_checkbox_click(option_c_var), font=("Helvetica", 12), fg=text_color, bg="#34495E")
# option_a_checkbox.pack(anchor="w")
# option_b_checkbox.pack(anchor="w")
# option_c_checkbox.pack(anchor="w")
# """
# def on_checkbox_click(checkbox_var):
#     option_a_var.set(0)
#     option_b_var.set(0)
#     option_c_var.set(0)
#     checkbox_var.set(1)
# 
# #Fonction de gestion des cases à cocher
# def on_checkbox_click(var):
#     print(var.get())  # Vous pouvez mettre ici le code que vous souhaitez exécuter
# # Couleurs et police
# text_color = "white"
# bg_color = "#34495E"
# font_style = ("Helvetica", 12)
# # Créez le style pour les Radiobuttons
# style = ttk.Style()
# style.configure("Custom.TRadiobutton", background=bg_color, foreground=text_color, font=font_style)
# 
# # Créez les Radiobuttons avec le même format que le reste de l'interface
# option_a_radiobutton = ttk.Radiobutton(content_frame, text="SimaPro", variable=selected_option, value="SimaPro", command=lambda: on_radiobutton_select(selected_option), style="Custom.TRadiobutton")
# option_b_radiobutton = ttk.Radiobutton(content_frame, text="OpenLCA", variable=selected_option, value="OpenLCA", command=lambda: on_radiobutton_select(selected_option), style="Custom.TRadiobutton")
# option_c_radiobutton = ttk.Radiobutton(content_frame, text="Brightway", variable=selected_option, value="Brightway", command=lambda: on_radiobutton_select(selected_option), style="Custom.TRadiobutton")
# 
# # Placez les Radiobuttons dans le frame de contenu
# option_a_radiobutton.pack(pady=5, anchor=tk.W)
# option_b_radiobutton.pack(pady=5, anchor=tk.W)
# option_c_radiobutton.pack(pady=5, anchor=tk.W)
# 
# 
# # File name
# root.title("File name")
# 
# label = tk.Label(content_frame, text="Enter file name:",font=("Helvetica", 14), fg=text_color, bg="#34495E")
# label.pack(anchor="w")
# # Champ de saisie pour entrer le nom de fichier
# entry = tk.Entry(content_frame, font=("Helvetica", 12), fg=text_color, bg="#34495E")
# entry.pack(anchor="w")
# 
# def select_name():
#     global file_name_saved  # Ajoutez cette ligne pour indiquer que vous utilisez la variable globale
#     file_name_saved = entry.get()  # Récupérer le nom de fichier depuis le champ de saisie
#     if file_name_saved:
#         # Ajouter l'extension .xlsx si nécessaire
#         if not file_name_saved.endswith('.xlsx'):
#             file_name_saved += '.xlsx'
#         # Sauvegarder le nom de fichier choisi dans un fichier pickle
#         with open('file_name_saved.pickle', 'wb') as file:
#             pickle.dump(file_name_saved, file)
#         file_name_saved_var.set(file_name_saved)
# 
# # Save Path
# save_var = tk.StringVar()
# save_label = tk.Label(content_frame, text="Save Path:", font=("Helvetica", 14), fg=text_color, bg="#34495E")
# save_label.pack(anchor="w", pady=(10, 0))
# save_entry = tk.Entry(content_frame, textvariable=save_var, width=35)
# save_entry.pack(anchor="w", pady=(5, 10))
# 
# select_save_button = tk.Button(content_frame, text="Select Folder", command=select_save_path, bg=button_color, fg=text_color)
# select_save_button.pack(anchor="w", pady=(5, 0))
# 
# 
# #Associate the script
# # Créez une variable pour stocker le nom de fichier choisi
# file_name_saved_var = tk.StringVar()
# # Fonction pour exécuter le script annex_c_simapro.py
# def run_annex_c_script():
#     try:
#         if option_a_var.get():
#             subprocess.run(["python", "annex_c_simapro.py", file_name_saved], shell=True, check=True)
#             #run_annex_c_script("annex_c_simapro.py")
#         if option_b_var.get():
#             subprocess.run(["python", "annex_c_simapro.py", file_name_saved], shell=True, check=True)
#             #run_annex_c_script("annex_c_openlca.py")
#         if option_c_var.get():
#             subprocess.run(["python", "Annex_C_Brightway_interface.py", file_name_saved], shell=True, check=True)
#         #run_annex_c_script("annex_c_brightway.py")
#         # Le script s'est terminé sans erreur
#         messagebox.showinfo("Script terminé", "Le script annex_c_simapro.py s'est terminé.")
#     except subprocess.CalledProcessError:
#         # Le script a rencontré une erreur
#         messagebox.showerror("Error", "An error occurred while running the script.")
# 
# # Fonction pour créer le fichier Annex C et exécuter le script
# def create_and_run():
#     global file_name_saved
#     file_name_saved = entry.get()
#     if file_name_saved:
#         if not file_name_saved.endswith('.xlsx'):
#             file_name_saved += '.xlsx'
#         with open('file_name_saved.pickle', 'wb') as file:
#             pickle.dump(file_name_saved, file)
#         file_name_saved_var.set(file_name_saved)
#         run_annex_c_script()
# 
# 
# create_button = tk.Button(root, text="Create", command=create_and_run, bg=button_color, fg=text_color, width=20)
# create_button.pack(anchor="n", pady=(0, 1), padx=(0, 10))  # Ajuste la marge inférieure et ajoute un peu d'espace horizontal
# 
# # Start the main loop
# root.mainloop()
# 

# def run_annex_c_script():
#     try:
#         subprocess.run(["python", "annex_c_simapro.py"], shell=True, check=True)
#         # Le script s'est terminé sans erreur
#         messagebox.showinfo("Script terminé", "Le script annex_c_simapro.py s'est terminé.")
#     except subprocess.CalledProcessError:
#         # Le script a rencontré une erreur
#         messagebox.showerror("Error", "An error occurred while running the script.")

# In[ ]:




