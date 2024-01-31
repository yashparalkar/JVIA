import customtkinter as ctk
import webbrowser
from PIL import Image
from main import scrape_data
import tkinter as tk

# Declare status_label as a global variable
global status_label

def run_scraping():
    codes = [int(code) for code in code_entry.get().split()]  # Convert space-separated codes to a list of integers
    file_name = file_name_entry.get()
    data_t = data_type_var.get()
    status_label.configure(text="Data Extraction in process, please wait")
    scrape_data(codes, file_name + ".xlsx", data_t)

def open_claude():
    webbrowser.open("https://claude.ai/chats")
app = ctk.CTk(fg_color="gray10")
app.title("JVIA")
app.iconbitmap("Source/Icon1.ico")

logo_img = Image.open("Source/Icon1.png")
logo_photo = ctk.CTkImage(logo_img, size=(80,80))

# Create a frame for the tab buttons
tab_buttons_frame = ctk.CTkFrame(app)
tab_buttons_frame.place(x=1, y=1)

# Create a button for the first tab
data_extractor_button = ctk.CTkButton(tab_buttons_frame, text="Data Extractor", command=lambda: show_frame(data_extractor_frame, data_extractor_button)
                                      , corner_radius=0, width=225)
data_extractor_button.pack(side='left')

# Create a button for the second tab
gpt_tab_button = ctk.CTkButton(tab_buttons_frame, text="Claude AI", command=lambda: show_frame(claude_frame, gpt_tab_button),
                               corner_radius=0, width=225)
gpt_tab_button.pack(side='left')

logo = ctk.CTkLabel(app, image=logo_photo, text='')
logo.place(x=180, y=50)

# Function to show a frame
def show_frame(frame, button):
    data_extractor_frame.place_forget()
    claude_frame.place_forget()
    data_extractor_button.configure(fg_color="#b81d1d")
    gpt_tab_button.configure(fg_color="#b81d1d")
    frame.place(x=40, y=140)
    button.configure(fg_color="dark red")

# Create a frame for the existing elements
data_extractor_frame = ctk.CTkFrame(app, width=300, height=400, fg_color='gray10')

# Add your existing elements to data_extractor_frame here...
title_label = ctk.CTkLabel(data_extractor_frame, text="Export/Import Data Extractor", font=("", 25), text_color="#ffffff")
title_label.grid(row=1, column=0, columnspan=2, sticky='ew', pady=20)

# Code entry
code_label = ctk.CTkLabel(data_extractor_frame, text="Enter Codes (space-separated):", text_color="#ffffff")
code_label.grid(row=2, column=0, padx=10, pady=10)
code_entry = ctk.CTkEntry(data_extractor_frame)
code_entry.grid(row=2, column=1, padx=10, pady=10)

# File name entry
file_name_label = ctk.CTkLabel(data_extractor_frame, text="Enter File Name:", text_color="#ffffff")
file_name_label.grid(row=3, column=0, padx=10, pady=10)
file_name_entry = ctk.CTkEntry(data_extractor_frame)
file_name_entry.grid(row=3, column=1, padx=10, pady=10)

# Data type selection
data_type_label = ctk.CTkLabel(data_extractor_frame, text="Select Data Type:", text_color="#ffffff")
data_type_label.grid(row=4, column=0, padx=10, pady=10)
data_type_var = tk.StringVar(value="export")
data_type_combobox = ctk.CTkComboBox(data_extractor_frame, variable=data_type_var, values=["import", "export"])
data_type_combobox.grid(row=4, column=1, padx=10, pady=10)

# Run button
run_button = ctk.CTkButton(data_extractor_frame, text="Get Data", command=run_scraping, fg_color="#b81d1d", corner_radius=15)
run_button.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

status_label = ctk.CTkLabel(data_extractor_frame, text="Status:", text_color="#ffffff")
status_label.grid(row=6, column=0, columnspan=2, padx=10, pady=10)

# Create a frame for the new tab
claude_frame = ctk.CTkFrame(app, width=300, height=400, fg_color='gray10')
claude_button = ctk.CTkButton(claude_frame, text='Open Claude', command=open_claude, corner_radius=15, fg_color="#b81d1d")
claude_button.grid(row=1, column=2, columnspan=2, padx=10, pady=10)

# Add your new tab content to claude_frame here...

# Show the first tab by default
show_frame(data_extractor_frame, data_extractor_button)

app.geometry("450x550")
# Start GUI loop
app.mainloop()
