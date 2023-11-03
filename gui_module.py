# gui_module.py
import tkinter as tk
from tkinter import font  # Import the font module

class HereApp:
    def __init__(self, window, action_on_input, doc_name):
        self.window = window
        self.window.title("Here App")
        self.window.geometry("300x250")  # Set window size
        
        self.action_on_input = action_on_input
        
        # Configure rows and columns to expand
        self.window.grid_rowconfigure(0, weight=1)  # This will push row 1 to the center
        self.window.grid_rowconfigure(2, weight=1)  # This will push row 1 to the center
        self.window.grid_rowconfigure(3, weight=1)  # This will push row 1 to the center
        self.window.grid_columnconfigure(0, weight=1)  # This will push columns 0 and 1 to the center
        self.window.grid_columnconfigure(1, weight=1)  # This will push columns 0 and 1 to the center
        
        # Remove the .docx extension from the document name
        doc_name_no_extension = doc_name.replace('.docx', '')
        
        self.doc_label = tk.Label(self.window, text=doc_name_no_extension)  # Label to hold the document name
        self.doc_label.grid(row=0, column=0, columnspan=2, sticky=tk.NSEW)  # Position the document name label
        
        # Set a larger font size for the label
        label_font = font.Font(size=20)  # Set font size to 20 
        self.label = tk.Label(self.window, font=label_font, wraplength=150)  # Set wraplength to 200 pixels
        self.label.grid(row=1, column=0, columnspan=2, sticky=tk.NSEW)  # Center the label
        
        self.one_button = tk.Button(self.window, text="Here", command=lambda: self.on_button(1))  # Text changed to "Here"
        self.one_button.grid(row=2, column=0, sticky=tk.NSEW, padx=8, pady=8)  # Stretch button
        
        self.zero_button = tk.Button(self.window, text="Absent", command=lambda: self.on_button(0))  # Text changed to "Absent"
        self.zero_button.grid(row=2, column=1, sticky=tk.NSEW, padx=8, pady=8)  # Stretch button
        
    def on_button(self, value):
        self.action_on_input(value, self.label)

    def update_label(self, text):
        self.label.config(text=text)

    def run(self):
        self.window.mainloop()

def create_app(action_on_input, doc_name):  # Updated function signature to include doc_name
    window = tk.Tk()
    app = HereApp(window, action_on_input, doc_name)  # Pass doc_name to HereApp
    return app
