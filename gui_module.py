# gui_module.py
import tkinter as tk

class HereApp:
    def __init__(self, window, action_on_input):
        self.window = window
        self.window.title("Here App")
        self.window.geometry("400x400")  # Set window size
        
        self.action_on_input = action_on_input
        
        # Configure rows and columns to expand
        self.window.grid_rowconfigure(0, weight=1)  # This will push row 1 to the center
        self.window.grid_rowconfigure(2, weight=1)  # This will push row 1 to the center
        self.window.grid_columnconfigure(0, weight=1)  # This will push columns 0 and 1 to the center
        self.window.grid_columnconfigure(1, weight=1)  # This will push columns 0 and 1 to the center
        
        self.label = tk.Label(self.window, width=20, height=2)  # Adjust width and height as needed
        self.label.grid(row=1, column=0, columnspan=2)  # Center the label
        
        self.one_button = tk.Button(self.window, text="1", command=lambda: self.on_button(1), width=10, height=2)
        self.one_button.grid(row=2, column=0, sticky=tk.EW)  # Align button to the bottom, float at the bottom
        
        self.zero_button = tk.Button(self.window, text="0", command=lambda: self.on_button(0), width=10, height=2)
        self.zero_button.grid(row=2, column=1, sticky=tk.EW)  # Align button to the bottom, float at the bottom
        
    def on_button(self, value):
        self.action_on_input(value, self.label)

    def update_label(self, text):
        self.label.config(text=text)

    def run(self):
        self.window.mainloop()

def create_app(action_on_input):
    window = tk.Tk()
    app = HereApp(window, action_on_input)
    return app

