# Import tkinter[GUI interface]
# gui.py
import tkinter as tk
from tkinter import filedialog, font  # Import additional tkinter modules for file dialog and font customization
import os  # Import os module for path operations

# Define the HereApp class to represent the GUI application
class HereApp:
    # Initialize the HereApp instance with window, update document callback, and update student callback
    def __init__(self, window, update_document_callback, update_student_callback):
        self.window = window  # The main window of the application
        self.window.title('Here App')
        self.window.geometry('520x480') 
        icon_path = 'here_app_logo.ico'
        self.window.iconbitmap(icon_path)

        # Save the callbacks for later use
        self.update_document_callback = update_document_callback
        self.update_student_callback = update_student_callback

        label_font = font.Font(family='Helvetica', size=14, weight='bold')

        # Create a label widget for instructions
        self.label = tk.Label(window, text='Upload Class Attendance Sheet', height = 3, width = 40, font=label_font)
        self.label.pack(expand=True, padx=15, pady=10)

        # Create a label widget for showing the current student's name or status messages
        student_font = font.Font(family='Helvetica', size=28, weight='bold')
        self.student_label = tk.Label(window, text='', font=student_font, wraplength=400)
        self.student_label.pack(padx=5, pady=5)

        button_font = font.Font(family='Helvetica', size=14, weight='bold')
        self.button_frame = tk.Frame(window)
        self.button_frame.pack(side=tk.BOTTOM, pady=10)

        self.present_button = tk.Button(window, text='Present', command=lambda: self.mark_attendance(1), height = 3, width = 10, bg='green', font=button_font)
        self.present_button.pack(side=tk.LEFT, expand=True, padx=10, pady=50)

        self.absent_button = tk.Button(window, text='Absent', command=lambda: self.mark_attendance(0), height = 3, width = 10, bg='red', font=button_font)
        self.absent_button.pack(side=tk.RIGHT, expand=True, padx=10, pady=50)

        # Create an upload button widget
        self.upload_button = tk.Button(window, text='Upload Document', command=self.upload_document,  height = 3, width = 15, font=button_font)
        self.upload_button.pack(side=tk.BOTTOM, pady=5)

        self.name_label = tk.Label(window, text='made by jude :)')
        self.name_label.place(relx=1.0, rely=1.0, x=-10, y=-10, anchor='se')

    # Method to update the displayed document name
    def set_document_name(self, file_path):
        # Extract and display the document name without its file extension
        document_name = os.path.basename(file_path)
        document_name_no_extension = os.path.splitext(document_name)[0]
        self.label.config(text=document_name_no_extension)

    # Method to open a file dialog for document upload and process the selected file
    def upload_document(self):
        # Prompt the user to select a .docx or .doc file
        file_path = filedialog.askopenfilename(filetypes=[('Word Documents', '*.docx'), ('Word Document', '.doc')])
        if file_path:
            self.update_document_callback(file_path)
            # Update the label to show the name of the uploaded document
            self.set_document_name(file_path)

    # Method to update the student label text
    def update_label(self, text):
        self.student_label.config(text=text)

    # Method to handle attendance marking action
    def mark_attendance(self, presence):
        # Notify the main app logic of the student's presence or absence
        self.update_student_callback(presence)

# Factory function to create and return a new instance
def create_app(update_document_callback, update_student_callback):
    window = tk.Tk()
    window.geometry('520x480')
    app = HereApp(window, update_document_callback, update_student_callback)
    return app
