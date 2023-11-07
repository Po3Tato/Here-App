# gui_module.py
import tkinter as tk
from tkinter import filedialog, font  # Import additional tkinter modules for file dialog and font customization
import os  # Import os module for path operations

# Define the HereApp class to represent the GUI application
class HereApp:
    # Initialize the HereApp instance with window, update document callback, and update student callback
    def __init__(self, window, update_document_callback, update_student_callback):
        self.window = window  # The main window of the application
        self.window.title('Here App')  # Title of the window
        self.window.geometry('350x300')  # Size of the window

        # Save the callbacks for later use
        self.update_document_callback = update_document_callback
        self.update_student_callback = update_student_callback

        # Create a label widget for instructions
        self.label = tk.Label(window, text='Please upload Class Attendance Sheet', wraplength=250)
        self.label.pack(pady=20)  # Place the label in the window with padding

        # Create an upload button widget
        self.upload_button = tk.Button(window, text='Upload Document', command=self.upload_document,  height = 3, width = 15)
        self.upload_button.pack(pady=10)  # Place the button in the window with padding

        # Create a label widget for showing the current student's name or status messages
        self.student_label = tk.Label(window, text='', font=font.Font(size=16), wraplength=200)
        self.student_label.pack(pady=10)  # Place the label in the window with padding

        # Create a button widget for marking attendance as present
        self.present_button = tk.Button(window, text='Present', command=lambda: self.mark_attendance(1), height = 3, width = 10)
        self.present_button.pack(side=tk.LEFT, expand=True, padx=5, pady=10)

        # Create a button widget for marking attendance as absent
        self.absent_button = tk.Button(window, text='Absent', command=lambda: self.mark_attendance(0), height = 3, width = 10)
        self.absent_button.pack(side=tk.RIGHT, expand=True, padx=5, pady=10)

    # Method to update the displayed document name
    def set_document_name(self, file_path):
        # Extract and display the document name without its file extension
        document_name = os.path.basename(file_path)
        document_name_no_extension = os.path.splitext(document_name)[0]
        self.label.config(text=document_name_no_extension)

    # Method to open a file dialog for document upload and process the selected file
    def upload_document(self):
        # Prompt the user to select a .docx file
        file_path = filedialog.askopenfilename(filetypes=[('Word Documents', '*.docx')])
        if file_path:  # Check if the user has selected a file
            # Trigger the update document callback with the selected path
            self.update_document_callback(file_path)
            # Update the label to show the name of the uploaded document
            self.set_document_name(file_path)

    # Method to update the student label text
    def update_label(self, text):
        self.student_label.config(text=text)  # Set the label text to the provided string

    # Method to handle attendance marking action
    def mark_attendance(self, presence):
        # Notify the main app logic of the student's presence or absence
        self.update_student_callback(presence)

# Factory function to create and return a new instance
def create_app(update_document_callback, update_student_callback):
    window = tk.Tk()
    window.geometry('350x300')
    app = HereApp(window, update_document_callback, update_student_callback)
    return app
