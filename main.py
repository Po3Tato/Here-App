import os
import tkinter as tk
from tkinter import filedialog, font
from docx import Document
from datetime import datetime

class HereApp:
    def __init__(self, window, update_document_callback, update_student_callback):
        # Initialize the HereApp gui box
        self.window = window
        self.window.title('Here App')
        self.window.geometry('520x480')

        self.update_document_callback = update_document_callback
        self.update_student_callback = update_student_callback

        label_font = font.Font(family='Helvetica', size=14, weight='bold')
        # Create and configure the main label
        self.label = tk.Label(window, text='Upload Class Attendance Sheet', height=3, width=40, font=label_font)
        self.label.pack(expand=True, padx=15, pady=10)

        student_font = font.Font(family='Helvetica', size=28, weight='bold')
        # Create and configure the label for student names
        self.student_label = tk.Label(window, text='', font=student_font, wraplength=400)
        self.student_label.pack(padx=5, pady=5)

        button_font = font.Font(family='Helvetica', size=14, weight='bold')
        # Create frame for buttons
        self.button_frame = tk.Frame(window)
        self.button_frame.pack(side=tk.BOTTOM, pady=10)

        # Create Present and Absent buttons
        self.present_button = tk.Button(window, text='Present', command=lambda: self.mark_attendance(1), height=3, width=10, bg='green', font=button_font)
        self.present_button.pack(side=tk.LEFT, expand=True, padx=10, pady=50)
        self.absent_button = tk.Button(window, text='Absent', command=lambda: self.mark_attendance(0), height=3, width=10, bg='red', font=button_font)
        self.absent_button.pack(side=tk.RIGHT, expand=True, padx=10, pady=50)

        # Create an Upload Document button
        self.upload_button = tk.Button(window, text='Upload Document', command=self.upload_document, height=3, width=15, font=button_font)
        self.upload_button.pack(side=tk.BOTTOM, pady=5)

        self.name_label = tk.Label(window, text='made by jude :)')
        self.name_label.place(relx=1.0, rely=1.0, x=-10, y=-10, anchor='se')

        # Variables to store document information
        self.doc = None
        self.table = None
        self.doc_path = None
        self.student_row_number = 0

    def run(self):
        # Start the main loop of the Tkinter window
        self.update_label('Upload Attendance Document')
        self.window.mainloop()

    def find_today_date_column(self):
    # Get today's date in the format "Jan 1" and "JAN 1"
        today_date_str = datetime.now().strftime("%b %d").replace(" 0", " ")
        uppercased_today_date_str = today_date_str.upper()

    # Checks through all cells in the first row
        header_row = self.table.rows[0]
        for date, cell in enumerate(header_row.cells):
            cell_text = cell.text.strip()

            if today_date_str in cell_text or uppercased_today_date_str in cell_text:
                return date
        return None

    def action_on_upload(self, file_path):
        # Actions when document is uploaded
        if not (file_path.endswith('.doc') or file_path.endswith('.docx')):
            self.update_label('Wrong Document type uploaded')
            return

        self.doc = Document(file_path)
        self.doc_path = file_path
        self.table = self.doc.tables[0]
        today = datetime.now()
        formatted_date = today.strftime("%b %d").lstrip("0")

        self.date_column_number = self.find_today_date_column()

        if self.date_column_number is not None:
            self.student_row_number = 1
            self.update_student_label()
        else:
            self.update_label(f"Date {formatted_date} not found in document")

    def action_on_button_press(self, is_present):
        # Actions when a button is pressed
        if self.student_row_number < len(self.table.rows):
            if not is_present:
                self.table.cell(self.student_row_number, self.date_column_number).text = 'Absent'

            self.student_row_number += 1
            self.update_student_label()
        else:
            self.save_document()

    def update_student_label(self):
        # Update the student label with the current student's information
        if self.student_row_number < len(self.table.rows):
            student_name = self.table.cell(self.student_row_number, 0).text
            self.update_label(student_name)
        else:
            self.save_document()
            self.update_label('Attendace Complete')

    def save_document(self):
        # Save the modified document
        if self.doc_path:
            self.doc.save(self.doc_path)
            self.update_label('Document saved')

    def set_document_name(self, file_path):
        # Set the label with the document name
        document_name = os.path.basename(file_path)
        document_name_no_extension = os.path.splitext(document_name)[0]
        self.label.config(text=document_name_no_extension)

    def upload_document(self):
        # Prompts the user to select file
        file_path = filedialog.askopenfilename(filetypes=[('Word Documents', '*.docx'), ('Word Document', '.doc')])
        if file_path:
            self.action_on_upload(file_path)
            self.set_document_name(file_path)

    def update_label(self, text):
        # Update the student label text
        self.student_label.config(text=text)

    def mark_attendance(self, presence):
        self.update_student_callback(presence)


def application(window, update_document_callback, update_student_callback):
    app = HereApp(window, update_document_callback, update_student_callback)
    return app

if __name__ == '__main__':
    window = tk.Tk()
    here_app = application(window, None, None)

    update_document_callback = lambda file_path: here_app.action_on_upload(file_path)
    update_student_callback = lambda is_present: here_app.action_on_button_press(is_present)

    here_app.update_document_callback = update_document_callback
    here_app.update_student_callback = update_student_callback

    here_app.run()
