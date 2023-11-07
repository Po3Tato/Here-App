# import needed modules. tkinter[GUI interface]
# and python-docx[Interact with the document]
# will be the modules used
# here_app.py
import os
from docx import Document
from gui import create_app

class HereApp:
    def __init__(self):
        self.doc = None
        self.table = None
        self.doc_path = None  # Keep track of the current document path
        self.student_row_number = 0
        self.app = create_app(self.action_on_upload, self.action_on_button_press)

    def run(self):
        self.app.update_label('Upload Attendance Document')
        self.app.window.mainloop()

    def action_on_upload(self, file_path):
        self.doc = Document(file_path)
        self.doc_path = file_path  # Save the document's path
        self.table = self.doc.tables[0]
        self.student_row_number = 1  # Start at the first student row
        self.update_student_label()

    def action_on_button_press(self, is_present):
        if self.student_row_number < len(self.table.rows):
            if not is_present:  # Only mark "Absent" if the button pressed corresponds to absence
                self.table.cell(self.student_row_number, 1).text = 'Absent'
            
            # Move to the next student regardless of presence
            self.student_row_number += 1
            self.update_student_label()
        else:
            self.save_document()  # Automatically save the document when done

    def update_student_label(self):
        # Check if there are more students to process
        if self.student_row_number < len(self.table.rows):
            student_name = self.table.cell(self.student_row_number, 0).text
            self.app.update_label(student_name)
        else:
            # No more students left to process, automatically save the document and notify the user
            self.save_document()
            self.app.update_label('Attendance marking complete. The document has been saved.')

    def save_document(self):
        if self.doc_path:  # Check if a document is loaded
            self.doc.save(self.doc_path)  # Save the document
            self.app.update_label('Document saved: ' + os.path.basename(self.doc_path))

if __name__ == '__main__':
    attendance_app = HereApp()
    attendance_app.run()
