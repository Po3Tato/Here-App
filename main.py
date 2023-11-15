# Import python-docx[Interact with the document]
# Import date module[Used to locate date's column and start from there]
# main.py
import os
from docx import Document
from gui import create_app
from datetime import datetime

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
    
    # Finds date in the first row    
    def find_date_column(self, date_str):
        header_row = self.table.rows[0]
        for idx, cell in enumerate(header_row.cells):
            if cell.text.strip() == date_str:
                return idx
            
        # If date is not formarted properly then go with abv. of month and today's date
        fallback_date_str = datetime.now().strftime("%b %d").lstrip("0")
        for idx, cell in enumerate(header_row.cells):
            if cell.text.strip() == fallback_date_str:
                return idx
            
        day_only_str = datetime.now().strftime("%d").lstrip("0")
        for idx, cell in enumerate(header_row.cells):
            if cell.text.strip() == day_only_str:
                return idx
        return None

    def action_on_upload(self, file_path):
         # Check if it is a .doc or .docx file being uploaded
        if not (file_path.endswith('.doc') or file_path.endswith('.docx')):
            self.app.update_label('Wrong Document type uploaded')
            return

        self.doc = Document(file_path)
        self.doc_path = file_path  # Save the document's path
        self.table = self.doc.tables[0]
        today = datetime.now()
        formatted_date = today.strftime("%B %d").lstrip("0")

        self.date_column_number = self.find_date_column(formatted_date)

        if self.date_column_number is not None:
            self.student_row_number = 1
            self.update_student_label()
        else:
            self.app.update_label(f"Date {formatted_date} not found in document")


    def action_on_button_press(self, is_present):
        if self.student_row_number < len(self.table.rows):
            if not is_present:  # Only mark "Absent" if the button pressed corresponds to absence
                self.table.cell(self.student_row_number, self.date_column_number).text = 'Absent'
            
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
