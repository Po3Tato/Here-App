# import docx module
from docx import Document

def print_name(table, row=1):
    # loop start here
    for row_number in range(1, len(table.rows)):
        student_name = table.cell(row_number, 0) 
        print(student_name.text) 
        # user input
        user_input = input("Enter 1 for 'here', 0 for 'absent' or press 'Enter' to close: ")

        if user_input == '1':
            continue
        elif user_input == '0':
            print("Not here.")
        elif user_input == '':
            print("Closing...")
            break
# path of file. Change later [refer to notes on notion]
doc_path = 'E://NO//CODE//projects//here-app-repo//class_attendance.docx'
doc = Document(doc_path)
table = doc.tables[0]

print_name(table)
