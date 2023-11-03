from docx import Document

def print_name(table):
    student_name = table.cell(1, 0) # cell start point
    print(student_name.text)

# document location
doc_path = 'E://NO//CODE//projects//here-app-repo//class_attendance.docx'  
doc = Document(doc_path)
table = doc.tables[0]
print_name(table)
input("Press Enter to close...")
