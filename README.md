# Here App

Here App is a simple GUI application for marking attendance in a `.docx` document. It allows users to upload a class attendance sheet, display each student's name, and mark them as present or absent with the click of a button. The application is built using Python with `tkinter` for the GUI interface and `python-docx` for interacting with Word documents.

## Features

- Upload `.docx` files containing attendance sheets.
- Navigate through student names in the document.
- Mark attendance as present or absent.
- Save the updated document automatically.

## Installation

Before running the application, ensure you have Python installed on your system. You can download Python from [here](https://www.python.org/downloads/). After installing Python, you need to install the required packages using pip by running this command:

`pip install -r requirements.txt`

Usage

To start the application, run the main.py file:

`python main.py`

Once the application starts, follow these steps:

- Click on the "Upload Document" button to upload the attendance sheet.
- The application will display the name of the first student in the list.
- Use the "Present" or "Absent" buttons to mark the attendance for the displayed student name.
- The application will automatically proceed to the next student.
- Once all students have been marked, the application will save the document and display a confirmation message.

How It Works

- gui.py contains the HereApp class which sets up the GUI.
- main.py contains the application logic, including document handling and attendance marking.
- The HereApp class in main.py is responsible for interacting with the file and updating the GUI accordingly.
