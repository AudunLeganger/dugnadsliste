import tkinter as tk
from tkinter import filedialog
import pandas as pd

class Person:
    def __init__(self, room_number, name, gender, email, phone_number):
        self.last_name, self.first_name = name.split(",")
        self.last_name = self.last_name.strip()
        self.first_name = self.first_name.strip()
        self.room_number = room_number
        self.gender = gender
        self.email = email
        self.phone_number = phone_number
    
    def __str__(self):
        return (
            "First name: " + self.first_name + "\n" + "Last name: " + self.last_name + "\n"  + "Gender: " + self.gender + "\n" + "Room number: " + str(self.room_number) + "\n" + "Email: " + self.email + "\n" + "Phone number: " + str(self.phone_number)
        )
    
    def get_dugnad_string(self):
        return f"<{self.last_name}>, <{self.first_name}>/<{self.room_number}>"

        

def get_nonempty_sheet(xl_file):
    # Iterate over the sheets
    for sheet_name in xl_file.sheet_names:
        # Read the current sheet
        temp_df = xl_file.parse(sheet_name)

        # Check if the dataframe is not empty
        if not temp_df.empty:
            return temp_df

def row_contains_person(row):
    return not (pd.isnull(row.iloc[1]) or row.iloc[1] == "Kunde" or pd.isnull(row.iloc[0]))


        
     
def make_person_object(row):
    return Person(row.iloc[0], row.iloc[1], row.iloc[2], row.iloc[3], row.iloc[7])

def write_to_txt_file(file_path, data):
    with open(file_path, "w") as file:
        for line in data:
            file.write(line + "\n")

# This function will be used to open the file selector
def select_file():
    # This creates the root window but doesn't display it
    root = tk.Tk()
    # This line makes the root window invisible (we just need the dialog)
    root.withdraw()

    # Open the file selector dialog and store the selected file path
    xl_file_path = filedialog.askopenfilename(
        initialdir=".",
        title="Velg Excel-fil",
        filetypes=(("Excel files", "*.xlsx"), ("alle filer", "*.*"))
    )
    root.destroy()
    return xl_file_path

txt_file_path = "./beboerliste_formatert.txt"
xl_file_path = select_file()

# Create an ExcelFile object
xl = pd.ExcelFile(xl_file_path)

# Retrive the relevant dataframe
df = get_nonempty_sheet(xl)

# Find the number of rows in the dataframe
row_count = df.shape[0]

# Initiliaze empty list for person objects
persons = []

# Iterate throug the rows
for i in range(row_count):
    row = df.iloc[i]
    # Check if the row contains a person
    if row_contains_person(row):
        # Add the row to the list
        person = make_person_object(row)
        persons.append(person)

dugnad_strings = []
for person in persons:
    dugnad_strings.append(person.get_dugnad_string())

dugnad_strings.sort()
write_to_txt_file(txt_file_path, dugnad_strings)

print(len(dugnad_strings))

'''
1. Read the excel file
2. Find the sheet with data
3. Read the sheet with data
4. Add row data where columns A, B, C, and D are non-empty to a list
5. Write the a .txt-file in the required format
'''




