"""
PCC Classlist Party Remixer
---------------------------
A tool for comparing first-day and second-week class lists
to identify adds and drops.
Part of the PCC Classlist Party Pack.
"""
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Set up the Tkinter root window (this is needed to open the file dialog)
root = tk.Tk()
root.withdraw()  # Hide the root window, we only want the file dialog

# Ask the user to select the first class list (from the first week)
first_week_file = filedialog.askopenfilename(title="Select the First Week Class List",
                                             filetypes=[("Excel files", "*.xlsx")])

# Ask the user to select the second class list (from the second Wednesday)
second_week_file = filedialog.askopenfilename(title="Select the Second Week Class List",
                                              filetypes=[("Excel files", "*.xlsx")])

# Load the selected files into pandas (Excel format)
first_week_df = pd.read_excel(first_week_file, sheet_name=None)  # Load all sheets
second_week_df = pd.read_excel(second_week_file, sheet_name=None)  # Load all sheets

# Initialize lists to hold the added and dropped students
added_students = []
dropped_students = []

# Iterate over each sheet (representing a class) in the first week's file
for crn, first_week_class in first_week_df.items():
    # Check if the same class (CRN) exists in the second week's file
    if crn in second_week_df:
        second_week_class = second_week_df[crn]

        # Merge the first and second week dataframes for the class based on G Number (students)
        merged_df = pd.merge(first_week_class, second_week_class, on=["G Number"], how="outer", indicator=True)

        # Add suffixes to the columns to distinguish between first week and second week data
        merged_df = pd.merge(first_week_class, second_week_class, on=["G Number"], how="outer", indicator=True,
                             suffixes=('_first_week', '_second_week'))

        # Identify added students (present in the second week but not the first week)
        added_students_class = merged_df[merged_df['_merge'] == 'right_only'][
            ['First Name_second_week', 'Last Name_second_week', 'G Number', 'PCC email address_second_week',
             'Class_second_week', 'Term_second_week', 'CRN_second_week']]
        added_students_class.columns = ['First Name', 'Last Name', 'G Number', 'PCC email address', 'Class', 'Term',
                                        'CRN']
        added_students.extend(added_students_class.values.tolist())

        # Identify dropped students (present in the first week but not the second week)
        dropped_students_class = merged_df[merged_df['_merge'] == 'left_only'][
            ['First Name_first_week', 'Last Name_first_week', 'G Number', 'PCC email address_first_week',
             'Class_first_week', 'Term_first_week', 'CRN_first_week']]
        dropped_students_class.columns = ['First Name', 'Last Name', 'G Number', 'PCC email address', 'Class', 'Term',
                                          'CRN']
        dropped_students.extend(dropped_students_class.values.tolist())

# Convert the lists to DataFrames for easy export to Excel
added_df = pd.DataFrame(added_students,
                        columns=['First Name', 'Last Name', 'G Number', 'PCC email address', 'Class', 'Term', 'CRN'])
dropped_df = pd.DataFrame(dropped_students,
                          columns=['First Name', 'Last Name', 'G Number', 'PCC email address', 'Class', 'Term', 'CRN'])

# **De-duplication Step**: Remove duplicates in the final DataFrames based on G Number and CRN
added_df = added_df.drop_duplicates(subset=['G Number', 'CRN'])
dropped_df = dropped_df.drop_duplicates(subset=['G Number', 'CRN'])

# Export the results to Excel
with pd.ExcelWriter('students_changes.xlsx') as writer:
    added_df.to_excel(writer, sheet_name="Added Students", index=False)
    dropped_df.to_excel(writer, sheet_name="Dropped Students", index=False)

print("Comparison completed. Added and dropped students have been saved as Excel files.")
