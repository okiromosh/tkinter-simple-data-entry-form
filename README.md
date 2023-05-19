# Data Entry Form

This script is a simple data entry form that allows users to input product information such as product name, category, and color. The entered data is then saved to an Excel file named "Data.xlsx".

## Prerequisites

- Python 3 installed
- Required libraries: `openpyxl`, `tkinter`

## Usage

1. Make sure you have the Excel file "Data.xlsx" in the same directory as the script.
2. Run the script using Python.
3. The data entry form window will appear.
4. Fill in the product name, select the category from the drop-down menu, and enter the color.
5. Click the "Submit" button to save the data to the Excel file.
6. To clear the input fields, click the "Clear" button.
7. A message box will pop up indicating the success of the data saving process.

## Notes

- The script uses the `openpyxl` library to interact with Excel files.
- The `tkinter` library is used for creating the graphical user interface (GUI) window.
- The script assumes that the Excel file "Data.xlsx" already exists. If it doesn't, please create an empty Excel file with that name before running the script.
- The script saves the data to the first sheet of the Excel file ("Data.xlsx").
- The product name, category, and color are saved in columns 1, 2, and 3 respectively.
- The script handles saving the data and displaying a success message using a message box.

Feel free to modify and extend this script according to your specific requirements.
