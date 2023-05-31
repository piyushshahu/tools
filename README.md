# tools
productivity tools : Certificate Maker
Change the following in the code:

First of all install libraries

pip install openpyxl Pillow fpdf

# Specify the paths to the Excel file, certificate template image, and output directory
Make sure to replace the placeholders 
"path_to_font.ttf",
"path_to_excel_file.xlsx", 
"path_to_certificate_template.jpg", 
and "path_to_output_directory" with the appropriate file paths and directories.

#The program renames the file using the corresponding name from the Excel sheet. 
 It assumes that the new name is present in the second column of the Excel sheet. **You can modify the index (row[1]) if the new name is in a different column**.
