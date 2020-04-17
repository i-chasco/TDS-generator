### Date created
This project was created on 17th April, 2020.

### Project Title
TDS-generator

### Description
Python script to automate the generation of Technical Data Sheets for a chemical product.

How does it work:

- Creates directory with today's date to store the output files.
- Loads product database (Excel)
- Loads TDS template (one of 3 different models based on info of the product name in the database)
- Fills Word MailMerge document with the relevant info from the database
- Saves Word file to the newly created directory
- Converts Word file to PDF and saves it in the same directory

Modules and packages used:

- mailmerge
- datetime
- docxs2pdf
- openpyxl
- os

### Files used
script_TDS.py, 'Crear Fichas en WORD y PDF (EN).bat', liquid_template_EN.docx, th_template_EN.docx,
paste_template_EN.dox, base_datos_TDS.xlsx, .gitignore


### Credits
Iker Chasco Llorente
