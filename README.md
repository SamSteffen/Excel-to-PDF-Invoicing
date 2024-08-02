# Excel-to-PDF-Invoicing
The following is a step-by-step guide for how to use the invoice_generator.py program to automatically generate multiple .PDF file-type invoices from an Excel spreadsheet containing data for multiple clients over a given timeframe, quickly and accurately, at the click of a button.

### PREREQUISITES
-	GitBash (Windows) or Linux (Mac)
-	Python
-	Docx (python library)
-	Docx2pdf (python library)
-	Microsoft Word
-	Microsoft Excel

### THE FILES
This program consists of (and requires) the following files:   
1.	Invoice_Data.xlsx – an Excel Spreadsheet (.xlsx) template consisting of three worksheets, “Clients”, “Timesheet”, and “Invoices.” The data the program expects to obtain from these xlsx worksheets is explained below.

2.	Invoice_generator.py – a Python program file that utilizes various programming libraries and methods to do the following:
(1) import data contained in the ‘invoice_data.xlsx’ worksheet, 
(2) compile the imported data as an invoice in a word doc (docx) file and 
(3) save the invoice as a .pdf in the ‘PDFs’ directory. 

This program must be run from the same directory (file folder) that contains the ‘invoice_data.xlsx’ spreadsheet and the (preferably empty) directory called ‘PDFs.’

3.	PDFs - An empty directory that is the output file location for the .pdf file that is created by the invoice_generator.py program. This folder should be empty when the invoice_generator.py program is run, just to avoid confusion between previous (old) invoices and new.

# Invoice_Data.xlsx
The ‘invoice_data.xlsx’ file is an Excel spreadsheet that contains the following three worksheets: (1) Clients, (2) Timesheet, (3) Invoices. An explanation of each of these worksheets and the expected data is provided below:

## .xlsx Worksheet 1: ‘Clients’
The ‘Clients’ worksheet is intended to act as a repository for all of the data pertaining to the clients of a particular business. The data contained in this worksheet is intended to be unique, meaning there should be no duplicate rows. Once the data is entered onto this page, it is added to other pages using formulas.
Column A: Client Number
 The client number is a unique ID number that is assigned to each new client in the ‘Clients’ worksheet.  ‘1’ represents the first client, ‘2’ the second client and so on. This number acts as a ‘key’ for the data associated with it, meaning it will be used to look up information about clients. In the ‘Timesheet’ and ‘Invoices’ worksheets, the client number can be used to reference all of the data associated with a particular client. This column cannot be blank, should be a sequential number. This data must be input manually and must be unique (no duplicates).
Column B: Client Name
The name of the client, written [First Name] [Middle Name] [Last Name]. The way client names are written is left to the discretion of the user; but it’s important to realize that the way the names are written in this cell is how they will appear on the ‘Timesheet’ and ‘Invoices’ worksheets, as well as on the invoice itself. This column should not be left blank. This data must be input manually.
Column C: Client Address
The complete address of the client, including street number, street name, city, province, and postal code. This column should not be left blank. This data must be input manually.
Column D: Client Phone (Primary)
The primary contact telephone number of the client, entered in the ###-###-#### format. This column should not be left blank. This data must be input manually.
Column E: Client Phone (Secondary)
The secondary contact telephone number of the client, entered in the ###-###-#### format. If there is no secondary contact telephone, this column may be left blank, or infilled with ‘N/A.’ This data must be input manually.
Column F: Client Email (Primary)
The primary email contact of the client. Email addresses must include an “@” symbol to be valid. This column should not be left blank. This data must be input manually.
Column G: Client Email (Secondary)
The secondary email contact of the client. Email addresses must include an “@” symbol to be valid. If there is no secondary contact email, this column may be left blank, or infilled with ‘N/A.’ This data must be input manually.
Column H: Preferred Payment Method
The client’s preferred method of paying invoices. Options may include (but are not limited to): (1) Cash, (2) Credit Card, (3) eTransfer, (4) Other. This data must be input manually.
Column I: Enrollment Date
The date the client was added to the ‘Clients’ page or first utilized the services offered by the business. This date should be entered in the format ‘m/d/yyyy’ or ‘mm/dd/yyyy’. This data must be entered manually.
Columns J-Z: Additional Data Columns, as needed.
There is always room to add more data to this sheet. Other things to capture from clients could include: (1) birthdates, (2) credit card numbers, (3) spouse names, etc. These columns are not used in the current iteration of this program.

## .xlsx Worksheet 2: ‘Timesheet’
The ‘Timesheet’ worksheet is intended to capture hours, rates of pay, and descriptions of services performed at client residences/ addresses. This sheet utilizes Excel’s built-in VLOOKUP(), DATE() and TEXT() formulas to save the user time in inputting information. The columns that contain formulas have been highlighted in light green, indicating to the user that NO INFORMATION SHOULD BE ADDED TO OR DELETED FROM THESE FIELDS. Based on the user’s input in the white (non-highlighted) cells, the formula-filled fields will be in-filled automatically.
The data on this page may contain duplicates, as long as the duplicate data pertains to identical clients at different times of the day, or on different days altogether. Users are cautioned to be mindful of how they are entering data in this sheet, to avoid duplicating entries and potentially overcharging clients. 
COLUMN A: Client Number
The client number is an ID number that is assigned to each new client in the ‘Clients’ worksheet.  ‘1’ represents the first client, ‘2’ the second client and so on. This unique number acts as a ‘key’ for the data associated with it, meaning it is intended to be used to look up information about clients. In the ‘Timesheet’ and ‘Invoices’ worksheets, the client number is used to reference all of the data associated with a particular client. This column cannot be blank, should be a sequential number and must be unique (no duplicates). This data must be input manually. The client number on the ‘Timesheet’ worksheet also must reference an existing client in the ‘clients’ worksheet.
When the client Number is entered on the ‘Timesheet’ worksheet, the data for column B and C (‘Client Name’ and ‘Client Address’) will be in-filled automatically.
COLUMN B: Client Name
The user need not touch this column.
This column contains a formula: ‘= IF(ISBLANK(A2), ‘’, VLOOKUP(A2, Clients!A;C, 2, FALSE))’ meaning, in row 2, if the Client Number is blank, then return a blank cell; if the Client Number is not blank and contains a valid client number from the ‘Clients’ worksheet, return the name of the client in the cell.
The client name is the name of the client, and is entered manually ONLY on the ‘Clients’ worksheet.
COLUMN C: Client Address
The user need not touch this column.

COLUMN D: Day of Service
COLUMN E: Month of Service
 COLUMN F: Year of Service
COLUMN G: Date of Service
The user need not touch this column.

COLUMN H: Weekday of Service
The user need not touch this column.

COLUMN I: Month of Service
The user need not touch this column.

COLUMN J: Start Time
COLUMN K: End Time
COLUMN L: Hours
The user need not touch this column.

COLUMN M: Description of Service(s)
COLUMN N: Rate/hr (CAD)

COLUMN O: Client Per Diem
The user need not touch this column.




## .xlsx Worksheet 3: ‘Invoices’

