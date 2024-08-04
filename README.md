# Excel-to-PDF-Invoicing
The following is a step-by-step guide for how to use the invoice_generator.py program to automatically generate multiple .PDF file-type invoices from an Excel spreadsheet containing data for multiple clients over a given timeframe, quickly and accurately, at the click of a button.

### PREREQUISITES
-	GitBash (Windows) or Linux (Mac)
-	Python
-	Docx (python library)
-	Docx2pdf (python library)
-	Microsoft Word
-	Microsoft Excel

# THE FILES
This program consists of (and requires) the following files:   
1.	**Invoice_Data.xlsx** – an Excel Spreadsheet (.xlsx) template consisting of three worksheets, “Clients”, “Timesheet”, and “Invoices.” The data the program expects to obtain from these xlsx worksheets is explained below.

2.	**Invoice_generator.py** – a Python program file that utilizes various programming libraries and methods to do the following:    

(1) import data contained in the ‘invoice_data.xlsx’ worksheet,     
(2) compile the imported data as an invoice in a word doc (docx) file and     
(3) save the invoice as a .pdf in the ‘PDFs’ directory.     
This program must be run from the same directory (file folder) that contains the ‘invoice_data.xlsx’ spreadsheet and the (preferably empty) directory called ‘PDFs.’

3.	**PDFs/** - An empty directory that is the output file location for the .pdf file that is created by the invoice_generator.py program. This folder should be empty when the invoice_generator.py program is run, just to avoid confusion between previous (old) invoices and new.

# PREPARING THE DATA
## Invoice_Data.xlsx
The ‘invoice_data.xlsx’ file is an Excel spreadsheet that contains the following three worksheets:    
(1) Clients     
(2) Timesheet    
(3) Invoices    
An explanation of each of these worksheets and the expected data is provided below:

### .xlsx Worksheet 1: ‘Clients’
The ‘Clients’ worksheet is intended to act as a repository for all of the data pertaining to the clients of a particular business. The data contained in this worksheet is intended to be unique, meaning there should be no duplicate rows. Once the data is entered onto this page, it is added to other pages using formulas.  

![clients.png](https://github.com/SamSteffen/Excel-to-PDF-Invoicing/blob/main/Images/clients.jpg)

Column A: **Client Number**     
The client number is a unique ID number that is assigned to each new client in the ‘Clients’ worksheet.  ‘1’ represents the first client, ‘2’ the second client and so on. This number acts as a ‘key’ for the data associated with it, meaning it will be used to look up information about clients. In the ‘Timesheet’ and ‘Invoices’ worksheets, the client number can be used to reference all of the data associated with a particular client. This column cannot be blank and should be a sequential number. This data must be input manually (as new clients are enrolled) and must be unique (no duplicates).       

Column B: **Client Name**     
The name of the client, written [First Name] [Middle Name] [Last Name]. The way client names are written is left to the discretion of the user; but it’s important to realize that the way the names are written in this cell is how they will appear on the ‘Timesheet’ and ‘Invoices’ worksheets, as well as on the invoice itself. This column should not be left blank. This data must be input manually (as new clients are enrolled).          

Column C: **Client Address**          
The complete address of the client, including street number, street name, city, province, and postal code. This column should not be left blank. This data must be entered manually (as new clients are enrolled).   

Column D: **Client Phone (Primary)**     
The primary contact telephone number of the client, entered in the ###-###-#### format. This column should not be left blank. This data must be entered manually (as new clients are enrolled).     

Column E: **Client Phone (Secondary)**     
The secondary contact telephone number of the client, entered in the ###-###-#### format. If there is no secondary contact telephone, this column may be left blank, or infilled with ‘N/A.’ This data must be entered manually (as new clients are enrolled).    

Column F: **Client Email (Primary)**     
The primary email contact of the client. Email addresses must include an “@” symbol to be valid. This column should not be left blank. This data must be entered manually (as new clients are enrolled).     

Column G: **Client Email (Secondary)**     
The secondary email contact of the client. Email addresses must include an “@” symbol to be valid. If there is no secondary contact email, this column may be left blank, or infilled with ‘N/A.’ This data must be entered manually (as new clients are enrolled).    

Column H: **Preferred Payment Method**     
The client’s preferred method of paying invoices. Options may include (but are not limited to): (1) Cash, (2) Credit Card, (3) eTransfer, (4) Other. This data must be entered manually (as new clients are enrolled).     

Column I: **Enrollment Date**     
The date the client was added to the ‘Clients’ page or first utilized the services offered by the business. This date should be entered in the format ‘m/d/yyyy’ or ‘mm/dd/yyyy’. This data must be entered manually (as new clients are enrolled).     

Columns J-Z: *Additional Data Columns, as needed.*     
There is always room to add more data to this sheet. Other things to capture from clients could include: (1) birthdates, (2) credit card numbers, (3) spouse names, (4)billing address (if different from residential address), etc. These columns are not used in the current iteration of this program.     

### .xlsx Worksheet 2: ‘Timesheet’
The ‘Timesheet’ worksheet is intended to capture hours, rates of pay, and descriptions of services performed at client residences/ addresses. This sheet utilizes Excel’s built-in VLOOKUP(), DATE() and TEXT() formulas to save the user time in entering information. The columns that contain formulas have been highlighted in light green, indicating to the user that NO INFORMATION SHOULD BE ADDED TO OR DELETED FROM THESE FIELDS. Based on the user’s input in the white (non-highlighted) cells, the formula-filled fields will be in-filled automatically.    
The data on this page may contain duplicates, as long as the duplicate data pertains to identical clients at different times of the day, or on different days altogether. Users are cautioned to be mindful of how they are entering data in this sheet, to avoid duplicating entries and potentially overcharging clients.    
In the event that multiple services occur for a client in the course of a single day, the information can be entered on multiple lines, but the lines will also appear separately on the invoice.

![timesheet.png](https://github.com/SamSteffen/Excel-to-PDF-Invoicing/blob/main/Images/Timesheet.jpg)

COLUMN A: **Client Number**    
The client number is an ID number that is assigned to each new client in the ‘Clients’ worksheet.  ‘1’ represents the first client, ‘2’ the second client and so on. This unique number acts as a ‘key’ for the data associated with it, meaning it is intended to be used to look up information about clients. In the ‘Timesheet’ and ‘Invoices’ worksheets, the client number is used to reference all of the data associated with a particular client. This column cannot be blank, should be a sequential number and must be unique (no duplicates). This data must be input manually. The client number on the ‘Timesheet’ worksheet also must reference an existing client in the ‘clients’ worksheet.    
When the client Number is entered on the ‘Timesheet’ worksheet, the data for column B and C (‘Client Name’ and ‘Client Address’) will be in-filled automatically.

COLUMN B: **Client Name**    
The user need not touch this column.    
This column contains a formula: ‘= IF(ISBLANK(A2), ‘’, VLOOKUP(A2, Clients!A;C, 2, FALSE))’ meaning, in row 2, if the Client Number is blank, then return a blank cell; if the Client Number is not blank and contains a valid client number from the ‘Clients’ worksheet, return the name of the client associated with that Client Number in the cell.    
The client name is the name of the client, and is entered manually ONLY on the ‘Clients’ worksheet.    

COLUMN C: **Client Address**
The user need not touch this column.
This column contains a formula: '= IF(ISBLANK(A2), '', VLOOKUP(Timesheet!A2, Clients!A:C, 3, FALSE))' meaning, in row 2, if the Client Number is blank, then return a blank cell; if the Client Number is not blank and contains a valid client number from the 'Clients' worksheet, return the address associated with that Client Number of the client in the cell.
The client address is the address associated with the client, and is entere manually ONLY on the 'Clients' worksheet.

COLUMN D: **Day of Service**    
An integer value between 1-31 representing the day of the month on which the billable services listed in the 'Description of Service(s)' column were performed.

COLUMN E: **Month of Service**    
An integer value between 1-12 representing the month of the year in which billable services listed in the 'Description of Service(s)' column were performed. (1= January, 2= February, 3= March, 4= April, 5= May, 6= June, 7= July, 8= August, 9= September, 10= October, 11=November, 12= December)

COLUMN F: **Year of Service**    
A 4 digit integer value representing the year in which billable services listed in the 'Description of Service(s)' column were performed.

COLUMN G: **Date of Service**    
The user need not touch this column.
This column contains a formula: '= IF(ISBLANK(D2), '', IF(ISBLANK(E2), '', IF(ISBLANK(F2), '', DATE(F2, E2, D2))))' meaning, in row 2, if the 'Day of Service', 'Month of Service' or 'Year of Service' values are blank, return a blank cell; if these values are all infilled with valid integer values, then return the generated date value in mm/dd/yyyy format.
The date of service is the date associated with the labor that the client is being invoiced for.

COLUMN H: **Weekday of Service**    
The user need not touch this column.
This column contains a formula: '= TEXT(G2, "dddd")' which returns the full word of the day of the week on which the billable services listed in the 'Description of Service(s)' column were performed (i.e., "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" or "Sunday")

COLUMN I: **Month of Service**    
The user need not touch this column.
This column contains a formula: '= TEXT(G2, "mmmm")' which returns the full word of the month of the year in which the billable services listed in the 'Description of Service(s)' column were performed (i.e., "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November" or "December")

COLUMN J: **Start Time**    
A time value entered in the hh:mm format indicating the time the services being billed for were begun. HH values should be an integer between 0-11 for a.m. values and 13-24 for p.m. values. MM values should be integers between 00 and 59. Start times must precede end times to avoid the return of negative hours. While any valid time value betwee 00:00 and 23:59 will be accepted here, it is recommended that the user either round up or down to the nearest quarter hour or half-hour, depending upon the user's preference, to make the billing easier to interpret on the client side. 

COLUMN K: **End Time**    
A time value entered in the hh:mm format indicating the time the services being billed for were completed. HH values should be an integer between 0-11 for a.m. values and 13-24 for p.m. values. MM values should be integers between 00 and 59. End times must follow start times to avoid the return of negative hours. While any valid time value betwee 00:00 and 23:59 will be accepted here, it is recommended that the user either round up or down to the nearest quarter hour or half-hour, depending upon the user's preference, to make the billing easier to interpret on the client side. 

COLUMN L: **Hours**    
The user need not touch this column.
This column contains a formula: '= ((K2-J2)*1440)/60' This formula subtracts the Start Time from the End Time and multiplies the hh:mm difference by 1440, then divides the result by 60 to retrieve the number of hours, as a float (decimal) elapsed between the start and end time. These are the billable hours to be listed and summed on the invoice.

COLUMN M: **Description of Service(s)**    
A description of services performed. The level of detail here is left to the business owner's discretion. In its current iteration, the invoice document is written to print the description of service on the invoice itself. Lengthy descriptions may result in multiple page-length invoices.

COLUMN N: **Rate/hr (CAD)**    
The rate of charge, per hour, for services, in Canadian dollars. Should be a whole number or float (decimal) value. Do not include dollar signs.

COLUMN O: **Client Per Diem**    
The user need not touch this column.
This column contains a formula: '= L2*N2'. This is the number of hours charged multiplied by the hourly pay rate. The result is the amount that the client would be invoiced if they were billed for the day.


### .xlsx Worksheet 3: 'Invoices'
The 'Invoices' worksheet is intended to provide a summary of the hours to be billed to each client in a given time period. It is also meant to provide a record of past invoices, for the benefit of the business owner. 

![invoices.png](https://github.com/SamSteffen/Excel-to-PDF-Invoicing/blob/main/Images/Invoices.jpg)

Column A: **Generate Invoice?**    
A "Yes" or "No" value that tells the program whether to generate the invoice for the associated client in the execution of the program. Selecting "Yes" will generate an invoice for the associated client for all the billable hours and services that took place between the 'Period Start Date' and the 'Period End Date' fields. Once the program is executed and the invoices are generated, the values for this column should be set to "No" to avoid duplication of identical invoices in future runs.

Column B: **Invoice Number**    
A unique 5-digit number that is assigned to each invoice. 

Column C: **Invoice Date**
The date that the invoice is being generated. This date will appear on the invoice itself as the invoice date.

Column D: **Period Start Date**    
The earliest date to include in the time period of billable services. This value should be entered in the 'mm/dd/yyyy' format. When the program is run, the invoice will include any billable services that occurred between this date and the 'Period End Date' values.

Column E: **Period End Date**    
The latest date to include in the time period of billable services. This value should be entered in the 'mm/dd/yyyy' format. When the program is run, the invoice will include any billable services that occurred between the 'Period Start Date' and this date value.

Column F: **Client Number**    
The client number is an ID number that is assigned to each new client in the ‘Clients’ worksheet.  ‘1’ represents the first client, ‘2’ the second client and so on. This unique number acts as a ‘key’ for the data associated with it, meaning it is intended to be used to look up information about clients. In the ‘Timesheet’ and ‘Invoices’ worksheets, the client number is used to reference all of the data associated with a particular client. This column cannot be blank, should be a sequential number and must be unique (no duplicates). This data must be input manually. The client number on the ‘Invoices’ worksheet also must reference an existing client in the ‘clients’ worksheet.    
When the client Number is entered on the 'Invoices' worksheet, the data for columns G, H, I, J, K, L, M, and N (‘Client Name',‘Client Address’, 'Client Phone (Primary)', 'Client Email (Primary)', 'Preferred Payment Method', 'Enrollment Date', 'Hrs Invoiced', and 'Subtotal') will be in-filled automatically.

Column G: **Client Name**    
The user need not touch this column.
This column contains a formula: '=VLOOKUP($F2, Clients!$A:$I, 2, FALSE)' meaning, in row 2, select the input client number from column F and locate client name associated with the number in the 'clients' sheet, and return the value in the cell.
The client name is the name of the client, and is entered manually ONLY on the ‘Clients’ worksheet.    

Column H: **Client Address**    
The user need not touch this column.
This column contains a formula: '=VLOOKUP($F2, Clients!$A:$I, 3, FALSE)' meaning, in row 2, select the input client number from column F and locate the client address associated with the number in the 'clients' sheet, and return the value in the cell.

Column I: **Client Phone (Primary)**    
The user need not touch this column.
This column contains a formula: '=VLOOKUP($F2, Clients!$A:$I, 4, FALSE)' meaning, in row 2, select the input client number from column F and locate the primary client phone number associated with the number in the 'clients' sheet, and return the value in the cell.

Column J: **Client Email (Primary)**    
The user need not touch this column.
This column contains a formula: '=VLOOKUP($F2, Clients!$A:$I, 6, FALSE)' meaning, in row 2, select the input client number from column F and locate the primary client email address associated with the number in the 'clients' sheet, and return the value in the cell.

Column K: **Preferred Payment Method**    
The user need not touch this column.
This column contains a formula: '=VLOOKUP($F2, Clients!$A:$I, 8, FALSE)' meaning, in row 2, select the input client number from column F and locate the client's preferred payment method associated with the number in the 'clients' sheet, and return the value in the cell.

Column L: **Enrollment Date**    
The user need not touch this column.
This column contains a formula: '=TEXT(VLOOKUP($F2, Clients!$A:$I, 9, FALSE), 'mm/dd/yyyy')' meaning, in row 2, select the input client number from column F and locate the client's enrollment date associated with the number in the 'clients' sheet, format it to return in the 'mm/dd/yyyy' format, and return the value in the cell.

Column M: **Hrs Invoiced**    
The user need not touch this column.
This column contains a formula: '=SUMIFS(Timesheet!$L:$L, Timesheet!$A:$A, $F2, Timesheet!$G:$G, ">"&$D2, Timesheet!$G:$G, "<"&$E2)' meaning, if the 'Date of Service' on the 'Timesheet' worksheet falls between the "Period Start Date" and "Period End Date" fields on the current (Invoices) worksheet, sum the value of the billable hours for the client with the associated client number.

Column N: **Subtotal**    
The user need not touch this column.
This column contains a formula: "="$ "& SUMIFS(Timesheet!$O:$O, Timesheet!$A:$A, $F2, Timesheet!$G:$G, ">"&$D2, Timesheet!$G:$G, ">"&$E2)" meaning, if the "Date of Service' on the 'Timesheet' worksheet falls between the "Period Start Date" and "Period End Date" fileds on the current (invoices) worksheet, sum the value of the Per Diem amounts billed for the clients with the associated client number.

# EXECUTING THE CODE
Once the data has been entered into the 'Invoice_Data.xlsx' workbook and the desired values set on the 'Invoices' worksheet, delete all the (prexisting PDF) files from the PDFs directory and execute the code.

To execute the Invoice_generator.py file, navigate to the project directory in your GitBash or Linux terminal. Enter the command:

```python Invoice_generator.py```

When the code has completed execution, open the 'PDFs' directory to view the generated invoices. The number of pdfs generated should correspond to the number of lines from the 'Invoices' worksheet in the 'Invoice_Data.xlsx' workbook for which a "Yes" value was indicated in the "Generate Invoices?" column.

![PDFs_folder_output.png](https://github.com/SamSteffen/Excel-to-PDF-Invoicing/blob/main/Images/PDFs_folder_output.jpg)

Note that the PDFs are saved as a string of data that includes the invoice number, the name of the client associated with the invoice and the invoice date. This string can be modified to suit the user's preference.

Open a PDF to view your automatically generated invoice:
![invoice_00002.png](https://github.com/SamSteffen/Excel-to-PDF-Invoicing/blob/main/Images/Invoice_00002.jpg)


