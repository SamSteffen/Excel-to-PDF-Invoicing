# import dependencies
import pandas as pd
import os
import openpyxl
import docx
import docx2pdf 

from os import remove, path
from docx import Document
from docx.shared import Inches
from docx.shared import Pt
from docx2pdf import convert
from datetime import datetime

# define the filepath to the desired xlsx workbook and docx document
cwd                = os.getcwd()
xlsx_filename      = 'Invoice_Data.xlsx'
xlsx_full_filepath = cwd + '\\' + xlsx_filename


# access the xlsx workbook
wb = openpyxl.load_workbook(filename = xlsx_full_filepath)

# create a dictionary of data from each worksheet

# 1. 'Clients'
# retrieve the client numbers from the 'Clients' worksheet
client_count            = max([item.value for item in wb['Clients']['A'] if isinstance(item.value, int) == True])
extra                   = client_count + 1
client_numbers_list_int = list(range(0, extra))
client_numbers_list_str = [str(int) for int in client_numbers_list_int]

# create a dictionary of all the client data, using the CLIENT_NUMBERS as keys and the remainder of the client data as values
client_data_dictionary  = {}
# remove the first item in the list so the full list can be iterated (row '0' is a header column)
client_numbers_list_int.pop(0)
# add the extra item to the list so the full list can be iterated
client_numbers_list_int.append(extra)
for row_num in client_numbers_list_int:
    client_data_list = []
    for row in wb['Clients'][row_num]:
        client_data_list.append(row.value)
        client_data_dictionary[(row_num-1)] = client_data_list
        
# modify the dictionary to make the data readable/usable
for datarow in list(client_data_dictionary.values())[1:]:
    datarow[8] = datarow[8].strftime('%m-%d-%Y')

# 2. 'Timesheet'
timesheet_data_dictionary = {}
entry_count = len([item.value for item in wb['Timesheet']['A'] if isinstance(item.value, int) == True])
extra = entry_count + 1
timesheet_entry_rows_list = list(range(0, extra))

# create a dictionary of all the timesheet data, using ROW_NUMBERS as kes and data as values
timesheet_data_dictionary = {}
# remove the first item in the list so full list can be iterated (row '0' is a header column)
timesheet_entry_rows_list.pop(0)
# add the extra item to the list so the full list can be iterated
timesheet_entry_rows_list.append(extra)
for row_num in timesheet_entry_rows_list:
    timesheet_data_list = []
    for row in wb['Timesheet'][row_num]:
        timesheet_data_list.append(row.value)
        timesheet_data_dictionary[(row_num-1)] = timesheet_data_list

# modify the dictionary to make the data readable/usable
for data_row in list(timesheet_data_dictionary.values())[1:]:
    for client_data in list(client_data_dictionary.values()):
        if data_row[0] == client_data[0]:                           
            data_row[1] = client_data[1]   # client name
            data_row[2] = client_data[2]   # client address
            # Date of Service
            data_row[6] = datetime(data_row[5], data_row[4], data_row[3]).strftime('%m-%d-%Y')   # Date of Service
            data_row[7] = datetime(data_row[5], data_row[4], data_row[3]).strftime('%A')         # Weekday of Service
            data_row[8] = datetime(data_row[5], data_row[4], data_row[3]).strftime('%B')         # Month of Service
            data_row[9] = data_row[9].strftime('%H:%M')                                          # start time
            data_row[10] = data_row[10].strftime('%H:%M')                                        # end time
            data_row[11] = (datetime.strptime(data_row[10], '%H:%M')\
                            -datetime.strptime(data_row[9],'%H:%M')).seconds/3600                # hrs billed
            data_row[14] = ((datetime.strptime(data_row[10], '%H:%M')\
                            -datetime.strptime(data_row[9],'%H:%M')).seconds/3600)*data_row[13]  # client per diem

# 3. 'Invoices'
# create a dictionary of all the invoice data, using the INVOICE_NUMBERS as keys and the remainder of the invoice data as values
# be sure to replace anything that appears as formulas with actual data from the client_data_list
invoice_count = len([item.value for item in wb['Invoices']['B']])-1 # subtract 1 because the first row is a header
extra = invoice_count + 1
invoice_list = [item.value for item in wb['Invoices']['B']]

invoice_data_dictionary = {}
invoice_list.pop(0)
invoice_list.insert(0, '00000')
invoice_data_rows_list = list(range(1, len(invoice_list)+1))

for i, invoice_number in enumerate(invoice_list):
    invoice_data_list = []
    for data in wb['Invoices'][i+1]:
        invoice_data_list.append(data.value)
        invoice_data_dictionary[invoice_number] = invoice_data_list

# create a unique list of the client_numbers that appear in the timesheet dataset
client_numbers_to_invoice_list = list(set([data_row[0] for data_row in list(timesheet_data_dictionary.values())[1:]]))
        
# modify the dictionary to make the data readable/usable
for data_row in list(invoice_data_dictionary.values())[1:]:
    for client_data in list(client_data_dictionary.values()):
        if data_row[5] == client_data[0]:
            data_row[2]  = data_row[2].strftime('%m-%d-%Y') # Invoice Date
            data_row[3]  = data_row[3].strftime('%m-%d-%Y') # Period Start Date
            data_row[4]  = data_row[4].strftime('%m-%d-%Y') # Period End Date
            data_row[6]  = client_data[1]                   # Client Name
            data_row[7]  = client_data[2]                   # Client Address
            data_row[8]  = client_data[3]                   # Client Phone (Primary)
            data_row[9]  = client_data[5]                   # Client Email (Primary)
            data_row[10] = client_data[7]                   # Preferred Payment Method
            data_row[11] = client_data[8]                   # Enrollment Date
        
    for timesheet_data in list(timesheet_data_dictionary.values()):
        if data_row[5] in client_numbers_to_invoice_list:
            if data_row[5] == timesheet_data[0]:
                data_row[12] = timesheet_data[11]               # Hrs Invoiced
                data_row[13] = timesheet_data[14]               # Subtotal
        else:
            data_row[12] = 0
            data_row[13] = 0

# retrieve the headers from the data_dictionaries and store them as lists
client_data_header_list    = client_data_dictionary[0]
invoice_data_header_list   = invoice_data_dictionary['00000']
timesheet_data_header_list = timesheet_data_dictionary[0]

# create a list of dictionaries that associate the header with the data
individual_invoice_list = []
for data_row in list(invoice_data_dictionary.values())[1:]:    
    temp_dict = {}
    for j, value in enumerate(data_row):
        temp_dict[invoice_data_header_list[j]] = data_row[j]
    individual_invoice_list.append(temp_dict)

# ### Create a docx doc from scratch, infilled with the desired invoice data
# create a function to make rows in a table bold
def make_rows_bold(*rows):
    for row in rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True

# loop through the invoice data to generate invoices for all clients                    
for i, data in enumerate(individual_invoice_list):
    # only target data for which invoice generation option is indicated
    if data['Generate Invoice?'] == 'Yes':
        #create a save string
        save_string = 'Invoice_'+data['Invoice Number']+'_'+data['Client Name']+'_'+data['Invoice Date']
        docx_save_string = save_string +'.docx'
        pdf_save_string  = 'PDFs/'+ save_string +'.pdf'
        
        if path.exists(docx_save_string):
            remove(docx_save_string)
            
        if path.exists(pdf_save_string):
            remove(pdf_save_string)
        
        # create a document object to store invoice information
        document = Document()
        
        # add header info here
        header_dict = {
            'business_owner_name'       : 'My Business' 
            , 'business_owner_address1' : '1234 Fake St'
            , 'business_owner_address2' : 'Edmonton, AB T6C 4C7'
            , 'business_owner_phone'    : '555-555-5555'
            , 'business_owner_email'    : 'mybusinessemail@me.com'
        }

        client_billing_info_list = list(individual_invoice_list[i].values())[6:10]
        
        ####################################
        # create the header 
        # create a line for business owner name
        business_owner_name_line = document.add_paragraph()

        # make the business_owner_name bold
        business_owner_name_line.add_run(list(header_dict.values())[0]).bold=True

        # add remaining header info to the doc
        for line in list(header_dict.values())[1:]:
            document.add_paragraph(line)

        ####################################
        document.add_paragraph().add_run('\n')
        ####################################
 
        # add invoice number
        document.add_paragraph('Invoice Number: ' + list(individual_invoice_list[i].values())[1])
        # add invoice date
        document.add_paragraph('Invoice Date: ' + list(individual_invoice_list[i].values())[2])
        document.add_paragraph('For services incurred between the dates of ' + list(individual_invoice_list[i].values())[3]\
                               + ' and ' + list(individual_invoice_list[i].values())[4])
        document.add_paragraph('Due on receipt')

        ####################################
        document.add_paragraph().add_run('\n')
        ####################################

        # create the bill_to section
        bill_to_line = document.add_paragraph()

        bill_to_line.add_run('BILL TO:').bold=True

        for line in client_billing_info_list:
            document.add_paragraph(line)

        ####################################
        document.add_paragraph().add_run('\n')
        ####################################
        
        #add a table of the description of services
        table = document.add_table(rows=1, cols=6)

        # create the table header rows by defining the cells
        header_cells        = table.rows[0].cells
        invoice_header_list = ['Date','','Description','Rate','Qty','Amount']

        for i in list(range(0, len(invoice_header_list))):
            header_cells[i].text = invoice_header_list[i]

        # create empty lists for desired data
        date_data        = []
        day_data         = []
        qty_data         = []
        description_data = []
        rate_data        = []
        amount_data      = []

        for data_row in list(timesheet_data_dictionary.values())[1:]:    
            # match data on client name
            if data_row[1] == client_billing_info_list[0]:
                date_data.append(data_row[6])                      # date
                day_data.append(data_row[7])                       # weekday
                qty_data.append(data_row[11])                      # hrs worked
                description_data.append(data_row[12])              # desription
                rate_data.append(data_row[13])                     # rate
                amount_data.append(data_row[14])                   # client per diem

        # calculate subtotals
        # total hours invoiced
        total_hours = sum(hrs for hrs in qty_data)

        # total amount due
        total_amount = sum(amount for amount in amount_data)   

        # add cumulative data for final row of table
        date_data.append('')
        day_data.append('')
        qty_data.append(total_hours)
        description_data.append('Total ($ CAD)')
        rate_data.append('')
        amount_data.append(total_amount)

        # convert ints and floats into strings
        day_data = [data[:3] for data in day_data]
        qty_data = [str(data)+' hrs' for data in qty_data]
        rate_data = ['$ '+str(data)+' /hr'for data in rate_data]
        amount_data = ['$ '+str(data) for data in amount_data]

        for i in list(range(0, len(description_data))):
            row_cells = table.add_row().cells
            row_cells[0].text = day_data[i]
            row_cells[1].text = date_data[i]
            row_cells[2].text = description_data[i]
            row_cells[3].text = rate_data[i]
            row_cells[4].text = qty_data[i]
            row_cells[5].text = amount_data[i]

        # reset the column widths in the table, where necessary
        for cell in table.columns[0].cells:
            cell.width = Inches(0.5)
        for cell in table.columns[1].cells:
            cell.width = Inches(2.0)
        for cell in table.columns[2].cells:
            cell.width = Inches(3.5)
        for cell in table.columns[3].cells:
            cell.width = Inches(2.0)
        for cell in table.columns[4].cells:
            cell.width = Inches(2.0)

        make_rows_bold(table.rows[0])
        make_rows_bold(table.rows[-1])
            
        ####################################
        # document.add_paragraph().add_run('\n')
        ####################################

        payment_info_line = document.add_paragraph()
        payment_info_line.add_run('Payment Info').bold = True
        document.add_paragraph('Send e-transfer payment to: '+ header_dict['business_owner_email'])

        # format the line spacing for all the lines in the doc
        for line in document.paragraphs:
            line.paragraph_format.space_after = Pt(1)  
            
        # save the document as a word doc
        document.save(docx_save_string)
                
        # convert the word doc to a pdf, store it in a separate folder
        convert(docx_save_string, pdf_save_string)
        
        # delete the word doc version
        remove(docx_save_string)


