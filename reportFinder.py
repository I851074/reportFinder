import PySimpleGUI as sg
import pyperclip
import os

# Functino to display a popup window with search results
def result_popup(foundCount, notfoundCount, outputFile):
    popup = sg.Window('Results', [[sg.Text('Reports Found: ' + foundCount, size=(30, 1))],
                          [sg.Text('Reports Not Found: ' + notfoundCount, size=(30, 1))],
                        [sg.Button('Open File'), sg.Button('Close')]], modal=True, element_justification='c', keep_on_top=True)
    while True:
        action, values = popup.read()
        # Allows user to open the Output file selected
        if action == 'Open File':
            os.startfile(outputFile)
            break
        # Closes popup window
        if action == sg.WINDOW_CLOSED or action == 'Close':
            break
    popup.close()

# Function to format the list of IDs into a string for DBSelect/SQL query
def format_ids(ids):
    if len(ids) == 1 :
        return f"('{ids[0]}')"
    else:
        formatted_ids = "', '".join(ids)
        return f"('{formatted_ids}')"
    
# Function to extrat report data from files
def extractReportData(ids, files):
    ids_found = []
    extracted_lines = []
    
    for file_path in files:
            with open(file_path, 'r', encoding="utf8") as file:
                # Read each line in the file
                lines = file.readlines()
                for line in lines:
                    found = False
                    for id in ids:
                        if id in line:
                            found = True
                            ids_found.append(id)
                    if found:
                        extracted_lines.append(line)
    
    return ids_found, extracted_lines



# GUI Layout: User input for report IDs, output file selection, function buttons
left_part = [
    [sg.Text('Enter the IDs to search:', font=("Arial Bold", 10))],
    [sg.Multiline(key='-IDS-', expand_x=True, expand_y=True, size=(30,15))],
    [sg.Radio("Report", "type", key='REPORT', default=True),
     sg.Radio("Invoice", "type", key='INVOICE'), sg.Push(),
     sg.Checkbox("Payment File", key='PAYMENTCHECK'),
     sg.Push(), sg.Button("Generate DBSelect")],
     [sg.HorizontalSeparator()],
    [sg.Text('Extracts Selected: ', font=("Arial Bold", 10)), sg.Text('0', key='-INPUT-'),
     sg.Push(),
     sg.Button('Select Extracts')],
    [sg.Text('Select the Output File to Store Results:', font=("Arial Bold", 10)), sg.Push(), sg.Checkbox("Include Header", key='HEADER', default=True)],
    [sg.Input(key='-OUTPUT_FILE-', disabled=True), sg.FileBrowse()],
    [sg.Button('Search and Extract'),
     sg.Button('Exit')]
]

# GUI Layout: Table/List displaying report IDs NOT found
right_part = [
    [sg.Text('Report IDs NOT Found', font=("Arial Bold", 12), visible=True, expand_x=True, justification='center', key='-NOTFOUNDLABEL-')],
    [sg.Table(values=[], headings=['Report ID'], 
              auto_size_columns=False, 
              display_row_numbers=False, 
              key='-NOTFOUND-', 
              expand_x=True, 
              expand_y=True,
              def_col_width=40,
              num_rows=10,
              enable_events=True,
              justification='left')],
    [sg.Push(),
     sg.Button('Copy')]
]


layout = [
    [sg.Column(left_part, vertical_alignment='top'),
    sg.VSeparator(),
    sg.Column(right_part, vertical_alignment='top')]
]

# Create the window
window = sg.Window('Report Finder Tool', layout, resizable=True)

db_report_script = '''select distinct ct_report.REPORT_ID as 'Report ID', CT_JOB_DEFINITION_LANG.NAME as 'Job name', END_TIME as 'Run date', RUN_NUMBER as 'Run number', RECORDS_LOADED as 'Records' from ct_report join ct_journal on ct_report.rpt_key = ct_journal.rpt_key join CT_JOB_RUN on ct_journal.jr_key = CT_JOB_RUN.jr_key join CT_JOB_DEFINITION_LANG on CT_JOB_RUN.jd_key = CT_JOB_DEFINITION_LANG.jd_key where ct_report.REPORT_ID in '''
db_invoice_script = '''select distinct ctp_request.REQUEST_ID as 'Request ID', CT_JOB_DEFINITION_LANG.NAME as 'Job name', END_TIME as 'Run date', RUN_NUMBER as 'Run number', RECORDS_LOADED as 'Records' from ctp_request join ctp_journal on ctp_request.req_key = ctp_journal.req_key join CT_JOB_RUN on ctp_journal.jr_key = CT_JOB_RUN.jr_key join CT_JOB_DEFINITION_LANG on CT_JOB_RUN.jd_key = CT_JOB_DEFINITION_LANG.jd_key where ctp_request.REQUEST_ID in '''
db_report_payment_script = '''SELECT DISTINCT CT_REPORT.REPORT_ID as 'REPORT ID', CT_JOB_DEFINITION_LANG.NAME as 'JOB NAME', END_TIME as 'RUN DATE', RUN_NUMBER as 'RUN NUMBER', RECORDS_LOADED as 'RECORDS' FROM CT_REPORT JOIN CT_EFT_ITEM_PAYEE on CT_REPORT.RPT_KEY = CT_EFT_ITEM_PAYEE.RPT_KEY JOIN CT_EFT_PAYMENT_JOURNAL on CT_EFT_ITEM_PAYEE.EFT_PD_KEY = CT_EFT_PAYMENT_JOURNAL.EFT_PD_KEY JOIN CT_JOB_RUN on CT_EFT_PAYMENT_JOURNAL.JR_KEY = CT_JOB_RUN.JR_KEY JOIN CT_JOB_DEFINITION_LANG on CT_JOB_RUN.JD_KEY = CT_JOB_DEFINITION_LANG.JD_KEY WHERE CT_REPORT.REPORT_ID in '''
db_invoice_payment_script = '''SELECT DISTINCT CTP_REQUEST.REQUEST_ID as 'REQUEST_ID', CT_JOB_DEFINITION_LANG.NAME as 'Job name', END_TIME as 'Run date', RUN_NUMBER as 'Run number', RECORDS_LOADED as 'Records' FROM CTP_REQUEST JOIN CTP_EFT_INVOICE_PAYMENT on CTP_REQUEST.REQ_KEY = CTP_EFT_INVOICE_PAYMENT.REQ_KEY JOIN CT_JOB_RUN on CTP_EFT_INVOICE_PAYMENT.JR_KEY = CT_JOB_RUN.JR_KEY JOIN CT_JOB_DEFINITION_LANG on CT_JOB_RUN.JD_KEY = CT_JOB_DEFINITION_LANG.JD_KEY WHERE CTP_REQUEST.REQUEST_ID in '''
orderby_script = ' order by CT_JOB_DEFINITION_LANG.NAME, RUN_NUMBER'
input_files = None




# Event loop to process GUI events
while True:
    event, values = window.read()

    # Exit the program if the window is closed or 'Exit' button is clicked
    if event == sg.WINDOW_CLOSED or event == 'Exit':
        break

    if event == 'Select Extracts':
        input_files = sg.popup_get_file('Select Extracts to Search', multiple_files=True)
        if input_files:
            window['-INPUT-'].update(len(input_files.split(';')))
        else:
             window['-INPUT-'].update('0')
             input_files=None

    # Generate DBSelect Script - Creates script to enter in DBSelect to get extract dates
    if event == 'Generate DBSelect':
        ids = [id.strip() for id in values['-IDS-'].split('\n')]
        ids = list(set(filter(None, ids)))
        
        if values['REPORT']:
            if values['PAYMENTCHECK']:
                sg.popup_scrolled(db_report_payment_script + format_ids(ids) + orderby_script, title="DBSelect for Report Payments")
            else:
                sg.popup_scrolled(db_report_script + format_ids(ids) + orderby_script, title="DBSelect for Reports")
        else:
            if values['PAYMENTCHECK']:
                sg.popup_scrolled(db_invoice_payment_script + format_ids(ids) + orderby_script, title="DBSelect for Invoice Payments")
            else:
                sg.popup_scrolled(db_invoice_script + format_ids(ids) + orderby_script, title="DBSelect for Invoices")

    
    # Perform the search and extraction when 'Search and Extract' button is clicked
    if event == 'Search and Extract':
    
        
        # Check if user has selected file(s) to search from and an output file to write to
        if (values['-OUTPUT_FILE-'] and (input_files is not None)): 

            #    ids = values['-IDS-'].split(',')
            #    ids = list(set(values['-IDS-'].split('\n')))

            # Split user input IDs by newline, remove any leading/trailing spaces, duplicates, and empty lines
            ids = [id.strip() for id in values['-IDS-'].split('\n')]
            ids = list(set(filter(None, ids)))

            files = input_files.split(';')
            output_file = values['-OUTPUT_FILE-']
            
            ids_found, extracted_lines = extractReportData(ids, files)

            # Update table with report IDs not found
            window['-NOTFOUND-'].update(values=[[elem] for elem in list(set(ids) - set(ids_found))])

            if (len(ids_found) > 0):
                # Write extracted lines to output file selected by suer
                with open(output_file, 'w') as file:
                    if values['HEADER']:
                        file.write('Header\n')
                    file.writelines(extracted_lines)

                result_popup(str(len(list(set(ids_found)))), str(len(window['-NOTFOUND-'].get())), output_file)
            else:
                sg.popup("No reports found!")
        
        else:
            if not (values['-OUTPUT_FILE-']):
                sg.popup("Please select an output file!")
            
            if input_files is None:
                sg.popup("Please select extracts to search!")

    # Copy IDs from Not found table
    if event == 'Copy':
        table_values = window['-NOTFOUND-'].get()
        copyValues = ""
        for i in table_values:
            copyValues += (i[0] + '\n')
        pyperclip.copy(copyValues)

# Close the window
window.close()
