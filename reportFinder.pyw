import PySimpleGUI as sg
import pyperclip
import os
import webbrowser
import re
import time

# Search and Extract Result popup
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

# DBSELECT popup
def db_popup(ids, script, title):
    popup = sg.Window(title, [
        [sg.Multiline(script + format_ids(ids) + orderby_script, expand_x=True, expand_y=True, key='-DBSELECT-')],
        [sg.Button('Copy'), sg.Button('Close')]], modal=True, element_justification='r', keep_on_top=True, size=(500,300), resizable=True )

    while True:
        action, values = popup.read()

        if action == 'Copy':
            pyperclip.copy(values['-DBSELECT-'])
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

def urlFormat(string):
    formatted = re.sub('[^A-Za-z0-9]+', '', string)
    return formatted

# GUI Layout: User input for report IDs, output file selection, function buttons
left_part = [
    [sg.Text('Enter the IDs to search:', font=("Arial Bold", 10))],
    [sg.Multiline(key='-IDS-', expand_x=True, expand_y=True, size=(30,15))],
    [sg.Radio("Expense", "type", key='EXPENSE', default=True),
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

# GUI Layout: Tenant search tab layout
tenant_search = [
    [sg.Text("Search Criteria:"), sg.Input(key='-ENTITYCODE-')],
    [sg.Text('Tenant: '), sg.Text(key='tenant')],
    [sg.Text('Client: '), sg.Text(key='client')],
    [sg.Button('Search'), sg.Button('Open')],
    [sg.Table(values=[], headings=['Tenant', 'Client Name', 'Entity'], key='-SEARCHRESULTS-', justification='center', enable_click_events=True, expand_x=True, enable_events=True, num_rows=15, visible=False)]

]

# GUI Layout: Table/List displaying report IDs NOT found
right_part = [[sg.Table(values=[], headings=['Report IDs Not Found'], 
              auto_size_columns=False, 
              display_row_numbers=False, 
              key='-NOTFOUND-', 
              expand_x=True, 
              expand_y=True,
              enable_events=True)],
    [sg.Push(),
     sg.Button('Copy')]
]





tab_layout = [
    [sg.TabGroup([
        [sg.Tab('Tenant Search', tenant_search),
         sg.Tab('Report Finder', left_part),
         sg.Tab('Results', right_part, visible=False, key='RESULT_TAB')
        ]
    ])]
]
# Create the window
window = sg.Window('Report Finder Tool', tab_layout, resizable=False)

db_report_script = '''select distinct ct_report.REPORT_ID as 'Report ID', CT_JOB_DEFINITION_LANG.NAME as 'Job name', END_TIME as 'Run date', RUN_NUMBER as 'Run number', RECORDS_LOADED as 'Records' from ct_report join ct_journal on ct_report.rpt_key = ct_journal.rpt_key join CT_JOB_RUN on ct_journal.jr_key = CT_JOB_RUN.jr_key join CT_JOB_DEFINITION_LANG on CT_JOB_RUN.jd_key = CT_JOB_DEFINITION_LANG.jd_key where ct_report.REPORT_ID in '''
db_invoice_script = '''select distinct ctp_request.REQUEST_ID as 'Request ID', CT_JOB_DEFINITION_LANG.NAME as 'Job name', END_TIME as 'Run date', RUN_NUMBER as 'Run number', RECORDS_LOADED as 'Records' from ctp_request join ctp_journal on ctp_request.req_key = ctp_journal.req_key join CT_JOB_RUN on ctp_journal.jr_key = CT_JOB_RUN.jr_key join CT_JOB_DEFINITION_LANG on CT_JOB_RUN.jd_key = CT_JOB_DEFINITION_LANG.jd_key where ctp_request.REQUEST_ID in '''
db_report_payment_script = '''SELECT DISTINCT CT_REPORT.REPORT_ID as 'REPORT ID', CT_JOB_DEFINITION_LANG.NAME as 'JOB NAME', END_TIME as 'RUN DATE', RUN_NUMBER as 'RUN NUMBER', RECORDS_LOADED as 'RECORDS' FROM CT_REPORT JOIN CT_EFT_ITEM_PAYEE on CT_REPORT.RPT_KEY = CT_EFT_ITEM_PAYEE.RPT_KEY JOIN CT_EFT_PAYMENT_JOURNAL on CT_EFT_ITEM_PAYEE.EFT_PD_KEY = CT_EFT_PAYMENT_JOURNAL.EFT_PD_KEY JOIN CT_JOB_RUN on CT_EFT_PAYMENT_JOURNAL.JR_KEY = CT_JOB_RUN.JR_KEY JOIN CT_JOB_DEFINITION_LANG on CT_JOB_RUN.JD_KEY = CT_JOB_DEFINITION_LANG.JD_KEY WHERE CT_REPORT.REPORT_ID in '''
db_invoice_payment_script = '''SELECT DISTINCT CTP_REQUEST.REQUEST_ID as 'REQUEST_ID', CT_JOB_DEFINITION_LANG.NAME as 'Job name', END_TIME as 'Run date', RUN_NUMBER as 'Run number', RECORDS_LOADED as 'Records' FROM CTP_REQUEST JOIN CTP_EFT_INVOICE_PAYMENT on CTP_REQUEST.REQ_KEY = CTP_EFT_INVOICE_PAYMENT.REQ_KEY JOIN CT_JOB_RUN on CTP_EFT_INVOICE_PAYMENT.JR_KEY = CT_JOB_RUN.JR_KEY JOIN CT_JOB_DEFINITION_LANG on CT_JOB_RUN.JD_KEY = CT_JOB_DEFINITION_LANG.JD_KEY WHERE CTP_REQUEST.REQUEST_ID in '''
orderby_script = ' order by CT_JOB_DEFINITION_LANG.NAME, RUN_NUMBER'
input_files = None

design_artifacts = {
    'US1': 'https://production-us-8n6ihepc.integrationsuite.cfapps.us10.hana.ondemand.com/shell/',
    'US2': 'https://production-us-2-ctysomoc.integrationsuite.cfapps.us10.hana.ondemand.com/shell/',
    'US3': 'https://production-us-3-fe00z6ig.integrationsuite.cfapps.us10.hana.ondemand.com/shell/',
    'US4': 'https://ns-production-us-4-pn18c7zw.it-cpi019.cfapps.us10-002.hana.ondemand.com/itspaces/shell/',
    'US5': 'https://ns-production-us-5-cmxcebye.integrationsuite.cfapps.us10-002.hana.ondemand.com/shell/',
    'US6': 'https://ns-production-us-6-416051om.integrationsuite.cfapps.us10-002.hana.ondemand.com/shell/',
    'US7': 'https://ns-production-us-7-b5z37r8o.integrationsuite.cfapps.us10-002.hana.ondemand.com/shell/',
    'EMEA1': 'https://ns-production-emea-1-cygt9aoi.integrationsuite.cfapps.eu10-003.hana.ondemand.com/shell/'
}


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
        
        if values['EXPENSE']:
            if values['PAYMENTCHECK']:
                db_popup(ids, db_report_payment_script, "DBSelect for Expense Payments")
            else:
                db_popup(ids, db_report_script, "DBSelect for Expense Reports")
        else:
            if values['PAYMENTCHECK']:
                db_popup(ids, db_invoice_payment_script, "DBSelect for Invoice Payments")
            else:
                db_popup(ids, db_invoice_script, "DBSelect for Invoice Reports")

    
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
        #    window['-NOTFOUND-'].update(values=[[elem] for elem in list(set(ids) - set(ids_found))])
            not_found = list(set(ids) - set(ids_found))
            if(len(not_found) > 0):
                window['-NOTFOUND-'].update(values=not_found)
                window['RESULT_TAB'].update(visible=True)
            else:
                window['RESULT_TAB'].update(visible=False)

            if (len(ids_found) > 0):
                # Write extracted lines to output file selected by user
                with open(output_file, 'w', encoding='utf8') as file:
                    if values['HEADER']:
                        file.write('Header\n')
                    file.writelines(extracted_lines)

            #    result_popup(str(len(list(set(ids_found)))), str(len(window['-NOTFOUND-'].get())), output_file)
                result_popup(str(len(list(set(ids_found)))), str(len(not_found)), output_file)
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
            copyValues += (i + '\n')
        pyperclip.copy(copyValues)

    #======================================================= Tenant Search
    
    if event == 'Search' :
        entity_code = values['-ENTITYCODE-']
    #    search_file = window['-SEARCHFILE-'].get()
        if os.path.exists('./Netsuite Packages.txt') != True:
            sg.popup('Unable to find NetSuite Packages file')
        else:
            search_file = os.getcwd() + '/Netsuite Packages.txt'

            result = []
            new_rows = []
            selected = []
            with open(search_file, 'r') as file:
                packages = file.readlines()
                for package in packages:
                    if entity_code.lower() in package.lower():
                        result.append(package.split(';')[:2])

            if len(result) > 1:
                for x in result:
                    temp = x[1].split(' - ')
                    if 'NetSuite' in temp:
                        temp.remove('NetSuite')
                    if len(temp) == 1:
                        new_rows.append([x[0], temp[0], ''])
                    else:
                        new_rows.append([x[0], temp[0], temp[1]])
                window['-SEARCHRESULTS-'].update(values=new_rows, visible=True)
            elif len(result) == 1 :
                selected = result[0]
                window['tenant'].update(selected[0].strip())
                window['client'].update(selected[1].strip())
                window['-SEARCHRESULTS-'].update(values=new_rows, visible=False)
            else:
                sg.popup("Client not found!")
    
    if '+CLICKED+' in event:
        selected = result[event[2][0]]
        window['tenant'].update(selected[0].strip())
        window['client'].update(selected[1].strip())


    if event == 'Open':
        if window['tenant'].get() != '':
            webbrowser.open(design_artifacts[selected[0]] + 'design/contentpackage/' + urlFormat(selected[1]) + '?section=ARTIFACTS')
            time.sleep(5)
            webbrowser.open(design_artifacts[selected[0]] + 'monitoring/Overview')
        else:
            sg.popup("Please search/select for a client!")

# Close the window
window.close()
