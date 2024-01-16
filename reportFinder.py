import PySimpleGUI as sg
import pyperclip


def result_popup(found, notfound):
    sg.Window('Results', [[sg.Text('Reports Found: ' + found, size=(30, 1))],
                          [sg.Text('Reports Not Found: ' + notfound, size=(30, 1))],
                        [sg.Button('Ok')]], modal=True, element_justification='c', keep_on_top=True).read(close=True)

# GUI: User input for report IDs, output file selection, function buttons
left_part = [
    [sg.Text('Enter the IDs to search:', font=("Arial Bold", 10))],
    [sg.Multiline(key='-IDS-', expand_x=True, expand_y=True, size=(30,15))],
    [sg.Text('Select Extracts to Search:', font=("Arial Bold", 10))],
    [sg.Text('Extracts Selected: ', font=("Arial Bold", 10)), sg.Text(key='-INPUT-'), sg.Push(), sg.Button('Select')],
    [sg.Text('Select the Output File to Store Results:', font=("Arial Bold", 10) )],
    [sg.Input(key='-OUTPUT_FILE-', disabled=True), sg.FileBrowse()],
    [sg.Button('Search and Extract'),
     sg.Button('Exit')]
]

# GUI: Table/List displaying report IDs found
#center_part = [
#    [sg.Text('Report IDs Found', font=("Arial Bold", 12))],
#    [sg.Table(values=[], headings=['Report ID'], auto_size_columns=True, display_row_numbers=False, key='-FOUND-', expand_x=True, expand_y=True)]
#]

# GUI: Table/List displaying report IDs NOT found
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
    [sg.Push(), sg.Button('Copy')]
]


layout = [
    [sg.Column(left_part, vertical_alignment='top'),
    sg.VSeparator(),
#    sg.Column(center_part, vertical_alignment='top'),
#    sg.VSeparator(),
    sg.Column(right_part, vertical_alignment='top')]
]

# Create the window
window = sg.Window('Report Finder Tool', layout, resizable=True)

# Event loop to process GUI events
while True:
    event, values = window.read()
    
    # Exit the program if the window is closed or 'Exit' button is clicked
    if event == sg.WINDOW_CLOSED or event == 'Exit':
        break

    if event == 'Select':
        input_files = sg.popup_get_file('Select Extracts to Search', multiple_files=True)
        if input_files:
            window['-INPUT-'].update(len(input_files.split(';')))
        else:
             window['-INPUT-'].update('')

    # Perform the search and extraction when 'Search and Extract' button is clicked
    if event == 'Search and Extract':
    #    ids = values['-IDS-'].split(',')
        ids = list(set(values['-IDS-'].split('\n')))
    #   input_file = values['-INPUT_FILE-']
        
        output_file = values['-OUTPUT_FILE-']
        
        files = input_files.split(';')
        extracted_lines = []
        ids_found = []

        for file_path in files:
            with open(file_path, 'r') as file:
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

        # Update table with report IDs not found
        window['-NOTFOUND-'].update(values=[[elem] for elem in list(set(ids) - set(ids_found))])

        # Write extracted lines to output file selected by suer
        with open(output_file, 'w') as file:
            file.write('Header\n')
            file.writelines(extracted_lines)

        if (len(ids_found)):
            result_popup(str(len(list(set(ids_found)))), str(len(window['-NOTFOUND-'].get())))
        else:
            sg.popup("No reports found!")

# Copy report IDs from table
    if event == 'Copy':
        table_values = window['-NOTFOUND-'].get()
        copyValues = ""
        for i in table_values:
            copyValues += (i[0] + '\n')
        pyperclip.copy(copyValues)

# DEBUG: Test button
    #if event == 'Test':
    #    result_popup('4', '6')
    #     sg.popup(
    #            'Reports Found: ' + str(len([3,6])),
    #            '-'*58,
    #            'Reports Not Found: ' + str(len(window['-NOTFOUND-'].get())),
    #            title='Results',
    #            button_justification='c')

# Close the window
window.close()
