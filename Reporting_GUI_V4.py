import tkinter
from tkinter import *
from tkinter import filedialog
import pandas as pd
from datetime import datetime
import traceback

"""
Hi, I am Martin, Anna's masters internship student, know that I'm not paid, so I can't assure you that this script
will stand the wrath of time.

This script will be used to parse Excel spreadsheet, if you edit this script at a later date, please don't judge me, 
it was made by a student learning python with the only goal of making something functional.
"""


def analyse_file(path, variable, sheet_name):
    sheet_name = sheet_name.get()
    file = pd.read_excel(path, sheet_name=sheet_name, header=10, parse_dates=True)
    choice = variable.get()

    body_dict = {}


    for i in range(0, len(file)):
        if file['Project'][i] == choice:
            next_step = file['Next Step'][i]
            if next_step not in body_dict.keys():
                body_dict[next_step] = {'body_header': [], 'body': []}

            match next_step:
                case "Permitting":
                    if not body_dict[next_step]['body_header']:
                        body_dict[next_step]['body_header'] = \
                            ['Site ID', 'Build Job ID (Netsite)', "Blocking Issue", 'Comment']

                case "On Hold":
                    if not body_dict[next_step]['body_header']:
                        body_dict[next_step]['body_header'] = \
                            ['Site ID', 'Build Job ID (Netsite)', "Blocking Issue", 'Comment']

                case "Correction BP":
                    if not body_dict[next_step]['body_header']:
                        body_dict[next_step]['body_header'] = \
                            ['Site ID', 'Build Job ID (Netsite)', "Blocking Issue", 'Comment']

                case "BP Signing":
                    if not body_dict[next_step]['body_header']:
                        body_dict[next_step]['body_header'] = \
                            ['Site ID', 'Build Job ID (Netsite)', "Blocking Issue", 'Comment']

                case "BP Application":
                    if not body_dict[next_step]['body_header']:
                        body_dict[next_step]['body_header'] = \
                            ['Site ID', 'Build Job ID (Netsite)', 'Comment']

                case "Draft":
                    if not body_dict[next_step]['body_header']:
                        body_dict[next_step]['body_header'] = \
                            ['Site ID', 'Build Job ID (Netsite)', 'Survey', 'Comment']

                case "NIS":
                    if not body_dict[next_step]['body_header']:
                        body_dict[next_step]['body_header'] = \
                            ['Site ID', 'Build Job ID (Netsite)', 'preNIS ready for QS',
                             'preNIS sent to Provider', 'preNIS approved by provider',
                             'Final NIS ready for QS', 'Final NIS sent to Provider', 'Comment']

                case "PA":
                    if not body_dict[next_step]['body_header']:
                        body_dict[next_step]['body_header'] = \
                            ['Site ID', 'Build Job ID (Netsite)', 'Comment']

                case "Survey":
                    if not body_dict[next_step]['body_header']:
                        body_dict[next_step]['body_header'] = \
                            ['Site ID', 'Build Job ID (Netsite)', 'Survey', 'Comment']

                case "AVOR":
                    if not body_dict[next_step]['body_header']:
                        body_dict[next_step]['body_header'] = \
                            ['Site ID', 'Build Job ID (Netsite)', 'Comment']

                case 'Closed' | 'Dialog' | 'EGT to sign Lease' | 'GA' | 'MBA Analysis' | 'Prep Lease (MV/DBV)' | \
                     'Recourse' | 'SFRO' | 'SFR1' | 'TC' | 'Unsuccessful Search' | 'ΙΡΑ' | 'RENEGO' | 'Survey' | \
                     'Measurem. Report' | 'New Site':
                    if not body_dict[next_step]['body_header']:
                        body_dict[next_step]['body_header'] = \
                            ['Site ID', 'Build Job ID (Netsite)', "Blocking Issue", 'Comment']

            body_dict[next_step]['body'].append(make_line(body_dict[next_step]['body_header'], file, i))

    return body_dict, choice


def make_line(header, file, i):
    line_list = []
    for el in header:
        if isinstance(file[el][i], pd.Timestamp):
            line_list.append(file[el][i].date())
        else:
            line_list.append(file[el][i])
    return line_list


def return_proper_time(date):
    try:
        return str(date)[0:10]
    except IndexError:
        return pd.nan


def output_file(body_dict, choice):
    for name in body_dict.keys():
        output_name = name
        body = body_dict[name]['body']
        header = body_dict[name]['body_header']
        df = pd.DataFrame.from_records(body, columns=header)
        with pd.ExcelWriter('extract' + output_name + "_" + choice + '.xlsx') as writer:
            df.to_excel(writer)


def get_options(path, sheet_name):
    if isinstance(sheet_name, tkinter.StringVar):
        sheet_name = sheet_name.get()

    file = pd.read_excel(path, sheet_name=sheet_name, header=10)
    option_list = []
    for el in file['Project']:
        if el not in option_list and str(type(el)) != "<class 'float'>":
            option_list.append(el)
    return option_list


def get_sheet_list(path):
    file = pd.read_excel(path, sheet_name=None, header=10)
    slide_list = list(file.keys())
    return slide_list


def main():
    # yes, I known there is probably no need for sub-functions had I structured the code better, did I do it? No.
    # Does the code works either way? yes
    def myclick_sub():
        myclick()

    def myfile_sub():
        myfile()

    root = Tk()
    root.resizable(width=False, height=False)
    root.title('Reporting GUI V4')
    variable = tkinter.StringVar(root)
    variable.set('')
    drop_down_menu = OptionMenu(root, variable, '')
    sheet_name_variable = tkinter.StringVar(root)
    sheet_name_variable.set('')
    drop_down_menu_slide = OptionMenu(root, sheet_name_variable, '')
    frame = LabelFrame(root, text="Excel file path")
    frame.grid(row=2, column=0, padx=10, pady=50)

    frame2 = LabelFrame(root, text="Select Excel file")
    frame2.grid(row=2, column=1, padx=10, pady=10)
    mypath = Entry(frame, width=50)

    mypath.grid(row=2, column=0)
    drop_down_menu_slide.grid(row=3, column=0)
    drop_down_menu.grid(row=4, column=0)

    message = StringVar()
    mylabel2 = Label(root, textvariable=message)

    mybutton = Button(root, text="Run scan", command=myclick_sub, state="disabled")
    mybutton2 = Button(frame2, text="Select file", command=myfile_sub)
    mybutton.grid(row=5, column=0)
    mybutton2.grid(row=2, columns=2)

    def show_option(path, sheet_name):
        dropdown_menu = drop_down_menu
        option_list = get_options(path, sheet_name)
        if option_list:
            variable.set(option_list[0])
            dropdown_menu1 = OptionMenu(root, variable, *option_list)
        else:
            variable.set('')
            dropdown_menu1 = OptionMenu(root, variable, '')

        dropdown_menu1.grid(row=4, column=0)
        dropdown_menu.destroy()

    def myclick():
        path = mypath.get()
        if path != "":
            body_dict, choice = analyse_file(path, variable, sheet_name_variable)
            output_file(body_dict, choice)
            button_message = "Everything went smoothly, the files should be in the folder " + \
                             "from which you executed the program ;)"
            mylabel2 = Label(root, text=button_message)
            mylabel2.grid(row=6, column=0)

    def myfile():
        def OptionMenu_SelectionEvent(event):
            show_option(path, event)

        try:
            message.set("Please allow up to 1 minute to read the file, depending on its size")
            mylabel2.grid(row=6, column=0)
            root.fasta_file = filedialog.askopenfilename()
            mypath.delete(first=0, last=tkinter.END)
            mypath.insert(0, root.fasta_file)
            path = mypath.get()
            slide_list = get_sheet_list(path)
            sheet_name_variable.set(slide_list[0])
            dropdown_menu_slide = drop_down_menu_slide
            dropdown_menu_slide1 = OptionMenu(root, sheet_name_variable, *slide_list,
                                              command=OptionMenu_SelectionEvent)
            dropdown_menu_slide1.grid(row=3, column=0)
            show_option(path, sheet_name_variable)
            mybutton = Button(root, text="Run scan", command=myclick, state="active")
            mybutton.grid(row=5, column=0)
            dropdown_menu_slide.destroy()

            root.update()
        except ValueError:
            message.set("Something went wrong, the most likely cause for this error is that \nyou selected the wrong "
                        "type of file.")

    try:
        root.mainloop()
    except Exception as e:
        # please remove my email address if you took over this script.
        # but don't hesitate to send me an email to tell me how much of an abomination you found my code to be
        email_address = 'martin.racoupeau@univ-tlse3.fr'
        message.set('Something went wrong, please send the errorlog that should have been\n created in the '
                    + 'folder from which you executed the program to '
                    + email_address
                    + ' or the person currently maintaining the script.')

        now = datetime.now()
        # dd/mm/YY H:M:S
        dt_string = now.strftime("%d-%m-%Y%H:%M:%S")
        with open('error_log_' + dt_string, 'w') as file:
            file.write(''.join(traceback.format_tb(e.__traceback__)) + "\n" + str(e))


if __name__ == "__main__":
    main()

