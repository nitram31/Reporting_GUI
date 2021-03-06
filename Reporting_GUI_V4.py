import tkinter
from datetime import datetime
from tkinter import *
from tkinter import filedialog
import pandas as pd
import time
import traceback

"""
Hi, I am Martin, Anna's masters internship student, know that I'm not paid, so I can't assure you that this script
will stand the wrath of time.

This script will be used to parse Excel spreadsheet, if you edit this script at a later date, please don't judge me, 
it was made by a student learning python with the only goal of making something functional.
"""


class ExcelFile:
    def __init__(self, path):
        self.path = path
        self.body_dict = {}
        self.file_dataframe = None
        self.sheet_name = None

    def analyse_file(self, variable):

        self.file_dataframe = pd.read_excel(self.path, sheet_name=self.sheet_name, header=10, parse_dates=True)
        file = self.file_dataframe
        choice = variable.get()
        body_dict = self.body_dict

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
                         'Recourse' | 'SFRO' | 'SFR1' | 'TC' | 'Unsuccessful Search' | '??????' | 'RENEGO' | 'Survey' | \
                         'Measurem. Report' | 'New Site':
                        if not body_dict[next_step]['body_header']:
                            body_dict[next_step]['body_header'] = \
                                ['Site ID', 'Build Job ID (Netsite)', "Blocking Issue", 'Comment']

                body_dict[next_step]['body'].append(self.make_line(body_dict[next_step]['body_header'], i))
        self.body_dict = body_dict

    def make_line(self, header, i):
        line_list = []
        for el in header:
            if isinstance(self.file_dataframe[el][i], pd.Timestamp):
                line_list.append(self.file_dataframe[el][i].date())
            else:
                line_list.append(self.file_dataframe[el][i])
        return line_list

    def output_file(self, choice):
        for name in self.body_dict.keys():
            output_name = name
            body = self.body_dict[name]['body']
            header = self.body_dict[name]['body_header']
            df = pd.DataFrame.from_records(body, columns=header)
            with pd.ExcelWriter('extract' + output_name + "_" + choice + '.xlsx') as writer:
                df.to_excel(writer)

    def get_options(self):
        file = pd.read_excel(self.path, sheet_name=self.sheet_name, header=10)
        option_list = []
        header_list = list(file.columns)
        if 'Project' in header_list:
            for el in file['Project']:
                if el not in option_list and not isinstance(el, float):
                    option_list.append(el)
            return option_list
        return

    def get_sheet_list(self):
        file = pd.read_excel(self.path, sheet_name=None, header=10)
        slide_list = list(file.keys())
        return slide_list

    def set_sheet_name(self, name):
        self.sheet_name = name
        if isinstance(self.sheet_name, tkinter.StringVar):
            self.sheet_name = self.sheet_name.get()


class Interface:

    def __init__(self):
        self.root = Tk()
        root = self.root
        root.resizable(width=False, height=False)
        root.title('Reporting GUI V4')
        self.variable = tkinter.StringVar(root)
        self.variable.set('')
        self.drop_down_menu = OptionMenu(root, self.variable, '')
        self.sheet_name_variable = tkinter.StringVar(root)
        self.sheet_name_variable.set('')
        self.drop_down_menu_slide = OptionMenu(root, self.sheet_name_variable, '')
        self.frame = LabelFrame(root, text="Excel file path")
        self.frame.grid(row=2, column=0, padx=10, pady=50)

        self.frame2 = LabelFrame(root, text="Select Excel file")
        self.frame2.grid(row=2, column=1, padx=10, pady=10)
        self.mypath = Entry(self.frame, width=50)

        self.mypath.grid(row=2, column=0)
        self.drop_down_menu_slide.grid(row=3, column=0)
        self.drop_down_menu.grid(row=4, column=0)

        self.message = StringVar()
        self.mylabel2 = Label(root, textvariable=self.message)

        self.mybutton = Button(root, text="Run scan", command=self.myclick, state="disabled")
        self.mybutton2 = Button(self.frame2, text="Select file", command=self.myfile)
        self.mybutton.grid(row=5, column=0)
        self.mybutton2.grid(row=2, columns=2)
        self.path = ""
        self.file = None
        self.sheet_name = ""
        self.root.mainloop()

    def myclick(self):
        try:
            self.path = self.mypath.get()
            if self.path != "":
                self.file.analyse_file(self.variable)
                choice = self.variable.get()
                self.file.output_file(choice)
                self.message.set("Everything went smoothly, the files should be in the folder " +
                                 "from which you executed the program ;)")
                self.mylabel2.grid(row=6, column=0)
        except Exception as e:
            self.manage_exception(e)

    def myfile(self):
        def OptionMenu_SelectionEvent(event):
            self.sheet_name = event
            self.message.set(f"Scanning {event} for content, please wait")
            self.file.set_sheet_name(event)
            self.show_option()
            time.sleep(4)
            self.message.set("")

        try:
            self.message.set("Please allow up to 1 minute to read the file, depending on its size")
            self.mylabel2.grid(row=6, column=0)
            self.root.fasta_file = filedialog.askopenfilename()
            self.mypath.delete(first=0, last=tkinter.END)
            self.mypath.insert(0, self.root.fasta_file)
            path = self.mypath.get()
            self.file = ExcelFile(path)
            slide_list = self.file.get_sheet_list()
            self.sheet_name_variable.set(slide_list[0])
            self.file.set_sheet_name(self.sheet_name_variable)

            self.drop_down_menu_slide = OptionMenu(self.root, self.sheet_name_variable, *slide_list,
                                                   command=OptionMenu_SelectionEvent)
            self.drop_down_menu_slide.grid(row=3, column=0)
            self.message.set("")

            self.drop_down_menu_slide.update()
            self.show_option()
            self.mybutton.config(state='active')

        except ValueError:
            self.message.set(
                "Something went wrong, the most likely cause for this error is that \nyou selected the wrong "
                "type of file.")

    def show_option(self):

        option_list = self.file.get_options()
        if option_list:
            self.variable.set(option_list[0])
            self.drop_down_menu = OptionMenu(self.root, self.variable, *option_list)
            self.mybutton.config(state='active')
            self.mybutton.update()
        else:
            self.variable.set('')
            self.drop_down_menu = OptionMenu(self.root, self.variable, '')
        self.drop_down_menu.grid(row=4, column=0)

    def manage_exception(self, ex):
        # please remove my email address if you took over this script.
        # but don't hesitate to email me to tell me how much you like my code
        email_address = 'martin.racoupeau@univ-tlse3.fr'
        self.message.set('Something went wrong, please send the errorlog that should have been\n '
                         + 'created in the folder from which you executed the program \nto '
                         + email_address
                         + ' or the person currently maintaining the script.')

        now = datetime.now()
        # dd/mm/YY H:M:S
        dt_string = now.strftime("%d-%m-%Y_%H:%M:%S")
        with open('error_log_' + dt_string, 'w') as file:
            file.write(''.join(traceback.format_tb(ex.__traceback__)) + "\n" + str(ex))


if __name__ == "__main__":
    inter = Interface()
