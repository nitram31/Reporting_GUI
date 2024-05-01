import tkinter
from datetime import datetime
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import Progressbar

import pandas as pd
import time
import traceback
from threadom import Threadom

"""
Hi, I am Martin, Anna's masters internship student, know that I'm not paid, so I can't assure you that this script
will stand the wrath of time.

This script will be used to parse Excel spreadsheet, if you edit this script at a later date, please don't judge me, 
it was made by a student learning python with the only goal of making something functional. Also I am french. (This seems
like an important piece of info.)

auto-py-to-exe was used to generate the executable, know that you will need to include the openpyxl folder manually, or 
else the user will not be able to launch it. 

debug: 
pyinstaller --noconfirm --onefile --console --name "Reporting_GUI_V4_debug.exe" --add-data "D:/Bureau/Cours/M1/pythonProject/venv3.10/Lib/site-packages/openpyxl;openpyxl/"  "D:/Bureau/Cours/M1/pythonProject/Reporting_GUI/Reporting_GUI_V5.py"
normal: 
pyinstaller --noconfirm --onefile --windowed --name "Reporting_GUI_V4.exe" --add-data "D:/Bureau/Cours/M1/pythonProject/venv3.10/Lib/site-packages/openpyxl;openpyxl/"  "D:/Bureau/Cours/M1/pythonProject/Reporting_GUI/Reporting_GUI_V5.py"
"""


class ExcelFile:
    def __init__(self, path):
        self.path = path
        self.body_dict = {}
        self.file_dataframe = None
        self.sheet_name = None

    def analyse_file(self, variable):
        """
        :param variable:  The name of the project
        """
        self.file_dataframe = pd.read_excel(self.path, sheet_name=self.sheet_name, header=10, parse_dates=True)
        file = self.file_dataframe
        choice = variable.get()
        body_dict = self.body_dict

        for i in range(0, len(file)):
            if file['Project'][i] == choice:
                next_step = file['Next Step'][i]
                if next_step not in body_dict.keys():
                    body_dict[next_step] = {'body_header': [], 'body': []}

                if not body_dict[next_step]['body_header']:  # initialize header
                    match next_step:

                        case "NIS":

                            body_dict[next_step]['body_header'] = \
                                ['Site ID', 'Build Job ID (Netsite)', 'preNIS ready for QS',
                                 'preNIS sent to Provider', 'preNIS approved by provider',
                                 'Final NIS ready for QS', 'Final NIS sent to Provider', 'Comment']

                        case "Survey" | "Draft":

                            body_dict[next_step]['body_header'] = \
                                ['Site ID', 'Build Job ID (Netsite)', 'Survey', 'Comment']

                        case "AVOR" | "PA":

                            body_dict[next_step]['body_header'] = \
                                ['Site ID', 'Build Job ID (Netsite)', 'Comment']

                        case 'Permitting' | 'On Hold' | 'Correction BP' | 'BP Signing' | 'Closed' | 'Dialog' \
                             | 'EGT to sign Lease' | 'GA' | 'MBA Analysis' | 'Prep Lease (MV/DBV)' | 'Recourse' \
                             | 'SFRO' | 'SFR1' | 'TC' | 'Unsuccessful Search' | 'ΙΡΑ' | 'RENEGO' | 'Survey' \
                             | 'Measurem. Report' | 'New Site':

                            body_dict[next_step]['body_header'] = \
                                ['Site ID', 'Build Job ID (Netsite)', "Blocking Issue", 'Comment']

                body_dict[next_step]['body'].append(self.make_line(body_dict[next_step]['body_header'], i))
        self.body_dict = body_dict

    def make_line(self, header, i):
        """
        Used to make the list of line that will be inserted in the dictionary, also used to properly format dates

        :param header: header of the project selected
        :param i: variable incremented in the for loop of analyse_file()
        :return: the list containing the lines relevant to the selected project
        """
        line_list = []
        for el in header:
            if isinstance(self.file_dataframe[el][i], pd.Timestamp):
                line_list.append(self.file_dataframe[el][i].date())
            else:
                line_list.append(self.file_dataframe[el][i])
        return line_list

    def check_name(self, name):
        """
        Used to make sure there are no windows forbidden characters in the name

        :param name:  variable containing next_step
        :return: the choice
        """
        if '/' in name:  # recursive because it can be, and its fancy
            sub = name.index("/")
            name_corrected = name[:sub] + '-' + name[sub + 1:]
            return self.check_name(name_corrected)
        return name

    def output_file(self, choice, directory):
        """
        It outputs files, I mean what could you expect it to do?

        :param choice: name of the project
        :param directory: Empty by default, contains the path to the custom output directory
        """
        for name in self.body_dict.keys():
            output_name = name
            body = self.body_dict[name]['body']
            header = self.body_dict[name]['body_header']
            df = pd.DataFrame.from_records(body, columns=header)
            file_name = directory + 'extract' + self.check_name(output_name) + "_" + choice + '.xlsx'
            with pd.ExcelWriter(file_name) as writer:
                df.to_excel(writer)

    def get_options(self):
        """
        Gets the different projects available from a selected sheet

        """
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
        """
        Am I not naming my function clearly enough? It looks at a file, then it gives you the list of sheets it contains

        :return: The names of the sheets in the file
        """
        file = pd.read_excel(self.path, sheet_name=None, header=10)
        slide_list = list(file.keys())
        return slide_list

    def set_sheet_name(self, name):
        """
        Used to set the sheet name in the class when a sheet name is selected in the Interface

        :param name: name of the sheet
        """
        self.sheet_name = name
        if isinstance(self.sheet_name, tkinter.StringVar):
            self.sheet_name = self.sheet_name.get()


class Interface:
    """
    Ohh boy. This thing is my frankenstein monster, it's basically my first python project ever but on its 7th version:
    its weird, but it works.
    """

    def __init__(self):
        self.root = Tk()
        root = self.root
        root.resizable(width=False, height=False)
        root.title('Reporting GUI V5')
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
        self.mybutton = Button(root, text="Run scan", command=self.myclickthread, state="disabled")
        self.mybutton2 = Button(self.frame2, text="Select file", command=self.myfilethread)
        self.mybutton.grid(row=5, column=0)
        self.mybutton2.grid(row=2, columns=2)

        self.menu = tkinter.Menu(root)
        self.menu1 = tkinter.Menu(root)
        self.menu.add_cascade(label="File options", menu=self.menu1)
        self.menu1.add_command(label="Change Output directory", command=self.change_output_directory)
        self.root.config(menu=self.menu, width=200)

        self.path = ""
        self.output_directory = ""
        self.file = None
        self.sheet_name = ""
        self.root.mainloop()

    def change_output_directory(self):
        self.output_directory = filedialog.askdirectory() + '/'

    @staticmethod
    def cut_string(string):
        string_formatted = string
        for i in range(1, len(string)):
            if i % 64 == 0:
                j = i
                while string_formatted[j] != " " and j < len(string) - 1:
                    j += 1
                if j != " ":
                    while string_formatted[j] != " ":
                        j -= 1
                string_formatted = string_formatted[:j] + '\n' + string_formatted[j + 1:]
        return string_formatted

    def myclick(self):
        try:
            self.path = self.mypath.get()
            if self.path != "":
                self.file.analyse_file(self.variable)
                choice = self.variable.get()
                self.file.output_file(choice, self.output_directory)
                message = "Everything went smoothly. By default, the files should be " \
                          "in the folder from which you executed the program ;)"
                self.message.set(self.cut_string(message))
                self.mylabel2.grid(row=6, column=0)
        except Exception as e:
            self.manage_exception(e)

    def myclickthread(self):
        threadom = Threadom(self.root)
        threadom.run(self.myclick())

    def myfilethread(self):
        threadom = Threadom(self.root)
        threadom.run(self.myfile())

    def myfile(self):
        def OptionMenu_SelectionEvent(event):
            self.sheet_name = event
            self.message.set(f"Scanning {event} for content, please wait")
            self.file.set_sheet_name(event)
            self.show_option()
            time.sleep(4)
            self.message.set("")
            if self.file:
                self.file.body_dict = {}

        try:
            message = "Please allow up to 1 minute to read the file, depending on its size"
            self.message.set(self.cut_string(message))
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
            message = "Something went wrong, the most likely cause for this error is that you selected the " \
                      "wrong type of file."
            self.message.set(self.cut_string(message))

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
        email_address = 'martin.racoupeau@gmail.com'
        message = 'Something went wrong, please send the errorlog that should have been created in the ' \
                  'folder from which you executed the program to ' \
                  + email_address \
                  + ' or the person currently maintaining the script.'
        self.message.set(self.cut_string(message))

        now = datetime.now()
        # dd/mm/YY H-M-S
        dt_string = now.strftime("%d-%m-%Y_%H-%M-%S")
        with open('error_log_' + dt_string, 'w') as file:
            file.write(''.join(traceback.format_tb(ex.__traceback__)) + "\n" + str(ex))


if __name__ == "__main__":
    inter = Interface()
