import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter.ttk import Progressbar
import tkinter.scrolledtext as tkst
import ttk
from PIL import ImageTk, Image

import time
import os
import gc
import re

import matplotlib.pyplot as plt
import numpy as np

import datetime
from datetime import timedelta
from dateutil import parser

import GSheet
import GMail
import GDrive


class AutoScrollbar(tk.Scrollbar):
    '''
    A scrollbar that hides itself if it's not needed.
    Only works if you use the grid geometry manager!
    '''

    def set(self, lo, hi):
        if float(lo) <= 0.0 and float(hi) >= 1.0:
            # grid_remove is currently missing from Tkinter!
            self.tk.call('grid', 'remove', self)
        else:
            self.grid()
        tk.Scrollbar.set(self, lo, hi)

    def pack(self, **kw):
        raise tk.TclError('cannot use pack with this widget')

    def place(self, **kw):
        raise tk.TclError('cannot use place with this widget')


class ScrollFrame:
    '''
    A Scrollbar that incorporates a frame inside it

    '''

    def __init__(self, master):

        # creates and grids the autoscrollbars
        self.vscrollbar = AutoScrollbar(master)  # AutoScrollbar(x)
        self.vscrollbar.grid(row=0, column=1, sticky=tk.NS)
        self.hscrollbar = AutoScrollbar(master, orient=tk.HORIZONTAL)
        self.hscrollbar.grid(row=1, column=0, sticky=tk.EW)
        ttk.Sizegrip(master).grid(column=1, row=1, sticky=('S', 'E'))

        # creates a canvas for the scrollbars to scrollthrough
        self.canvas = tk.Canvas(
            master,
            yscrollcommand=self.vscrollbar.set,
            xscrollcommand=self.hscrollbar.set,
            width=700,
            height=200)
        self.canvas.grid(row=0, column=0, sticky=tk.NSEW)

        # links scrollbars to canvas
        self.vscrollbar.config(command=self.canvas.yview)
        self.hscrollbar.config(command=self.canvas.xview)

        # make the canvas expandable
        master.grid_rowconfigure(0, weight=1)
        master.grid_columnconfigure(0, weight=1)

        # create canvas contents. The frame allows widgets to be placed inside
        self._active_area = None
        self.frame = tk.Frame(self.canvas)
        self.frame.rowconfigure(1, weight=1)
        self.frame.columnconfigure(1, weight=1)
        self.frame.bind('<Configure>', self.reset_scrollregion)

    def update(self):
        ''' Updates scrollregion based on convas contents'''
        self.canvas.create_window(0, 0, anchor=tk.NW, window=self.frame)
        self.frame.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox('all'))

    def reset_scrollregion(self, event):
        '''Does the same but is triggered according to configure'''
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))

    def autosize(self):
        '''If the canvas is smaller than the scrollregion, make the canvas size
        equal to scrollregion'''
        if self.frame.winfo_reqwidth() != self.canvas.winfo_width():
            # update the canvas's width to fit the inner frame
            self.canvas.config(width=self.frame.winfo_reqwidth())
        if self.frame.winfo_reqheight() != self.canvas.winfo_height():
            # update the canvas's width to fit the inner frame
            self.canvas.config(height=self.frame.winfo_reqheight())

    def resize(self, l, w):
        '''Changes size of canvas based on input'''
        self.canvas.config(width=w)
        self.canvas.config(height=l)


class Form2Email:
    '''Main class for Form2Email application. Contains subclasses and functions
    that are vital to its operations, and also contains the gui'''

    class LoadWindow():
        '''Displayed when program is first booted up for aesthetics'''

        def __init__(self, cls):
            self.cls = cls
            cls.top = tk.Toplevel()
            cls.top.title("Form2Email")
            # sets window size and centers it on screen
            self.cls.top.geometry('250x300+' +
                                  str((self.cls.root.winfo_screenwidth() //
                                       2) -
                                      125) +
                                  '+' +
                                  str((self.cls.root.winfo_screenheight() //
                                       2) -
                                      150))
            # sets icon
            self.cls.top.wm_iconbitmap(
                os.path.abspath(
                    os.path.abspath(
                        os.path.join(
                            __file__,
                            os.pardir))) +
                '\\icon.ico')

            # creates main image in the window
            self.path = (
                os.path.abspath(
                    os.path.abspath(
                        os.path.join(
                            __file__,
                            os.pardir))) +
                '\ICON.png')
            self.img = Image.open(self.path)
            self.img = self.img.resize((250, 250), Image.ANTIALIAS)
            self.img = ImageTk.PhotoImage(self.img)
            self.panel = tk.Label(self.cls.top, image=self.img)
            self.panel.pack(side="top", fill="both", expand="yes")
            # creates loading bar and displays information
            self.progress = Progressbar(
                self.cls.top,
                orient='horizontal',
                length=100,
                mode='determinate')
            self.progress.pack(side='bottom', fill='x', expand='yes')
            for i in range(100):
                self.progress.step(1)
                self.progress.update()
                # Busy-wait
                time.sleep(.015)
            self.stop()

        def stop(self):
            '''Used to close the loading window'''
            self.cls.top.destroy()

    class RunWindow():
        '''Displayed when the program is running to remind the user that it is running'''

        def __init__(self, cls):

            self.cls = cls
            cls.top = tk.Toplevel()
            cls.top.title("Form2Email")
            cls.top.protocol("WM_DELETE_WINDOW", self.stop)

            cls.msg = tk.Message(
                cls.top, text="Form2Email is currently Running")
            cls.msg.pack()

            cls.button = tk.Button(cls.top, text="Stop", command=self.stop)
            cls.button.pack()

            print('Run Window Initiated.')

        def stop(self):
            '''Used to close window'''
            if messagebox.askokcancel(
                    'Stop', 'Are you sure you want to stop Form2Email?'):
                self.cls.top.destroy()
                # stops running tasks
                self.cls.running = False
                # unhides main window
                self.cls.root.deiconify()

    class TimeInterval:
        '''One of three classes for the email sending options. This class is responsible
        for sending an email according to a time interval set by the user'''

        def __init__(self, cls, group_n, time, interval, email_addresses, subject_entry,
                     html_entry, d):

            self.cls = cls
            self.group_number = group_n

            # error checking for user input
            try:
                self.time = parser.parse(time)
            except BaseException:
                self.create_error(
                    "Could Not Parse Time for TimeInterval in Group" + str(self.group_number))

            # error checking for user input
            try:
                self.regex = re.compile(
                    r'((?P<days>\d+?)d)?((?P<hours>\d+?)hr)?((?P<minutes>\d+?)m)?((?P<seconds>\d+?)s)?')
            except BaseException:
                self.create_error(
                    "Could Not Parse Interval for TimeInterval in Group" + str(self.group_number))

            self.interval = self.get_timedelta(interval)
            self.email_addresses = email_addresses
            self.subject = subject_entry
            self.htmlbody = html_entry
            self.data = d

            print('TimeInterval Initiated...')
            print('Email will be sent at ', self.time, 'with a ', self.interval,
                  'delay')

        def check(self, cls, d):
            '''Checks to see if email sending conditions are met. If they are,
            it sends an email'''
            # takes the difference of the time of the next email and current
            # time
            time_difference = self.time - datetime.datetime.now()
            print('TimeInterval: Time until next email: ',
                  timedelta.total_seconds(time_difference))
            # if that time difference is less than zero meaning that the time has
            # passed, the program will send an email
            if timedelta.total_seconds(time_difference) < 0:
                print('Time is up. Sending Email...')
                # adds the time interval to the time when the email was supposed
                # to be sent
                self.time += self.interval
                self.data = d
                cls.email_command_execution(self.htmlbody, self.data)
                self.send_email(cls)

        def get_timedelta(self, time_str):
            '''Creates a timedelta from a time interval'''
            parts = self.regex.match(time_str)
            if not parts:
                return
            parts = parts.groupdict()
            time_params = {}
            for (name, param) in parts.items():
                if param:
                    time_params[name] = int(param)
            return timedelta(**time_params)

        def send_email(self, cls):
            '''Sends Email'''
            html, attach = cls.email_command_execution(
                self.htmlbody,
                self.data,
                self.time)
            GMail.SendMessage(
                'me',
                self.email_addresses,
                self.subject,
                html,
                '', attach)

    class Keyword:
        '''One of three classes for the email sending options. This class is responsible
        for sending an email according to a keyword in set by the user'''

        def __init__(self, cls, group_n, sheet_identifier, keyword, email_addresses,
                     subject_entry, html_entry, d):

            self.cls = cls
            self.group_number = group_n
            self.sheet_number, self.special_column = self.identify(
                sheet_identifier)
            self.keyword = keyword
            self.email_addresses = email_addresses
            self.subject = subject_entry
            self.htmlbody = html_entry
            self.data = d
            self.old_rows = len(self.data[self.sheet_number])
            print('Keyword Initiated...')
            print(
                'Email will be sent when ',
                self.keyword,
                'is seen in ',
                self.special_column,
                'of Sheet',
                self.sheet_number)

        def identify(self, identifier):
            '''Takes the sheet identifier from the user, and splits it into
            sheet number and column'''
            for x in range(0, len(identifier)):
                if identifier[x] == '\\':
                    if identifier[x + 1:].isdigit():
                        return int(identifier[:x]) - 1, identifier[x + 1:]
                    return int(
                        identifier[:x]) - 1, self.cls.column_id_to_int(identifier[x + 1:])
            # if there is no \, it returns an error
            self.create_error(
                "Incorrect Sheet Identifier for Keyword in Group" + str(self.group_number))

        def check(self, cls, d):
            '''Checks to see if email sending conditions are met. If they are,
            it sends an email'''
            self.data = d
            self.new_rows = len(self.data[self.sheet_number])
            # checks to see if there is a new row
            if self.new_rows > self.old_rows:
                print('Keyword: New Row Found. Testing Keyword...')
                # if keyword matches specific column on specific sheet, send
                # email
                if self.data[self.sheet_number -
                             1][len(self.data[self.sheet_number -
                                              1]) -
                                1][self.special_column -
                                   1] == self.keyword:
                    print('Keyword Match! Sending Email...')
                    self.old_rows = self.new_rows
                    cls.email_command_execution(self.htmlbody, self.data)
                    self.send_email(cls)
                else:
                    print('Keyword Mismatch')

        def send_email(self, cls):
            html, attach = cls.email_command_execution(
                self.htmlbody,
                self.data)
            GMail.SendMessage(
                'me',
                self.email_addresses,
                self.subject,
                html,
                '', attach)

    class Response:
        def __init__(self, cls, sheet_number, interval, email_addresses,
                     subject_entry, html_entry, d):
            try:
                self.sheet_number = int(sheet_number) - 1
            except BaseException:
                self.create_error(
                    "Incorrect SheetNumber for Response# in Group" + str(self.group_number))
            try:
                self.interval = int(interval)
            except BaseException:
                self.create_error(
                    "Incorrect interval for Response# in Group" + str(self.group_number))
            self.email_addresses = email_addresses
            self.subject = subject_entry
            self.htmlbody = html_entry
            self.data = d
            self.old_rows = len(self.data[self.sheet_number])
            print('Response# Initiated...')
            print(
                'Email will be sent for every ',
                self.interval,
                'responses on Sheet',
                self.sheet_number)

        def check(self, cls, d):
            self.data = d
            self.new_rows = len(self.data[self.sheet_number])
            if self.new_rows > self.old_rows:
                print('Response: New Row Found. Testing Interval...')
                if self.interval % self.new_rows == 0:
                    print('Interval satisfied. Sending Email...')
                    self.old_rows = self.new_rows
                    cls.email_command_execution(self.htmlbody, self.data)
                    self.send_email(cls)

        def send_email(self, cls):
            html, attach = cls.email_command_execution(
                self.htmlbody,
                self.data)
            GMail.SendMessage(
                'me',
                self.email_addresses,
                self.subject,
                html,
                '', attach)

    def __init__(self):
        self.sheets = []
        self.GSheetURLEntries = []
        self.GSheetTitleEntries = []
        self.groups = []
        self.email_entries = []
        self.subject_entries = []
        self.send_method = []
        self.send_entries = []
        self.send_entries2 = []
        self.html_entries = []
        self.sheet_data = []
        self.running = True
        self.savefilename = None
        self.check_dirs()
        gc.enable()

        self.root = tk.Tk()
        self.root.title('Form2Email')
        self.root.geometry('250x100+' +
                           str((self.root.winfo_screenwidth() //
                                2) -
                               125) +
                           '+' +
                           str((self.root.winfo_screenheight() //
                                2) -
                               50))
        self.root.wm_iconbitmap(
            os.path.abspath(
                os.path.abspath(
                    os.path.join(
                        __file__,
                        os.pardir))) +
            '\\icon.ico')

        self.main = ScrollFrame(self.root)

        self.menus = tk.Menu(self.root)

        self.file_menu = tk.Menu(self.root)
        self.file_menu.add_command(label='Open', command=self.open_file)
        self.file_menu.add_command(label='Save', command=self.save_file)
        self.file_menu.add_command(label='SaveAs', command=self.save_as_file)
        self.file_menu.add_command(label='Close', command=self.close)

        self.edit_menu = tk.Menu(self.root)
        self.edit_menu.add_command(label='LoadData', command=self.load_data)
        self.edit_menu.add_command(
            label='SetCredentials',
            command=self.set_credentials)
        self.edit_menu.add_command(label='Options')
        self.edit_menu.add_command(label='GroupWizard')

        self.run_menu = tk.Menu(self.root)
        self.run_menu.add_command(label='Run', command=self.run)
        self.run_menu.add_command(label='RunBackground')
        self.run_menu.add_command(label='ForceEmail', command=self.force_email)

        self.help_menu = tk.Menu(self.root)
        self.help_menu.add_command(label='Help')
        self.help_menu.add_command(label='About', command=self.about)

        self.menus.add_cascade(label='File', menu=self.file_menu)
        self.menus.add_cascade(label='Edit', menu=self.edit_menu)
        self.menus.add_cascade(label='Run', menu=self.run_menu)
        self.menus.add_cascade(label='Help', menu=self.help_menu)

        self.root.config(menu=self.menus)

        #______________________________________________________________________
        self.main_book = ttk.Notebook(self.main.frame)
        self.main_book.grid()
        self.main_book.bind("<<NotebookTabChanged>>", self.main.update())

        self.sheet_page = ttk.Frame(self.main_book)
        self.main_book.add(self.sheet_page, text='Sheets')

        self.data_page = ttk.Frame(self.main_book)
        self.main_book.add(self.data_page, text='Data')

        self.group_page = ttk.Frame(self.main_book)
        self.main_book.add(self.group_page, text='Groups')

        self.sheet_book = ttk.Notebook(self.sheet_page)
        self.sheet_book.grid()
        self.sheet_book.bind('<Button-3>', self.delete_sheet)
        self.sheet_book.bind('<Button-1>', self.check_create_sheet)

        self.add_page = ttk.Frame(self.sheet_book)
        self.sheet_book.add(self.add_page, text='+')
        self.sheets.append(self.add_page)

        self.data_book = ttk.Notebook(self.data_page)
        self.data_book.grid()

        self.group_book = ttk.Notebook(self.group_page)
        self.group_book.grid()
        self.group_book.bind('<Button-3>', self.delete_group)
        self.group_book.bind('<Button-1>', self.check_create_group)

        self.add_page = ttk.Frame(self.group_book)
        self.group_book.add(self.add_page, text='+')
        self.groups.append(self.add_page)

        self.main.update()
        self.main.autosize()
        self.root.protocol('WM_DELETE_WINDOW', self.close)

        self.loadwindow = self.LoadWindow(self)
        self.root.mainloop()

    def check_create_sheet(self, event):
        ''' Creates a new spreadsheet entry form
        in the application. Function is triggered by clicking on
        the plus tab of the sheets tab
        '''

        clicked_tab = self.sheet_book.tk.call(
            self.sheet_book._w, 'identify', 'tab', event.x, event.y)

        if clicked_tab == (len(self.sheets) - 1):
            self.create_sheet()

    def create_sheet(self):

        current_sheet = tk.StringVar()
        current_sheet.set('Sheet' + str(len(self.sheets)))

        frame = tk.Frame(self.sheet_book)
        self.sheet_book.insert(
            (len(
                self.sheets) - 1),
            frame,
            text=current_sheet.get())
        self.sheets.insert(len(self.sheets) - 1, frame)

        frame.GSheetURLLabel = tk.Label(frame, text='GSheetURL:')
        frame.GSheetURLLabel.grid()

        frame.GSheetURLEntry = tk.Entry(frame)
        frame.GSheetURLEntry.grid(row=0, column=1)
        self.GSheetURLEntries.append(frame.GSheetURLEntry)

        frame.GSheetTitleLabel = tk.Label(frame, text='Sheet Title:')
        frame.GSheetTitleLabel.grid(row=1, column=0)

        frame.GSheetTitleEntry = tk.Entry(frame)
        self.GSheetTitleEntries.append(frame.GSheetTitleEntry)
        frame.GSheetTitleEntry.grid(row=1, column=1)

        self.main.update()
        self.main.autosize()

    def delete_sheet(self, event):

        if messagebox.askyesno(
            'Confirmation',
                'Are you sure that you want delete this sheet?'):
            clicked_tab = self.sheet_book.tk.call(
                self.sheet_book._w, 'identify', 'tab', event.x, event.y)
            print(clicked_tab)
            if clicked_tab != len(self.sheets) - 1:
                self.sheet_book.forget(clicked_tab)
                del self.sheets[clicked_tab]
                del self.GSheetURLEntries[clicked_tab]
                del self.GSheetTitleEntries[clicked_tab]
                for x in range(len(self.sheets) - 1):
                    rename = tk.StringVar()
                    rename.set('Sheet' + str(x + 1))
                    if x != len(self.sheets) - 1:
                        self.sheet_book.tab(x, text=rename.get())

    def change_var(self, variable, frame):

        if variable == 'TimeInterval':
            frame.send_label.config(text='Time:')
            frame.send_label2.config(text='Interval: ')
        elif variable == 'Keyword':
            frame.send_label.config(text='SheetIdentifier: ')
            frame.send_label2.config(text='Keyword: ')
        elif variable == 'Response#':
            frame.send_label.config(text='SheetNumber: ')
            frame.send_label2.config(text='Interval: ')

    def check_create_group(self, event):
        clicked_tab = self.group_book.tk.call(
            self.group_book._w, 'identify', 'tab', event.x, event.y)

        if clicked_tab == (len(self.groups) - 1):
            self.create_group()

    def create_group(self):
        group_counter = tk.StringVar()
        group_counter.set('Group' + str(len(self.groups)))

        frame = ttk.Frame(self.group_book)
        self.group_book.insert(
            (len(self.groups) - 1),
            frame,
            text=group_counter.get())
        self.groups.insert(len(self.groups) - 1, frame)

        frame.emailAddressesLabel = tk.Label(
            frame, text='Email Addresses: ')
        frame.emailAddressesLabel.grid()

        frame.email_entry = tk.Entry(frame)
        frame.email_entry.grid(row=0, column=1)
        self.email_entries.append(frame.email_entry)

        frame.subject_label = tk.Label(frame, text='Subject: ')
        frame.subject_label.grid(row=1)

        frame.subject_entry = tk.Entry(frame)
        frame.subject_entry.grid(row=1, column=1)
        self.subject_entries.append(frame.subject_entry)

        frame.emailOccuranceLabel = tk.Label(
            frame, text='Email Occurance: ')
        frame.emailOccuranceLabel.grid(row=2)

        frame.variable = tk.StringVar(frame)
        frame.variable.set('')  # default value
        self.send_method.append(frame.variable)
        frame.variable.trace(
            'w',
            lambda name,
            index,
            mode,
            variable=frame.variable: self.change_var(
                frame.variable.get(),
                frame))

        self.send_options = tk.OptionMenu(
            frame, frame.variable, 'TimeInterval', 'Keyword', 'Response#')
        self.send_options.grid(row=2, column=1)

        frame.send_label = tk.Label(frame, text='')
        frame.send_label.grid(row=3)

        frame.send_entry = tk.Entry(frame)
        frame.send_entry.grid(row=3, column=1)
        self.send_entries.append(frame.send_entry)

        frame.send_label2 = tk.Label(frame, text='')
        frame.send_label2.grid(row=4)

        frame.send_entry2 = tk.Entry(frame)
        frame.send_entry2.grid(row=4, column=1)
        self.send_entries2.append(frame.send_entry2)

        frame.html_format_label = tk.Label(frame, text='Html: ')
        frame.html_format_label.grid(row=5)

        frame.html_email = tkst.ScrolledText(frame, height=20, width=30)
        frame.html_email.grid(row=6, column=0, columnspan=2, rowspan=2)
        self.html_entries.append(frame.html_email)

        self.main.update()
        self.main.autosize()

    def delete_group(self, event):

        if messagebox.askyesno(
            'Confirmation',
                'Are you sure that you want delete this group?'):
            clicked_tab = self.group_book.tk.call(
                self.group_book._w, 'identify', 'tab', event.x, event.y)

            if clicked_tab != len(self.groups) - 1:
                self.group_book.forget(clicked_tab)
                del self.groups[clicked_tab]
                del self.email_entries[clicked_tab]
                del self.subject_entries[clicked_tab]
                del self.send_method[clicked_tab]
                del self.send_entries[clicked_tab]
                del self.send_entries2[clicked_tab]
                del self.html_entries[clicked_tab]
                for x in range(len(self.groups) - 1):
                    rename = tk.StringVar()
                    rename.set('Group' + str(x + 1))
                    if x != len(self.groups) - 1:
                        self.group_book.tab(x, text=rename.get())

    def load_data(self):

        for widget in self.data_book.winfo_children():
            widget.destroy()

        self.sheet_data = []
        datacounter = 1
        for x in range(0, len(self.GSheetURLEntries)):

            currentData = tk.StringVar()
            currentData.set('Data' + str(datacounter))

            frame = ttk.Frame(self.data_book)
            self.data_book.add(frame, text=currentData.get())

            scroll = ScrollFrame(frame)

            temp = GSheet.getData(self.GSheetURLEntries[x].get(),
                                  self.GSheetTitleEntries[x].get())

            self.sheet_data += [temp]

            self.xcounter = 0
            self.ycounter = 0
            for x in range(0, len(temp)):
                for y in range(0, len(temp[x])):
                    if (x == 0):
                        self.label = tk.Label(
                            scroll.frame,
                            text=self.int_to_column_id(
                                self.ycounter).upper(),
                            borderwidth=1,
                            width=25,
                            height=3,
                            wraplength=150,
                            relief='groove')
                        self.label.grid(row=x, column=y, sticky='NSEW')
                        self.ycounter += 1
                    if (y == 0):
                        self.label = tk.Label(
                            scroll.frame,
                            text=self.xcounter,
                            borderwidth=1,
                            width=25,
                            height=3,
                            wraplength=150,
                            relief='groove')
                        self.label.grid(row=x, column=y, sticky='NSEW')
                        self.xcounter += 1

                    self.label = tk.Label(
                        scroll.frame,
                        text=temp[x][y],
                        borderwidth=1,
                        width=25,
                        height=3,
                        wraplength=150,
                        relief='groove')
                    self.label.grid(row=x + 1, column=y + 1, sticky='NSEW')

            self.label = tk.Label(
                scroll.frame,
                text=self.int_to_column_id(
                    self.ycounter).upper(),
                borderwidth=1,
                width=25,
                height=3,
                wraplength=150,
                relief='groove')
            self.label.grid(row=0, column=len(temp[0]), sticky='NSEW')

            self.label = tk.Label(
                scroll.frame,
                text=self.xcounter,
                borderwidth=1,
                width=25,
                height=3,
                wraplength=150,
                relief='groove')
            self.label.grid(row=len(temp), column=0, sticky='NSEW')

            datacounter += 1
            scroll.update()

            self.main.update()
            self.main.autosize()

    def get_data(self, GSheetURLEntries, GSheetTitleEntries):

        data = []
        for x in range(0, len(GSheetURLEntries)):
            data += [GSheet.getData(GSheetURLEntries[x].get(),
                                    GSheetTitleEntries[x].get())]
        return data

    def int_to_column_id(self, num):
        ''' Converts any positive integer to Base26(letters only) with no 0th case.
        Useful for applications such as spreadsheet columns to determine which
        Letterset goes with a positive integer.
        '''
        if num <= 0:
            return ''
        elif num <= 26:
            return chr(96 + num)
        else:
            return self.int_to_column_id(
                int((num - 1) / 26)) + chr(97 + (num - 1) % 26)

    def column_id_to_int(self, string):
        ''' Converts a string from Base26(letters only) with no 0th case to a
        positive integer. Useful for figuring out column numbers from letters so
        that they can be called from a list.
        '''
        string = string.lower()
        if string == ' ' or len(string) == 0:
            return 0
        if len(string) == 1:
            return ord(string) - 96
        else:
            return self.column_id_to_int(string[1:]) \
                + (26**(len(string) - 1)) \
                * (ord(string[0]) - 96)

    def email_command_execution(self, string, data=[], time=None):
        attachment_list = []
        while '\\\\' in string:
            x, y = self.find_next_command(string, 0)
            string, attachments = self.execute_command(
                x, y, string, data, time)
            if attachments is not None:
                for x in attachments:
                    attachment_list += [x]
        return string, attachment_list

    def find_next_command(self, string, start):
        first = -1
        for x in range(start, len(string)):
            if string[x:x+2] == '\\\\' and first == -1:
                first = x
            if (first != -1 and
                    not(string[x].isalpha()
                        or string[x].isdigit()
                        or string[x] == '\\'
                        or string[x] == '-'
                        or string[x] == ':')):
                last = x
                return(first, last)

    def execute_command(self, first, last, string, data, time):

        placeholder = string[first:last]
        counter = 0
        v1 = ''
        v2 = ''
        v3 = ''
        v4 = ''
        v5 = ''
        v6 = ''
        r1 = ''
        r2 = ''
        for x in placeholder:
            if x == '\\':
                counter += 1
            elif counter == 2:
                v1 += x
            elif counter == 3:
                v2 += x
            elif counter == 4:
                v3 += x
            elif counter == 5:
                v4 += x
            elif counter == 6:
                v5 += x
            elif counter == 7:
                v6 += x

        if 'time' in v1.lower():
            r1, r2 = self.time_statistics_command(
                first, last, string, data, v2, v3, time, v4)
            return r1, r2

        elif 'stat' in v1.lower():
            r1, r2 = self.statistics_command(
                first, last, string, data, v2, v3, v4, v5, v6)

        else:
            r1, r2 = self.replace_with_variable(
                first, last, string, data, v1, v2, v3)
            r1 = self.fill_no_with_red(r1,string,first,last)
        return r1, r2
    
    def fill_no_with_red(self, r1, string, first, last):
        if 'no' in r1 or 'No' in r1:
            front = first
            while front >= 0:
                print(string[front:front+5])
                if string[front:front+5] == '<span':
                    return string[:front+5] + ' style="color: #ff0000;"' + string[front] + r1 + string[last:]
                if string[front:front+5] == '/span' or front == 0:
                    return string[:front] + '<span style="color: #ff0000;">'+ r1 + '</span>' + string[last:]
                front -= 1
        return r1

    def statistics_command(self, first, last, string, data, option,
                           sheet_number, column, row_range, title=''):

        num1 = 0
        num2 = 0
        options = []
        numbers = []
        _return_string = ''

        if row_range != '':

            if row_range[0] == ':' or row_range[len(row_range) - 1] == ':':
                if row_range[0] == ':':
                    row_range = str(
                        len(data[int(sheet_number) - 1])) + row_range
                if row_range[len(row_range) - 1] == ':':
                    row_range = row_range + \
                        str(len(data[int(sheet_number) - 1]))

            for x in range(0, len(row_range)):
                if row_range[x] == ':':
                    temp = int(row_range[:x])
                    if temp <= int(row_range[x + 1:]):
                        num1 = temp
                        num2 = int(row_range[x + 1:])
                    else:
                        num1 = int(row_range[x + 1:])
                        num2 = temp
                    continue
        else:
            num1 = 1
            num2 = len(data[int(sheet_number) - 1])

        if num1 < 0 or num2 < 0:
            if num1 < 0:
                num1 = len(data[int(sheet_number) - 1]) + num1 - 1
            if num2 < 0:
                num2 = len(data[int(sheet_number) - 1]) + num2 - 1

        if column.isdigit() == False:
            column = self.column_id_to_int(column)

        options, numbers = self.statistics_internal(
            data, sheet_number, column, num1, num2)

        if 'num' in option.lower():
            for x in range(0, len(options)):
                _return_string += (options[x] + ':' + str(numbers[x]) + ' ')
            return string[:first] + _return_string + string[last:], None
        elif 'per' in option.lower():
            total = sum(numbers)
            for x in range(0, len(options)):
                _return_string += (options[x] + ':' +
                                   str(int(numbers[x] / total * 100)) + '% ')
            return string[:first] + _return_string + string[last:], None
        elif 'pie' in option.lower():
            _return_string, attach = self.create_pie(title, options, numbers)
            return string[:first] + _return_string + string[last:], attach
        elif 'bar' in option.lower():
            _return_string, attach = self.create_bar(title, options, numbers)
            return string[:first] + _return_string + string[last:], attach
        elif 'line' in option.lower():
            _return_string, attach = self.create_line(title, options, numbers)
            return string[:first] + _return_string + string[last:], attach
        else:
            print('statistics error: no choice')
            return string[:first] + string[last:], None

    def time_statistics_command(
            self,
            first,
            last,
            string,
            data,
            option,
            sheet_number,
            time,
            title=''):
        options = []
        numbers = []
        start = -1
        stop = 0
        _return_string = ''
        if time is None:
            print("Time Stats Error")
            return string[:first] + string[last:]
        else:
            for x in range(1, len(data[int(sheet_number) - 1])):
                change = parser.parse(data[int(sheet_number) - 1][x][0]) - time
                if timedelta.total_seconds(change) > 0:
                    if start == -1:
                        start = x
                    else:
                        stop = x
            options, numbers = self.statistics_internal(
                self, data, sheet_number, 0, start, stop)
        if 'num' in option.lower():
            for x in range(0, len(options)):
                _return_string += (options[x] + ':' + str(numbers[x]) + ' ')
            return string[:first] + _return_string + string[last:], None
        elif 'per' in option.lower():
            total = sum(numbers)
            for x in range(0, len(options)):
                _return_string += (options[x] + ':' +
                                   str(int(numbers[x] / total * 100)) + '% ')
            return string[:first] + _return_string + string[last:], None
        elif 'pie' in option.lower():
            _return_string, attach = self.create_pie(title, options, numbers)
            return string[:first] + _return_string + string[last:], attach
        elif 'bar' in option.lower():
            _return_string, attach = self.create_bar(title, options, numbers)
            return string[:first] + _return_string + string[last:], attach
        elif 'line' in option.lower():
            _return_string, attach = self.create_line(title, options, numbers)
            return string[:first] + _return_string + string[last:], attach
        else:
            print('time statistics error: no choice')
            return string[:first] + string[last:], None

    def statistics_internal(self, data, sheet_number, column, first, last):
        options = []
        numbers = []
        for x in range(first - 1, last):
            if ',' in data[int(sheet_number) - 1][x][int(column) - 1]:
                values = data[int(sheet_number) - 1][x][int(column) - 1]
                var_list = []
                place_holder = 0
                for x in range(0, len(values)):
                    if values[x] == ',':
                        var_list += [values[place_holder:x]]
                        place_holder = x + 1
                    if values[x:x + 2] == ', ':
                        place_holder += 1
                    if x == len(values) - 1:
                        var_list += [values[place_holder:x + 1]]
            else:
                var_list = [data[int(sheet_number) - 1][x][int(column) - 1]]
            for x in var_list:
                if x in options:
                    for y in range(0, len(options)):
                        if options[y] == x:
                            numbers[y] += 1
                else:
                    options += [x]
                    numbers += [1]
        return options, numbers
        
    def replace_with_variable(
            self,
            first,
            last,
            string,
            data,
            sheet_number,
            column,
            row):
        return_string = ''

        if column.isdigit() == False:
            column = self.column_id_to_int(column)
        if row == '':
            if sheet_number == '' or column == '':
                return string[:first] + string[last:]
            else:
                return_string = str(data[int(sheet_number) - 1][len(
                    data[int(sheet_number) - 1]) - 1][int(column) - 1])
        elif '-' in row:
            return_string = str(data[int(sheet_number) -
                                     1][len(data[int(sheet_number) -
                                                 1]) -
                                        int(row[1:]) -
                                        1][int(column) -
                                           1])
        else:
            return_string = str(data[int(sheet_number) - 1]
                                [int(row) - 1][int(column) - 1])
        return_string = self.fill_no_with_red(return_string,string,first,last)
        if self.create_images(return_string)[0] is None:
            return string[:first] + return_string + string[last:], None
        else:
            img_string, img_attach = self.create_images(return_string)
            return string[:first] + img_string + string[last:], img_attach

    def create_images(self, value):
        if 'https://drive.google.com/open?id=' in value:
            if ',' in value:
                img_list = []
                img_filelist = []
                place_holder = 0
                return_string = ''
                for x in range(0, len(value)):
                    if value[x] == ',':
                        img_list += [value[place_holder:x]]
                        place_holder = x + 1
                    if value[x] == ' ':
                        place_holder += 1
                    if x == len(value) - 1:
                        img_list += [value[place_holder:x + 1]]
                for x in img_list:
                    img_filename = GDrive.download(x)
                    img_filelist += [img_filename]
                    return_string += '<img src="cid:' + img_filename + '"/>'
                return return_string, img_filelist
            else:
                img_filename = GDrive.download(value)
                return '<img src="cid:' + img_filename + '"/>', [img_filename]
        else:
            return None, None

    def create_pie(self, title, options, numbers):
        def make_autopct(values):
            def my_autopct(pct):
                total = sum(values)
                val = int(round(pct * total / 100.0))
                return '{p:.2f}% ({v:d})'.format(p=pct, v=val)
            return my_autopct
        plt.pie(numbers, labels=options, autopct=make_autopct(numbers))
        plt.title(title)
        plt.axis('equal')
        home_dir = os.path.abspath(os.path.join(__file__, os.pardir))
        graph_dir = os.path.join(home_dir, '.images')
        counter = 1
        while True:
            graphfile = os.path.join(
                graph_dir,
                'pie' +
                self.int_to_column_id(counter).upper() +
                '.png')
            if not os.path.exists(graphfile):
                print(graphfile)
                plt.savefig('pie' + self.int_to_column_id(counter).upper())
                break
            counter += 1
        plt.clf()
        return '<img src="cid:' + 'pie' + self.int_to_column_id(counter).upper(
        ) + '.png' + '"/>', ['pie' + self.int_to_column_id(counter).upper() + '.png']

    def create_bar(self, title, options, numbers):
        y_pos = np.arange(len(options))
        plt.bar(y_pos, numbers, align='center')
        plt.xticks(y_pos, options)
        plt.title(title)
        home_dir = os.path.abspath(os.path.join(__file__, os.pardir))
        graph_dir = os.path.join(home_dir, '.images')
        counter = 1
        while True:
            graphfile = os.path.join(
                graph_dir,
                'bar' +
                self.int_to_column_id(counter).upper() +
                '.png')
            if not os.path.exists(graphfile):
                print(graphfile)
                plt.savefig('bar' + self.int_to_column_id(counter).upper())
                break
            counter += 1
        plt.clf()
        return '<img src="cid:' + 'bar' + self.int_to_column_id(counter).upper(
        ) + '.png' + '"/>', ['bar' + self.int_to_column_id(counter).upper() + '.png']

    def create_line(self, title, options, numbers):
        list_x = []
        for x in range(0, len(options)):
            list_x += [x]
        x = np.array(list_x)
        y = np.array(numbers)
        plt.xticks(x, options)
        plt.plot(x, y)
        plt.title(title)
        home_dir = os.path.abspath(os.path.join(__file__, os.pardir))
        graph_dir = os.path.join(home_dir, '.images')
        counter = 1
        while True:
            graphfile = os.path.join(
                graph_dir,
                'line' +
                self.int_to_column_id(counter).upper() +
                '.png')
            if not os.path.exists(graphfile):
                print(graphfile)
                plt.savefig(
                    'line' + self.int_to_column_id(counter).upper())
                break
            counter += 1
        plt.clf()
        return '<img src="cid:' + 'line' + self.int_to_column_id(counter).upper(
        ) + '.png' + '"/>', ['line' + self.int_to_column_id(counter).upper() + '.png']

    def check_dirs(self):
        home_dir = os.path.abspath(os.path.join(__file__, os.pardir))
        image_dir = os.path.join(home_dir, '.images')
        if not os.path.exists(image_dir):
            os.makedirs(image_dir)
        os.chdir(image_dir)

    def validate_emails(self, addresses, group):
        first = 0
        last = 0
        for x in range(0, len(addresses)):
            if addresses[x] == ',':
                last = x
                if '.' not in addresses[first:last] or '@' not in addresses[first:last]:
                    self.create_error("Invalid email address in " + group)
                    return False
                first = x
        if '.' not in addresses[first:] or '@' not in addresses[first:]:
            self.create_error("Invalid email address in " + group)
            return False
        return True

    def create_error(self, message):
        tk.messagebox.showinfo("Error", message)
        self.running = False
        del(self.running_groups)
        gc.collect()

    def open_file(self):
        self.openfilename = filedialog.askopenfilename(
            initialdir='/',
            title='Select file',
            filetypes=(
                ('text files',
                 '*.txt'),
                ('all files',
                 '*.*')))
        self.clear_all()
        openfile = open(self.openfilename, 'r')
        self.savefilename = self.openfilename
        line = openfile.readline()
        while line:
            line = openfile.readline()
            if line == 'Sheet\n':
                print('Loading Sheet...')
                self.create_sheet()
                line = openfile.readline()
                self.GSheetURLEntries[len(
                    self.GSheetURLEntries) - 1].insert(0, line[:len(line) - 1])
                line = openfile.readline()
                self.GSheetTitleEntries[len(
                    self.GSheetTitleEntries) - 1].insert(0, line[:len(line) - 1])
                line = openfile.readline()

            if line == 'Group\n':
                print('Loading Group...')
                self.create_group()
                line = openfile.readline()
                self.email_entries[len(self.email_entries) -
                                   1].insert(0, line[:len(line) - 1])
                line = openfile.readline()
                self.subject_entries[len(
                    self.subject_entries) - 1].insert(0, line[:len(line) - 1])
                line = openfile.readline()
                self.send_method[len(self.send_method) -
                                 1].set(line[:len(line) - 1])
                line = openfile.readline()
                self.send_entries[len(self.send_entries) -
                                  1].insert(0, line[:len(line) - 1])
                line = openfile.readline()
                self.send_entries2[len(self.send_entries2) -
                                   1].insert(0, line[:len(line) - 1])
                while True:
                    line = openfile.readline()
                    if line == 'Sheet\n' or line == 'Group\n' or line is None or line == '':
                        break
                    self.html_entries[len(self.html_entries) -
                                      1].insert(tk.END, line[:len(line) - 1])
        openfile.close()
        print('File Loaded')

    def clear_all(self):
        print(len(self.sheets))
        if len(self.sheets) != 0 and len(self.groups) != 0:
            for x in range(len(self.sheets) - 1):
                self.sheet_book.forget(x)
                del self.sheets[x]
                del self.GSheetURLEntries[x]
                del self.GSheetTitleEntries[x]
            for x in range(len(self.groups) - 1):
                self.group_book.forget(x)
                del self.groups[x]
                del self.email_entries[x]
                del self.subject_entries[x]
                del self.send_method[x]
                del self.send_entries[x]
                del self.send_entries2[x]
                del self.html_entries[x]

    def save_file(self):
        if self.savefilename:
            self.save_internal(self.savefilename)
        else:
            self.savefilename = filedialog.asksaveasfilename(
                initialdir='/',
                title='Select file',
                filetypes=(
                    ('text files',
                     '*.txt'),
                    ('all files',
                     '*.*')))
        if self.savefilename is not None or self.savefilename != '':
            self.save_internal(self.savefilename)
        else:
            print('\tWarn: Did not save: No File name')

    def save_as_file(self):
        self.savefilename = filedialog.asksaveasfilename(
            initialdir='/',
            title='Select file',
            filetypes=(
                ('text files',
                 '*.txt'),
                ('all files',
                 '*.*')))
        if self.savefilename is not None and self.savefilename != '':
            self.save_internal(self.savefilename)
        else:
            print('\tWarn: Did not save: No File name')

    def save_internal(self, filename):
        if '.txt' in filename:
            savefile = open(filename, 'w')
        else:
            savefile = open(filename + '.txt', 'w')

        savefile.write('SaveFile for an Email Formatter\n')
        savefile.write('\tby William Crawford\n\n')

        for x in range(0, len(self.sheets) - 1):
            savefile.write('Sheet\n')
            savefile.write(self.GSheetURLEntries[x].get() + '\n')
            savefile.write(self.GSheetTitleEntries[x].get() + '\n')
            savefile.write('\n')
        for x in range(0, len(self.groups) - 1):
            savefile.write('Group\n')
            savefile.write(self.email_entries[x].get() + '\n')
            savefile.write(self.subject_entries[x].get() + '\n')
            savefile.write(self.send_method[x].get() + '\n')
            savefile.write(self.send_entries[x].get() + '\n')
            savefile.write(self.send_entries2[x].get() + '\n')
            savefile.write(self.html_entries[x].get(1.0, tk.END) + '\n')
        savefile.close()
        print("\tSave Complete")

    def set_credentials(self):
        self.root.openfilename = filedialog.askopenfilename(
            initialdir='/',
            title='Select file',
            filetypes=(
                ('json files',
                 '*.json'),
                ('all files',
                 '*.*')))

    def about(self):
        self.top = tk.Toplevel()
        self.top.title('About')

        self.msg = tk.Label(self.top, text='Programmed by:')
        self.msg.pack()
        self.msg = tk.Label(self.top, text='William Crawford')
        self.msg.pack()

        self.button = tk.Button(
            self.top,
            text='Close',
            command=self.top.destroy)
        self.button.pack()

    def checkgroups(self):
        new_data = self.get_data(self.GSheetURLEntries, self.GSheetTitleEntries)
        for x in self.running_groups:
            x.check(self, new_data)
        if self.running:
            self.root.after(10000, self.checkgroups)

    def run(self):

        self.load_data()
        self.running_groups = []
        self.running = True
        self.root.withdraw()
        self.run_window = self.RunWindow(self)

        for i in range(0, len(self.send_method)):
            if self.validate_emails(self.email_entries[i].get()):
                if self.send_method[i].get() == 'TimeInterval':
                    if self.send_entries[i].get(
                    ) is None or self.send_entries[i].get() == "":
                        self.create_error(
                            "Time is empty for TimeInterval in Group" + str(i + 1))
                        break
                    if self.send_entries2[i].get(
                    ) is None or self.send_entries2[i].get() == "":
                        self.create_error(
                            "Interval is empty for TimeInterval in Group" + str(i + 1))
                        break
                    self.running_groups += [self.TimeInterval(self, i,
                                                              self.send_entries[i].get(
                                                              ),
                                                              self.send_entries2[i].get(
                                                              ),
                                                              self.email_entries[i].get(
                                                              ),
                                                              self.subject_entries[i].get(
                                                              ),
                                                              str(self.html_entries[i].get(1.0,
                                                                                           tk.END)),
                                                              self.sheet_data
                                                              )]
                elif self.send_method[i].get() == 'Keyword':
                    if self.send_entries[i].get(
                    ) is None or self.send_entries[i].get() == "":
                        self.create_error(
                            "SheetIdentifier is empty for Keyword in Group" + str(i + 1))
                        break
                    if self.send_entries2[i].get(
                    ) is None or self.send_entries2[i].get() == "":
                        self.create_error(
                            "Keyword is empty for Keyword in Group" + str(i + 1))
                        break
                    self.running_groups += [self.Keyword(self, i,
                                                         self.send_entries[i].get(
                                                         ),
                                                         self.send_entries2[i].get(
                                                         ),
                                                         self.email_entries[i].get(
                                                         ),
                                                         self.subject_entries[i].get(
                                                         ),
                                                         str(self.html_entries[i].get(1.0,
                                                                                      tk.END)),
                                                         self.sheet_data
                                                         )]
                elif self.send_method[i].get() == 'Response#':
                    if self.send_entries[i].get(
                    ) is None or self.send_entries[i].get() == "":
                        self.create_error(
                            "SheetNumber is empty for Response# in Group" + str(i + 1))
                        break
                    if self.send_entries2[i].get(
                    ) is None or self.send_entries2[i].get() == "":
                        self.create_error(
                            "Interval is empty for Response# in Group" + str(i + 1))
                        break
                    self.running_groups += [self.Response(self, i,
                                                          self.send_entries[i].get(
                                                          ),
                                                          self.send_entries2[i].get(
                                                          ),
                                                          self.email_entries[i].get(
                                                          ),
                                                          self.subject_entries[i].get(
                                                          ),
                                                          str(self.html_entries[i].get(1.0,
                                                                                       tk.END)),
                                                          self.sheet_data
                                                          )]
                else:
                    self.create_error("Group" + str(i + 1) +
                                      " does not have a send method")
                    break
            else:
                break

        self.checkgroups()

    def force_email(self):

        if messagebox.askyesno(
            'Confirmation',
                'Are you sure that you want to force send emails?'):

            self.load_data()

            for x in range(0, len(self.groups) - 1):
                if self.validate_emails(
                        self.email_entries[x].get(), "Group" + str(x + 1)):
                    html, attach = self.email_command_execution(str(
                        self.html_entries[x].get(1.0, tk.END)),
                        self.sheet_data)
                    GMail.SendMessage('me', self.email_entries[x].get(),
                                      self.subject_entries[x].get(),
                                      html, "", attach)

    def close(self):
        if self.savefilename:
            self.save_file()
        elif messagebox.askyesno('Quit',
                                 'You have not saved your work. Do you want to save it?'):
            self.save_as_file()
        self.root.destroy()


if __name__ == '__main__':
    app = Form2Email()
