import sys
import os

import math
import string
import re

import pandas as pd
import csv

import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog

class Root(Tk):
    def __init__(self):
        super(Root, self).__init__()
        self.title("Spreadsheet separate sheet adding program")
        self.minsize(500, 500)

        self.labelFrame = ttk.LabelFrame(self, text = "Select Spreadsheet")
        self.labelFrame.place(relx = .5, rely = .075, anchor=CENTER)

        self.button()



    def button(self):
        self.button = ttk.Button(self.labelFrame, text = "Browse...",
            command = self.fileDialog)
        self.button.grid(column = 1, row = 1)
        self.label = ttk.Label(self.labelFrame, text = "")
        self.label.grid(column = 1, row = 2)


    def fileDialog(self):
        self.filename = filedialog.askopenfilename(initialdir =  ".",
            title = "Select A Spreadsheet",
            filetype = [("excel files","*.xls*" )])
        self.label.configure(text = self.filename)
        self.load_spreadsheet(self.filename)

    def load_spreadsheet(self, filepath):
        self.excel_file = pd.ExcelFile(filepath)
        self.all_sheets = pd.read_excel(self.excel_file, None, header=None)
        print(self.all_sheets, flush=True)
        self.create_dialogues()
        self.submit_frame()

    def create_dialogues(self):
        self.dialogue_frame = ttk.LabelFrame(self, text = "Select Cells")
        self.dialogue_frame.place(relx = .5, rely = .35, anchor=CENTER)

        instructions = ttk.Label(self.dialogue_frame,
            text = """
            For each sheet, enter the cell you would like added.
            Please enter the row as a lower case letter, e.g. d6, g8.
            If you would like to skip a sheet, simply leave that field as \"none\"
            """
        )
        instructions.grid(row=1, column=1, columnspan=4)

        self.dialogues = []
        for sheet in range(len(self.all_sheets)):
            curr = ttk.LabelFrame(
                self.dialogue_frame,
                text="Sheet " + str(sheet +1)
            )

            curr.grid(
                column = (sheet % 4) +1,
                row = math.floor(sheet / 4) +2,
                padx = 10,
                pady = 5
            )

            entry = ttk.Entry(curr, width=5)
            entry.pack()
            entry.insert(END, "none")
            self.dialogues.append(entry)

    def submit_frame(self):
        self.submission_frame = ttk.LabelFrame(self, text = "Add it up")
        self.submission_frame.place(relx = .5, rely = .625, anchor=CENTER)

        self.submit_button = ttk.Button(
            self.submission_frame,
            text="Compute",
            command = self.process
        )
        self.submit_button.grid(column = 2, row = 1, padx=20, pady=10)

    def process(self):
        item_positions = [entry.get() for entry in self.dialogues]
        alphabet = string.ascii_lowercase
        converted_positions = []
        for s in item_positions:
            if s == "none":
                converted_positions.append(s)
                continue
            decoupled = re.findall(r'(\d+|\D+)', s)
            converted = [alphabet.find(decoupled[0])] + [int(decoupled[1]) - 1]
            converted_positions.append(converted)

        index = 0
        numbers = []
        for item in self.all_sheets:
            coordinates = converted_positions[index]
            index += 1
            if coordinates == 'none':
                continue
            column, row = coordinates[0], coordinates[1]
            numbers.append(self.all_sheets[item].iat[row, column])

        RESULTS = [
            ["All numbers: "] + numbers,
            ["Average: ", sum(numbers)/len(numbers)],
            ["Total: ", sum(numbers)]
        ]

        readable = [
            RESULTS[0][0] + ", ".join(map(str,RESULTS[0][1:])),
            RESULTS[1][0] + str(RESULTS[1][1]),
            RESULTS[2][0] + str(RESULTS[2][1])
        ]

        result_list = Listbox(self.submission_frame)
        result_list.grid(row=1, column=1)
        result_list.config(width=30, height=0)
        index = 1
        for entry in readable:
            result_list.insert(index, entry)
            index += 1

        self.result(RESULTS)

    def result(self, res):
        self.save_frame = ttk.LabelFrame(self, text = "Save Result")
        self.save_frame.place(relx = .5, rely = .85, anchor=CENTER)

        self.browse_button = ttk.Button(
            self.save_frame,
            text = "Select directory",
            command = self.select_save_file
        )

        self.save_entry = ttk.Entry(
            self.save_frame,
            width = 10,
        )
        self.save_entry.insert(END, "output")

        self.save_update = ttk.Button(
            self.save_frame,
            text = "Update save location",
            command = self.update_save
        )

        self.save_file = self.filename[:self.filename.rindex(r'/')]
        self.save_path = self.save_file + "/output.csv"
        self.save_label = ttk.Label(self.save_frame, text=self.save_path)

        self.save_button = ttk.Button(
            self.save_frame,
            text = "Save CSV",
            command = lambda a1=res : self.save(a1)
        )

        self.browse_button.grid(column=1, row=1)
        self.save_entry.grid(column=2, row=1)
        self.save_update.grid(column=3, row=1)
        self.save_label.grid(column=1, row=2, columnspan=3)
        self.save_button.grid(column = 1, row = 3,
            columnspan = 3, padx=20, pady=10)

    def save(self, res):
        print(self.save_path)
        with open(self.save_path,'w+') as result_file:
            wr = csv.writer(result_file, dialect='excel')
            wr.writerows(res)

    def select_save_file(self):
        self.save_file = filedialog.askdirectory()

    def update_save(self):
        self.save_path = self.save_file + "/" + self.save_entry.get() + ".csv"
        self.save_label.configure( text = self.save_path )

root = Root()
root.mainloop()
