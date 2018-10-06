#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Andrew Li
EOL 2017
Text Analysis Tool
"""

import docx2txt
import matplotlib; matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
import pandas as pd
from textstat.textstat import textstat
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import tkinter as tk


class Window:

    def __init__(self, master):
        self.filename = ''
        self.ease = ''
        self.grade = ''
        self.files = []
        self.easeScores = []
        self.gradeScores = []

        # Sizing window
        master.minsize(width=400, height=400)
        master.resizable(width=False, height=False)

        # Centering window
        ws = master.winfo_screenwidth()  # width of the screen
        hs = master.winfo_screenheight()  # height of the screen
        x = (ws / 2) - (400 / 2)
        y = (hs / 2) - (400 / 2)
        master.geometry('%dx%d+%d+%d' % (400, 400, x, y))
        master.title('EOL 2017')

        # Creating title label
        self.title = Label(master, text='Text Analysis', fg='blue', font='Helvetica 20 bold', padx=10)
        self.title.pack()

        # Creating info label
        self.info = Label(master,
                          justify=LEFT,
                          text=
                          'Flesch Reading Ease grading scale:\n'
                          '71-100: Very Confusing\n'
                          '51-70: Difficult\n'
                          '41-50: Fairly Difficult\n'
                          '31-40: Standard\n'
                          '21-30: Fairly Easy\n'
                          '11-20: Easy\n'
                          '0-10: Very Easy\n\n'
                          'Flesch-Kincaid Grade grading scale:\n'
                          'Returns grade level approximation.\n'
                          'Ex. 9.3 means that a ninth grader could read the text.\n',
                          fg='red')
        self.info.pack()

        # Creating filename label
        self.file = Label(master, text='Filename: %s' % self.filename, font='Helvetica 14 bold')
        self.file.pack()

        # Creating reading ease score label
        self.ease = Label(master, text='Flesch Reading Ease scale score: %s' % self.ease, font='Helvetica 14 bold')
        self.ease.pack()

        # Creating grading scale score label
        self.grade = Label(master, text='Flesch-Kincaid Grade scale score: %s' % self.grade, font='Helvetica 14 bold')
        self.grade.pack()

        # Creating file selection button
        self.button = tk.Button(master, text='Choose File', width=25, command=self.open_file)
        self.button.pack()

        # Creating plot button
        self.exit = tk.Button(master, text='Plot Data', width=25, command=self.plot)
        self.exit.pack()

        # Creating about button
        self.exit = tk.Button(master, text='About Grading Scale', width=25, command=about)
        self.exit.pack()

    def open_file(self):
        self.filename = filedialog.askopenfilename()
        if self.filename.endswith('.txt') or self.filename.endswith('.docx'):
            self.read_file()
        elif self.filename is not '':
            messagebox.showerror(title='ERROR', message='Wrong file format selected. Please try again.')

    def read_file(self):
        text = ''
        with open(self.filename, 'r', encoding='UTF-8', errors='ignore') as file:
            if self.filename.endswith('.docx'):
                file = docx2txt.process(self.filename)
            self.filename = self.filename[self.filename.rfind('/') + 1:]
            for line in file:
                text += line.replace('\n', '')
            self.score_text(text)

    def score_text(self, txt):
        ease = float('%.2f' % (100 - textstat.flesch_reading_ease(txt)))
        grade0 = float('%.2f' % textstat.flesch_kincaid_grade(txt))
        grade = grade0
        if grade >= 13:
            grade = '12+ (%s)' % grade

        self.file.config(text='Filename: %s' % self.filename)
        self.ease.config(text='Flesch Reading Ease scale score: %s' % ease)
        self.grade.config(text='Flesch-Kincaid Grade scale score: %s' % grade)
        if self.filename not in self.files:
            self.files.append(self.filename[:self.filename.find('.')])
            self.easeScores.append(ease)
            self.gradeScores.append(grade0)

    def plot(self):
        try:
            df = pd.DataFrame({'Flesch-Kincaid Grade level': self.gradeScores,
                               'Flesch Reading Ease': self.easeScores},
                              index=self.files)

            ax1 = df.plot(kind='bar', figsize=(10, 5))
            plt.savefig('image.png', bbox_inches='tight', dpi=100)
            ax1.set_xticklabels(ax1.xaxis.get_majorticklabels(), rotation=0)
            ax1.set_ylabel('Flesch Reading Ease', color='b')
            ax2 = ax1.twinx()
            for r in ax1.patches[len(df):]:
                r.set_transform(ax2.transData)
            ax1.set_ylim(0, 100)
            ax2.set_ylim(0, 25)

            for p in ax1.patches:
                ax1.annotate(str(p.get_height()), (p.get_x() + p.get_width() / 2, 5),
                             ha='center', va='center')

            ax2.set_ylabel('Flesch-Kincaid Grade level', color='r')
            plt.title('Text Analysis Stats')
            plt.show()
        except TypeError:
            messagebox.showerror(title='ERROR', message='No data to plot. Please try again.')


def about():
    about_window = Toplevel()

    about_window.resizable(width=False, height=False)
    ws = about_window.winfo_screenwidth()  # width of the screen
    hs = about_window.winfo_screenheight()  # height of the screen
    x = (ws / 2) - (400 / 2)
    y = (hs / 2) - (400 / 2)
    about_window.geometry('%dx%d+%d+%d' % (400, 200, x, y))

    fre = PhotoImage(file='/Users/MacBook/PycharmProjects/Fun/fre.gif')
    fkg = PhotoImage(file='/Users/MacBook/PycharmProjects/Fun/fkg.gif')

    fre_text = Label(about_window, text='Scoring formula for the Flesch Reading Ease test.')
    fre_text.pack()
    fre_image = Label(about_window, image=fre)
    fre_image.pack()

    fkg_text = Label(about_window, text='Scoring formula for the Flesch-Kincaid Grade Level test.')
    fkg_text.pack()
    fkg_image = Label(about_window, image=fkg)
    fkg_image.pack()

    about_window.title('About')
    about_window.mainloop()

if __name__ == '__main__':
    root = Tk()
    window = Window(root)
    root.mainloop()
