import os
import re
import tkinter
from tkinter import filedialog, messagebox
import string
from tkinter.filedialog import asksaveasfile

import nltk
from nltk.corpus import stopwords
from string import punctuation
from nltk.probability import FreqDist
import xlrd
import pandas as pd


def show_results(text):
    tkinter.messagebox.showinfo("Results", text)

def save():
    files = [('All Files', '*.*'),
             ('Python Files', '*.py'),
             ('Text Document', '*.txt')]
    app.filename = filedialog.asksaveasfilename(initialdir='/', title='Select file', filetypes = files, defaultextension = files)
    print(app.filename)

def show_error(text):
    tkinter.messagebox.showinfo("Error",text)


def genlist(file):
    decks={}
    results=""
    if os.path.abspath(file).endswith('.txt'):
        decks[(os.path.abspath(file))]=0
    else:
        show_error("the file %s is not a .txt file and will be ignored." %(file))

    f = open(file, 'r', encoding='UTF-8')
    text = f.read()
    text = text.lower()

    # strip punctuation
    exclude = set(string.punctuation)
    exclude.remove(".")
    for punctuation in exclude:
        text = text.replace(punctuation, "")
    text = text.replace(".", "")
    text = text.replace("—", " ")
    ntext = ''.join(i for i in text if not i.isdigit())
    s = ntext.split()

    # remove stopwords
    sw = set(stopwords.words("english"))
    fl = [w for w in s if not w in sw]


    # remove duplicate
    def unique_list(l):
        ulist = []
        [ulist.append(x) for x in l if x not in ulist]
        return ulist

    fl = unique_list(fl)

    fd = FreqDist(fl)

    f.close()

    words = set(nltk.corpus.words.words())
    nfl = [y for y in fl if y in words or not y.isalpha()]

    # remove known
    def remove_known(known):
        wb = xlrd.open_workbook(known)
        sheet = wb.sheet_by_index(0)
        sheet.cell_value(0, 0)
        wordlist = []
        for i in range(sheet.nrows):
            wordlist.append(sheet.cell_value(i, 0))
        nnfl = [x for x in nfl if not x in wordlist]
        return nnfl

    nfl = remove_known("CET4.xlsx")
    #print(len(nfl))
    nfl = remove_known("highschool.xlsx")
    results += ("%s\t%s\n" % (os.path.basename(file),len(nfl)))
    show_results(results)

    test = pd.DataFrame(data=nfl)
    test.to_csv('G:\hobby\python\excel\list_try_maggie.csv')


app = tkinter.Tk()
app.geometry('275x175')
app.title("wordlist")


def clicked():
    t = tkinter.filedialog.askopenfilename()
    wordfile = genlist(t)
    savefile = tkinter.Button(app, text='save', command=save)
    savefile.pack(fill='x')
    return wordfile


header=tkinter.Label(app,text="Welcome to worklist",fg="blue",font=("Arial Bold",16))
header.pack(side="top",ipady=10)

text=tkinter.Label(app,text="Select .txt file")
text.pack()
'''
show_summary=tkinter.BooleanVar()
show_summary.set(True)
summary=tkinter.Checkbutton(app,text="show summary",var=show_summary)
summary.pack(ipady=10)
'''
open_files=tkinter.Button(app,text="Choose file...",command=clicked)
open_files.pack(fill="x")

app.mainloop()

