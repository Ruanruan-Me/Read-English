# -*- coding: utf-8 -*-
"""
Created on Sat Apr 25 16:59:20 2020

@author: Administrator
"""
import string
import nltk
from nltk.corpus import stopwords
from string import punctuation
from nltk.probability import FreqDist
import xlrd
import pandas as pd
from nltk.corpus import wordnet as wn

f = open('Email Assignment.txt', 'r', encoding='UTF-8')
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
print(len(fl))


# plural

def plural(word):
    if word[-3:] == 'ies':
        return word[:-3] + 'y'
    elif word[-2:] == 'es' and (word[-4:-2] in ['sh', 'ch'] or word[-3:-2] in 'sx'):
        return word[:-2]
    elif word[-1:] == 's':
        return word[:-1]
    else:
        return word


rfl = []
for w in fl:
    word = plural(w)
    rfl.append(word)
print(len(rfl))


# remove duplicate
def unique_list(l):
    ulist = []
    [ulist.append(x) for x in l if x not in ulist]
    return ulist


fl = unique_list(rfl)
print(len(fl))
fd = FreqDist(fl)

f.close()

# if words
words = set(nltk.corpus.words.words())
nfl = [y for y in fl if y in words]
print(len(nfl))

#adv convert
def advtoadj(wordtoinv):
    s = []
    winner = ""
    for ss in wn.synsets(wordtoinv):
        for lemmas in ss.lemmas(): # all possible lemmas.
            s.append(lemmas)

    for pers in s:
        if len(pers.pertainyms()) == 0:
            continue
        posword = pers.pertainyms()[0].name()
        if posword[0:3] == wordtoinv[0:3]:
            winner = posword
        break

    return winner

kfl=[]
for win in nfl:
    if win[-2:] == 'ly':
        win=advtoadj(win)
    else:
        win
    kfl.append(win)

#remove None
while '' in kfl:
    kfl.remove('')

# remove known
def remove_known(known):
    wb = xlrd.open_workbook(known)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)
    wordlist = []
    for i in range(sheet.nrows):
        wordlist.append(sheet.cell_value(i, 0))
    nnfl = [x for x in kfl if not x in wordlist]
    return nnfl


kfl = remove_known("CET4.xlsx")
kfl = remove_known("augCET4.xls")
print(len(kfl))
kfl = remove_known("aughighschool.xls")
print(len(kfl))
kfl = remove_known("augcet6.xls")
print(len(kfl))
kfl = remove_known("augcleaneasy.xls")
print(len(kfl))


test = pd.DataFrame(data=kfl)
test.to_csv('G:\hobby\python\excel\Maggie_word.csv')
