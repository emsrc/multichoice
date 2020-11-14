#!/usr/bin/env python

"""
Generate multiple choice questions in Bokmal and Nynorsk and corresponding
answer key from Excel spreadsheet
"""

# Requires Python2, Numpy, Pandas, xlrd 

# EM - 27-05-2019

from codecs import open
from pandas import ExcelFile
from os.path import splitext
from textwrap import wrap
from numpy.random import shuffle


def generate(xls_fname, randomize=False, select_col=0, encoding="utf-8",
             width=120, tab_names=False, quest_col=0):
    xls = ExcelFile(xls_fname)
    fname_pat = splitext(xls_fname)[0] + "_{}.txt"
    languages = "bokmal", "nynorsk"
    lang_files = [open(fname_pat.format(lang), "w", encoding) 
                  for lang in languages]
    key_file = open(fname_pat.format("key"), "w", encoding)
    quest_n = 0
    rand_index = range(4)
    choices = "ABCD"

    if quest_col == 0:
        quest_offset = 0
    else:
        quest_offset = quest_col - 1
    
    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name)
        
        if tab_names:
            for outf in lang_files:
                outf.write(u"\n{}\n\n".format(sheet_name.upper()))

        for _, row in df.iterrows():
            if not select_col or row.iat[select_col - 1] == 1:
                quest_n += 1
                if randomize: shuffle(rand_index)
                correct = choices[rand_index.index(0)]
                key_file.write(u"{}\t{}\n".format(quest_n, correct))
                    
                for i, outf in enumerate(lang_files):
                    question = row.iat[quest_offset + i * 5]
                    if not isinstance(question, basestring): continue
                    question = u"\n\t".join(wrap(question, width))
                    outf.write(u"{}.\t{}\n".format(quest_n, question))
                    answers = row.iloc[quest_offset + i * 5 + 1: quest_offset + (i + 1) * 5]
                    shuffled_answers = answers[rand_index]
                    for choice, answer in zip(choices, shuffled_answers):
                        answer = u"\n\t\t".join(wrap(unicode(answer), width))
                        outf.write(u"\t{})\t{}\n".format(choice, answer))
                    if quest_offset:
                        ids = row.iloc[0:quest_offset]
                        ids = u", ".join(map(unicode, ids.values))
                        outf.write(u"\t<<< {} >>>\n".format(ids))
                    outf.write(u"\n")
                                   


if __name__ == "__main__":
    from argparse import ArgumentParser

    parser = ArgumentParser(description=__doc__)    
    
    parser.add_argument(
        "xls_fname",
        metavar="FILENAME",
        help="Excel spreadsheet file")
    
    parser.add_argument(
        "-t", "--tab_names",
        default=False,
        action="store_true",
        help="print tab names")
    
    parser.add_argument(
        "-r", "--randomize",
        default=False,
        action="store_true",
        help="randomize order of answers")
          
    parser.add_argument(
        "-s", "--select_col",
        type=int,
        metavar="NUMBER",
        default=0,
        help="number of column to select answers")
    
    parser.add_argument(
        "-e", "--encoding",
        choices=["windows-1252", "utf-8"],
        default="windows-1252",
        help="character encoding of output (defaults to 'windows-1252')")    
    
    parser.add_argument(
        "-w", "--width",
        type=int,
        metavar="NUMBER",
        default=120,
        help="max number of characters per line (defaults to 120)")

    parser.add_argument(
        "-q", "--quest_col",
        type=int,
        metavar="NUMBER",
        default=0,
        help="number of column containing first question")      
    
    args = parser.parse_args()
    generate(args.xls_fname, args.randomize, args.select_col, args.encoding,
             args.width, args.tab_names, args.quest_col)

    
    
