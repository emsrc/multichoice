#!/usr/bin/python

"""
Parse multiple choice questions in Bokmal or Nynork and corresponding answer
key to create a comma-separated file, which can be imported into Excel
"""

# Requires Python2

# EM - 2015

import re
import csv


QUEST_START = re.compile(r"\d+\.\W+(.+)")

ANSWER_START = re.compile(r"\W+[ABCD]\)\W+(.+)")

#CONTINUATION = re.compile(r"\W+(\w.+)")
CONTINUATION = re.compile(r"\W*(\w.+)")

SECTION = re.compile(r"(\w.*)")

EMPTY = re.compile("(\W*)$")


def parse_exam(exam_fname):
    # ignore encoding
    f = open(exam_fname, "rb")
    # split lines on \n only, not on \r inserted by MS Word
    lines = f.read().split("\n")
    item = []
    in_question = False
    in_answer = False

    for i, line in enumerate(lines):
        # SECTION also matches start of question,
        # so QUEST_START must come first
        match = ( QUEST_START.match(line) or
                  ANSWER_START.match(line) or
                  CONTINUATION.match(line) or
                  SECTION.match(line) or
                  EMPTY.match(line) )

        if not match:
            raise Exception("Error parsing line {}:\n{!r}".format(i+1, line))

        text = match.group(1).strip()

        if match.re is QUEST_START:
            print('QUEST_START', line)
            item = [text]
            in_question = True
            in_answer = False
        elif match.re is ANSWER_START:
            print('ANSWER START', line)
            item.append(text)
            in_answer = True
            in_question = False
        elif match.re is CONTINUATION and (in_question or in_answer):
            print('CONTINUATION', line)
            item[-1] += " " + text
        elif match.re is SECTION:
            print('SECTION', line)
            in_answer = False
            in_question = False
            yield text
        elif match.re is EMPTY and len(item) == 5:
            print('EMPTY', line)
            in_answer = False
            in_question = False
            yield item
            item = []


def mcparse(exam1_fname, exam2_fname, out_fname):
    exam1 = parse_exam(exam1_fname)
    exam2 = parse_exam(exam2_fname)

    # encoding is the same as in the input files
    with open(out_fname, 'wb') as outf:
        writer = csv.writer(outf, lineterminator="\n")

        for item1, item2 in zip(exam1, exam2):
            #print(item1,item2)
            if isinstance(item1, basestring):
                writer.writerow([item1] + 9 * [""])
            else:
                writer.writerow(item1 + item2)


#mcparse("MC1011_HaraldHamre_bokmal2.txt",
#        "MC1011_HaraldHamre_bokmal2.txt",
#        "output.csv")


# Save ms word doc as txt file with UTF-8 encoding and LF only line breaks.
# Import csv in excel with utf-8 encoding (under "File origin"!) and with comma as delimiter.

# mcparse("MC1011b_del1_HH.txt",
#         "MC1011b_del1_HH.txt",
#         "MC1011b_del1_HH.csv")
#
#
# mcparse("MC1011b_del2_HH.txt",
#         "MC1011b_del2_HH.txt",
#         "MC1011b_del2_HH.csv")


mcparse("MC1011b_del1_HH.txt",
        "MC1011n_del1_HH.txt",
        "MC1011_del1_HH.csv")

mcparse("MC1011b_del2_HH.txt",
        "MC1011n_del2_HH.txt",
        "MC1011_del2_HH.csv")


