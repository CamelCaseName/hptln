# This is a shitty Python script.
import os
import sys
import subprocess
import csv

try:
    import pandas as pd
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", 'pandas'])
finally:
    import pandas as pd
try:
    import regex as re
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", 'regex'])
finally:
    import regex as re
try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", 'openpyxl'])
path = sys.argv[1]
files = os.listdir(path)
for person in files:
    f_name, f_ext = os.path.splitext(person)
    if f_ext == ".xlsx":
        hp = pd.read_excel(path + "\\" + person)
        hp.drop(hp.columns[1], axis=1, inplace=True)
        while hp.columns.size > 2:
            hp.drop(hp.columns[2], axis=1, inplace=True)
        for sentence in hp.loc[0]:
            try:
                sentence.rstrip(sentence[-2])
            except AttributeError:
                print(sentence)
        hp.to_csv(path + "\\" + f_name + ".txt", header=["[General]",""], index=False, 
                  quoting=csv.QUOTE_NONE, escapechar="|", doublequote=False, sep="|", mode='w')
        fin = open(path + "\\" + f_name + ".txt", "r", encoding="utf-8", errors="ignore")
        adj_lines = []
        fin.read
        for lines in fin.readlines():
            lines = re.sub(r"\|{1}\n", "\n", lines)
            lines = re.sub(r"\|* *\|", "|", lines)
            lines = re.sub(r"\|\|", "|", lines)
            adj_lines.append(lines)
        fin.close()
        fout = open(path + "\\" + f_name + ".txt", "w", encoding="utf-8")
        fout.write('\ufeff')
        fout.writelines(adj_lines)
