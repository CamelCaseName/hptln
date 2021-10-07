# This is a shitty script.
import csv
import os
import subprocess
import sys
import warnings

from pandas.core.series import Series

from requests_html import HTMLSession

# fix imports in case they dont exist
try:
    import pandas as pd
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas"])
finally:
    import pandas as pd
try:
    import regex as re
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "regex"])
finally:
    import regex as re
try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])

args = sys.argv  # 0 is the file itself, 1 the deciding factor

if os.path.exists(args[1]):
    print(f"Converting files in {args[1]}")
    path = args[1]
    files = []
    tnpath = path + "\\converted"

    for root, directories, dirfiles in os.walk(path):
        for name in dirfiles:
            files.append(os.path.join(root, name)[len(path) :])
        for dirs in directories:
            if not os.path.isdir(tnpath + "\\" + dirs):
                os.makedirs(tnpath + "\\" + dirs)

    if not os.path.isdir(tnpath):
        os.makedirs(tnpath)

    file_counter = 0

    for person in files:
        f_name, f_ext = os.path.splitext(path + person)
        if f_ext == ".xlsx":
            file_counter += 1
            print(".", end="")
            hp = pd.read_excel(path + person)
            hp.drop(hp.columns[1], axis=1, inplace=True)
            while hp.columns.size > 2:
                hp.drop(hp.columns[2], axis=1, inplace=True)
            for sentence in hp.loc[0]:
                try:
                    sentence.rstrip(sentence[-2])
                except AttributeError:
                    print(sentence)
            f_name, f_ext = os.path.splitext(person)
            f_name = os.path.join(tnpath + f_name)
            hp.to_csv(
                f_name + ".txt",
                header=["[General]", ""],
                index=False,
                quoting=csv.QUOTE_NONE,
                escapechar="|",
                doublequote=False,
                sep="|",
                mode="w",
            )
            fin = open(f_name + ".txt", "r", encoding="utf-8", errors="ignore")
            adj_lines = []
            fin.read
            for lines in fin.readlines():
                lines = re.sub(r"\|{1}\n", "\n", lines)
                lines = re.sub(r"\|* *\|", "|", lines)
                lines = re.sub(r"\|\|", "|", lines)
                adj_lines.append(lines)
            fin.close()
            fout = open(f_name + ".txt", "w", encoding="utf-8")
            fout.write("\ufeff")
            fout.writelines(adj_lines)
    print(f"\nConverted {file_counter} files!")

elif len(args) > 2:
    path = args[2]
    lang = args[1]
    languages = pd.Series(
        [
            "cs",
            "da",
            "de",
            "nl",
            "fi",
            "fr",
            "hu",
            "it",
            "ja",
            "ko",
            "pl",
            "pt",
            "ptbr",
            "ru",
            "es",
            "esmx",
            "tr",
        ]
    )
    if languages.isin([lang]).any():
        session = HTMLSession()
        # session.cookies.
        response = session.get(
            "https://crowdin.com/translate/house-party/158/en-de?filter=basic&value=0"
        )
        if response.status_code == 200:
            print("Connected to crowdin.com")
            response.html.render()
            print(response.html.html)
            button = response.html.find("editor_download_button")
            print(button)


else:
    warnings.warn("Please use the correct syntax", SyntaxWarning)
    print(
        "Either specify a first path to the downloaded .xlsx files, or provide the language you want with a target directory as the second option."
    )
    print(
        "Example: 'py.exe hptlndir.py de C:/translations' to download the german files, convert them and place them in the given folder."
    )
    print(
        "Or: 'py.exe hptlndir.py C:/translations' to convert the files at the location."
    )
