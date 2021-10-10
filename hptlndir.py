# This is a shitty script.
import csv
import os
import subprocess
import sys
import warnings

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

try:
    from requests_html import HTMLSession
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "requests_html"])
finally:
    from requests_html import HTMLSession

    # import requests


args = sys.argv  # 0 is the file itself, 1 the deciding factor

if os.path.exists(args[1]):
    print(f"Converting files in {args[1]}")
    path = args[1]
    files = []
    tnpath = path + "\\converted"

    for root, directories, dirfiles in os.walk(path):
        for name in dirfiles:
            files.append(os.path.join(root, name)[len(path) :])
        for rot in root:
            if not os.path.isdir(tnpath + "\\" + root[len(path) :]):
                os.makedirs(tnpath + "\\" + root[len(path) :])

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
            fin.read()
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
        print("Connecting to crowdin.com...")
        session = HTMLSession()
        REQUESTURL = (
            "https://crowdin.com/translate/house-party/158/en-de?filter=basic&value=0"
        )
        LOGINURL = "https://accounts.crowdin.com/login"
        proxies = {"http": "http://127.0.0.1:8090", "https": "https://127.0.0.1:8080"}

        login_page = session.get(
            LOGINURL,
            # proxies=proxies,
            # # verify=False
        )

        if login_page.status_code == 200:
            print("Connected to crowdin.com")
            print("Snatching Tokens...")

            login_page.html.render()
            login_cookies = login_page.cookies
            token = login_cookies.get("CSRF-TOKEN")
            print(f"Your current login token is {token}")

            headers = {
                "Referer": "https://crowdin.com/",
                # "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/94.0.4606.71 Safari/537.36",
            }
            payload = {
                "domain": "",
                "cname": "",
                "continue": "/translate/house-party/186/en-de?filter=basic&value=0",
                "locale": "en",
                "intended": "/auth/token",
                "_token": token,
                "email_or_login": "hpdownloader",
                "password": "HousePartyGame",
            }

            print("Logging in...")
            post = session.post(
                LOGINURL,
                data=payload,
                headers=headers,
                # proxies=proxies,
                # verify=False,
                cookies=login_cookies,
            )

            if post.status_code == 200 or 302:
                print("Login succesful!")
                print("Crowdin needs to rember us")
                # we need to first check the remember me "box"
                payload = {
                    "domain": "",
                    "cname": "",
                    "continue": "/translate/house-party/186/en-de?filter=basic&value=0",
                    "locale": "en",
                    "intended": "/auth/token",
                    "_token": token,
                }

                rember_cookies = post.cookies

                response = session.post(
                    "https://accounts.crowdin.com/remember-me/accept?continue=%2Ftranslate%2Fhouse-party%2F186%2Fen-de%3Ffilter%3Dbasic%26value%3D0",
                    data=payload,
                    headers=headers,
                    # proxies=proxies,
                    # verify=False,
                    cookies=rember_cookies,
                )

                if response.status_code == 200 or 302:
                    print("Crowdin rembered xD")
                    print("Beginning downloads now")
                    response.html.render()  # we need to remember first sometimes

                    button = response.html.find("editor_download_button")
                    print(button)
            else:
                print(
                    f"Crowdin is silly and gives a {post.status_code} because '{post.reason}'"
                )


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
