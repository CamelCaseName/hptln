# This is a shitty script.
import csv
import os
import subprocess
import sys
import warnings
import json

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


def convert_files(args):

    print(f"Converting files in {args[1]}")
    path = args[1]
    files = []
    tnpath = path + "\\converted"

    for root, _, dirfiles in os.walk(path):
        for name in dirfiles:
            files.append(os.path.join(root, name)[len(path) :])
        for _ in root:
            if not os.path.isdir(tnpath + "\\" + root[len(path) :]):
                os.makedirs(tnpath + "\\" + root[len(path) :])

    if not os.path.isdir(tnpath):
        os.makedirs(tnpath)

    file_counter = 0
    error_occurred = 0

    for person in files:
        f_name, f_ext = os.path.splitext(path + person)
        if f_ext == ".xlsx":
            file_counter += 1
            print(".", end="")
            try:
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
                for lines in fin.readlines():
                    lines = re.sub(r"\|{1}\n", "\n", lines)
                    lines = re.sub(r"\|* *\|", "|", lines)
                    lines = re.sub(r"\|\|", "|", lines)
                    adj_lines.append(lines)
                fin.close()
                fout = open(f_name + ".txt", "w", encoding="utf-8")
                fout.write("\ufeff")
                fout.writelines(adj_lines)
            except ValueError:
                error_occurred = 1
    print(f"\nConverted {file_counter} files!")
    return error_occurred


global_args = sys.argv  # 0 is the file itself, 1 the deciding factor
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

# convert files
if os.path.exists(global_args[1]):
    convert_files(global_args)

# download first, then convert
elif len(global_args) > 2:
    dfile_counter = 0
    dlpath = global_args[2]
    lang = global_args[1]

    if languages.isin([lang]).any():
        language = ""
        if lang == "cs":
            language = "Czech"
        elif lang == "da":
            language = "Danish"
        elif lang == "de":
            language = "German"
        elif lang == "nl":
            language = "Dutch"
        elif lang == "fi":
            language = "Finnish"
        elif lang == "fr":
            language = "French"
        elif lang == "hu":
            language = "Hungarian"
        elif lang == "it":
            language = "Italian"
        elif lang == "ja":
            language = "Japanese"
        elif lang == "ko":
            language = "Korean"
        elif lang == "pl":
            language = "Polish"
        elif lang == "pt":
            language = "Portuguese"
        elif lang == "ptbr":
            language = "Portuguese, Brazilian"
        elif lang == "es":
            language = "Spanish"
        elif lang == "esmx":
            language = "Spanish, Mexico"
        elif lang == "tr":
            language = "Turkish"

        print("Connecting to crowdin.com...")
        session = HTMLSession()
        REQUESTURL = (
            "https://crowdin.com/translate/house-party/158/en-de?filter=basic&value=0"
        )
        LOGINURL = "https://accounts.crowdin.com/login"

        login_page = session.get(
            LOGINURL,
        )

        if login_page.status_code == 200:
            print("Connected to crowdin.com")
            print("Snatching tokens...")

            login_page.html.render()
            login_cookies = login_page.cookies
            token = login_cookies.get("CSRF-TOKEN")
            print(f"Your current login token is {token}")

            headers = {
                "Referer": "https://crowdin.com/",
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
                    cookies=rember_cookies,
                )

                if response.status_code == 200 or 302:
                    print("Crowdin rembered xD")
                    print("Downloading Files")
                    download_payload = {
                        "as_xliff": "",
                        "task": "",
                    }
                    download_cookies = response.cookies

                    for i in [
                        200,
                        214,
                        198,
                        202,
                        212,
                        206,
                        208,
                        210,
                        148,
                        146,
                        144,
                        150,
                        154,
                        178,
                        176,
                        158,
                        188,
                        180,
                        182,
                        186,
                        164,
                        162,
                        172,
                        166,
                        168,
                        170,
                        160,
                        190,
                        174,
                        184,
                        192,
                        220,
                        218,
                        136,
                        216,
                        138,
                        140,
                    ]:
                        download = session.post(
                            f"https://crowdin.com/backend/project/house-party/{lang}/{i}/export",
                            json=download_payload,
                            headers=headers,
                            cookies=download_cookies,
                        )
                        if download.status_code == 200:
                            file_cookies = download.cookies
                            download_data_url_dict = json.loads(download.html.html)
                            file_download_url = download_data_url_dict.get("url")
                            file_download = session.get(
                                file_download_url, headers=headers, cookies=file_cookies
                            )
                            if file_download.status_code == 200:
                                dfile_counter += 1
                                print(".", end="")

                                subpath = "Languages"

                                if i in [200, 214, 198, 202]:
                                    subpath = "Languages\\A Vickie Vixen Valentine\\"
                                elif i in [212, 210, 206, 208]:
                                    subpath = "Languages\\Combat Training\\"
                                elif i in [146, 150, 144, 148]:
                                    subpath = "Languages\\Date Night with Brittney\\"
                                elif i == 154:
                                    subpath = "Hints\\"
                                elif i in [
                                    178,
                                    176,
                                    158,
                                    188,
                                    180,
                                    182,
                                    186,
                                    164,
                                    162,
                                    172,
                                    166,
                                    168,
                                    170,
                                    160,
                                    190,
                                    174,
                                    184,
                                    192,
                                ]:
                                    subpath = "Languages\\Original Story\\"
                                elif i in [220, 218, 136, 216, 138, 140]:
                                    subpath = "Languages\\"

                                if i != 154:
                                    if not os.path.exists(
                                        os.path.join(dlpath, subpath, language)
                                    ):
                                        os.makedirs(
                                            os.path.join(dlpath, subpath, language)
                                        )

                                    dlfout = open(
                                        os.path.join(
                                            dlpath,
                                            subpath,
                                            language,
                                            re.findall(
                                                r"((\"(\w*|\d*|\s*)+)(\.((xlsx){1})))",
                                                file_download.headers.get(
                                                    "Content-Disposition"
                                                ),
                                            )[0][0][1:],
                                        ),
                                        "wb",
                                    )
                                else:
                                    if not os.path.exists(
                                        os.path.join(dlpath, subpath)
                                    ):
                                        os.makedirs(os.path.join(dlpath, subpath))

                                    dlfout = open(
                                        os.path.join(
                                            dlpath,
                                            subpath,
                                            language + ".xlsx",
                                        ),
                                        "wb",
                                    )
                                dlfout.write(file_download.content)
                            else:
                                print(
                                    f"Crowdin is silly and gives a {file_download.status_code} because '{file_download.reason}'"
                                )
                        else:
                            print(
                                f"Crowdin is silly and gives a {download.status_code} because '{download.reason}'"
                            )
                    print(f"\nDownloaded {dfile_counter} files, converting next")

                    if convert_files(["easter egg", global_args[2]]):
                        print(
                            "One file could not automatically be converted, it is advised you run the script in CONVERSION only mode again,"
                        )
                        print(
                            f"often it works then, if not, it is probably the PauseMenu.xlsx file. So >py.exe ./hptlndir.py {dlpath}"
                        )

                else:
                    print(
                        f"Crowdin is silly and gives a {response.status_code} because '{response.reason}'"
                    )
            else:
                print(
                    f"Crowdin is silly and gives a {post.status_code} because '{post.reason}'"
                )
        else:
            print(
                f"Crowdin is silly and gives a {login_page.status_code} because '{login_page.reason}'"
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
    print("Currently available languages:")
    for l in languages:
        print(l, end=", ")
        print(" ")
