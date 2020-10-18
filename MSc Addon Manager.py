# Python Base Libraries
import os.path
import pickle
import datetime
import concurrent.futures
import json
import zipfile

# external Libraries
import PySimpleGUI as Sg
import requests
import cloudscraper


# FUNCTIONS
# get and return current time
def get_time():
    time = datetime.datetime.now()
    time = time.strftime("%d.%m.%Y %H:%M:%S")
    return time


# Addon Update Main Function
def addon_updater():
    # entry = {name, download_link, version,}
    to_update_list = list()
    progress_bar = window['-PROGRESSMETER-']
    progress_text = window['-PROGRESSTEXT-']

    window['-LOGBOX-' + Sg.WRITE_ONLY_KEY].print(f"{get_time()}   Checking AddOns for newest version.")
    i = 0
    number_addons = len(config['addon_list'])
    progress_text.Update(f"{i:03d} / {number_addons:03d}")
    # create threading for getting newest addon information
    with concurrent.futures.ThreadPoolExecutor() as executor:
        results = [executor.submit(get_addon_request,
                                   addon_name=addon_name,
                                   addon_link=addon_info["link"]) for addon_name, addon_info in config['addon_list'].items()]

        for f in concurrent.futures.as_completed(results):
            i += 1
            progress_text.Update(f"{i:03d} / {number_addons:03d}")
            addon_name, download_link, addon_version = f.result()
            progress_bar.UpdateBar(i / len(config['addon_list']) * 1000)

            # noinspection PyTypeChecker
            if "version" not in config['addon_list'][addon_name] or config['addon_list'][addon_name]['version'] != addon_version:
                window['-LOGBOX-' + Sg.WRITE_ONLY_KEY].print(f"{get_time()}   Addon {addon_name} has a new version v#{addon_version}")
                to_update_list.append({
                    "name": addon_name,
                    "download_link": download_link,
                    "version": addon_version,
                })

    # print addons to update:
    window['-LOGBOX-' + Sg.WRITE_ONLY_KEY].print(f"\n{get_time()}   {len(to_update_list)} AddOns will need an update.\n")

    # token, agent = cloudscraper.get_tokens(url="https://www.curseforge.com/")
    scraper = cloudscraper.create_scraper()  # returns a CloudScraper instance

    j = 0
    number_updates = len(to_update_list)
    progress_text.Update(f"{j:03d} / {number_updates:03d}")
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        results = [executor.submit(download_and_extract_addon_file, wow_path=config['wow_path'], name=addon["name"], download_link=addon["download_link"], version=addon["version"], scraper=scraper) for addon in to_update_list]
        for f in concurrent.futures.as_completed(results):
            j += 1
            addon_name, addon_version = f.result()
            window['-LOGBOX-' + Sg.WRITE_ONLY_KEY].print(f"{get_time()}   Finished updating {addon_name} to v#{addon_version}")
            progress_text.Update(f"{j:03d} / {number_updates:03d}")
            progress_bar.UpdateBar(j / number_updates * 1000)
            config['addon_list'][addon_name]["version"] = addon_version

    # remove downloaded files
    for root, dirs, files in os.walk(download_folder):
        for dowm_file in files:
            os.remove(os.path.join(root, dowm_file))

    return


# Contact Curseforge API to get newest versions and links
def get_addon_request(addon_name, addon_link):
    download_link = None
    addon_version = None

    if "curseforge" in addon_link:
        api_link = addon_link.replace("https://www.curseforge.com/", "https://api.cfwidget.com/")
        r = requests.get(api_link)
        r_json = r.content
        addon_json = json.loads(r_json.decode('utf-8'))

        if "error" in addon_json and addon_json["error"] == "in_queue":
            return addon_name, None, None

        if "versions" not in addon_json:
            return addon_name, None, None
        elif type(addon_json["versions"]) == list:
            return addon_name, None, None

        addon_files = list()
        for version, subfiles in addon_json["versions"].items():
            for subfile in subfiles:
                addon_files.append(subfile)

        addon_files = sorted(addon_files, key=lambda k: k['id'], reverse=True)

        for addon in addon_files:
            if addon["version"].startswith("1") and config['classic_retail'] == "classic":
                if config['prefered_release_type'] == "release" and addon["type"] == "release":
                    addon_version = addon["id"]
                    download_link = addon["url"].replace("files", "download")
                    break
                elif config['prefered_release_type'] == "beta" and addon["type"] in ["release", "beta"]:
                    addon_version = addon["id"]
                    download_link = addon["url"].replace("files", "download")
                    break
                elif config['prefered_release_type'] == "alpha":
                    addon_version = addon["id"]
                    download_link = addon["url"].replace("files", "download")
                    break
            elif not addon["version"].startswith("1"):
                if config['prefered_release_type'] == "release" and addon["type"] == "release":
                    addon_version = addon["id"]
                    download_link = addon["url"].replace("files", "download")
                    break
                elif config['prefered_release_type'] == "beta" and addon["type"] in ["release", "beta"]:
                    addon_version = addon["id"]
                    download_link = addon["url"].replace("files", "download")
                    break
                elif config['prefered_release_type'] == "alpha":
                    addon_version = addon["id"]
                    download_link = addon["url"].replace("files", "download")
                    break

    elif "tukui" in addon_link:
        addon_name = "TUKUI/ELVUI"
        r = requests.get(addon_link)
        content = r.text

        start = 0
        end = content.find(".zip") + 4
        for i in range(end - 1, 0, -1):
            if content[i] == '"':
                start = i + 1
                break
        download_link = f"https://www.tukui.org{content[start:end]}"
        addon_version = download_link[download_link.find("-") + 1:download_link.find(".zip")]

    return addon_name, download_link, addon_version


# Download addon-file and extract it to wow-folder
def download_and_extract_addon_file(wow_path, name, download_link, version, scraper):
    addon_download_file = f"{download_folder}{name}-{version}.zip"

    if not os.path.isfile(addon_download_file):
        if "curseforge" in download_link:
            download_link = download_link.replace("files", "download")
            download_link = download_link + "/file"

            # cloudscraper
            # scraper = cloudscraper.create_scraper()  # returns a CloudScraper instance
            addon_file = scraper.get(download_link)

            with open(addon_download_file, 'wb') as f:
                f.write(addon_file.content)

        if "tukui" in download_link:
            addon_file = requests.get(download_link)
            with open(addon_download_file, 'wb') as f:
                f.write(addon_file.content)

    if os.path.isdir(wow_path):
        with zipfile.ZipFile(addon_download_file, 'r') as zip_ref:
            zip_ref.extractall(wow_path)

    return name, version


# Link-Wrapper
def link_wrapper(addon_link):
    # Curseforge
    # https://www.curseforge.com/wow/addons/details
    if "curseforge" in addon_link:
        link_args = addon_link.split("/")
        for i in range(len(link_args)):
            if link_args[i] == "addons":
                addon_name = link_args[i+1]
                return addon_name, addon_link

    # ElvUI
    # https://www.tukui.org/download.php?ui=elvui
    if "tukui" in addon_link:
        addon_name = "TUKUI/ELVUI"
        return addon_name, addon_link


def update_config_with_addonlist(w_values):
    box_addon_list = w_values["-AL-"].split("\n")
    box_addon_list = [x for x in box_addon_list if x]
    box_addon_name_list = list()

    for box_addon_link in box_addon_list:
        box_addon_name, box_link = link_wrapper(addon_link=box_addon_link)
        box_addon_name_list.append(box_addon_name)
        if box_addon_name in config["addon_list"]:
            config["addon_list"][box_addon_name]['link'] = box_link
        else:
            config["addon_list"][box_addon_name] = {
                "link": box_link,
                "version": None,
            }

    to_del = list()
    for c_addon_name, c_addon_info in config["addon_list"].items():
        if c_addon_name not in box_addon_name_list:
            to_del.append(c_addon_name)

    for del_addon in to_del:
        del config["addon_list"][del_addon]

    with open('config.pickle', 'wb') as c_file:
        pickle.dump(config, c_file)


# MAIN SCRIPT
# try to read config, else start with default
global config
try:
    home_path = os.getcwd()
    with open(f'{home_path}\\config.pickle', 'rb') as file:
        config = pickle.load(file)
except Exception as e:
    config = {
        "wow_path": "",
        "addon_list": {},
        "prefered_release_type": "release",
        "classic_retail": "retail"
    }

# create download folder
download_folder = os.getcwd() + "\\downloads\\"
if not os.path.isdir(download_folder):
    os.makedirs(download_folder)

# get list of addonlinks to display in addon-box
addon_display_string = ""
for cbox_addon_name, box_addon_info in config["addon_list"].items():
    addon_display_string += f"{box_addon_info['link']}\n"

# radio1: Set default Release_Type
prt_r = True if config["prefered_release_type"] == "release" else False
prt_b = True if config["prefered_release_type"] == "beta" else False
prt_a = True if config["prefered_release_type"] == "alpha" else False

# radio02: Set default WoW-Version
cr_r = True if config["classic_retail"] == "retail" else False
cr_c = True if config["classic_retail"] == "classic" else False

# Create left part of AddOn-Manager
config_column = [
    [
        Sg.Text("WoW/Interface/AddOn Folder"),
        Sg.In(config["wow_path"], size=(50, 1), enable_events=True, key="-FOLDER-"),
        Sg.FolderBrowse(),
    ],
    [
        Sg.HSeparator()
    ],
    [
        Sg.Text("Select Release-Type and paste a list of addon-links (one addon per line)"),
    ],
    [
        Sg.Radio("Retail", "RADIO02", default=cr_r, enable_events=True, key="-CR_R-"),
        Sg.Radio("Classic", "RADIO02", default=cr_c, enable_events=True, key="-CR_C-"),
    ],
    [
        Sg.Radio("Release", "RADIO01", default=prt_r, enable_events=True, key="-RT_R-"),
        Sg.Radio("beta", "RADIO01", default=prt_b, enable_events=True, key="-RT_B-"),
        Sg.Radio("alpha", "RADIO01", default=prt_a, enable_events=True, key="-RT_A-"),
    ],
    [
        Sg.Button('Save AddOn-List', enable_events=True, key="-ALU-"),
    ],
    [
        Sg.Multiline(addon_display_string, size=(75, 50), key="-AL-"),
    ],
]

control_column = [
    [
        Sg.Button('Update Addons', enable_events=True, key="-UAA-"),
        Sg.Button('Close AddOn Manager', enable_events=True, key="-CLOSE-", button_color=('black', '#bb0000')),
     ],
    [
        Sg.Multiline("", size=(75, 56), enable_events=True, key="-LOGBOX-"+Sg.WRITE_ONLY_KEY, auto_refresh=True, autoscroll=True),
    ],
    [
        Sg.ProgressBar(max_value=1000, size=(44, 25), key='-PROGRESSMETER-'),
        Sg.Text("000 / 000", size=(8, 1), key='-PROGRESSTEXT-'),
    ],
]

layout = [
    [
        Sg.Column(config_column),
        Sg.VSeperator(),
        Sg.Column(control_column),
    ]
]

# Create the window
window = Sg.Window("MSc Addon Manager", layout, margins=(0, 0))

# Create an event loop
while True:
    event, values = window.read(timeout=100)

    # if event == new Folder selection
    if event == "-FOLDER-":
        config["wow_path"] = values['-FOLDER-']

    # if event == new prefered release type selected
    elif event == "-RT_R-" or event == "-RT_B-" or event == "-RT_A-":
        if values["-RT_R-"]:
            config["prefered_release_type"] = "release"
        elif values["-RT_B-"]:
            config["prefered_release_type"] = "beta"
        elif values["-RT_A-"]:
            config["prefered_release_type"] = "alpha"
        # save config
        with open('config.pickle', 'wb') as file:
            pickle.dump(config, file)

    # if event == new wow_version selected
    elif event == "-CR_R-" or event == "-CR_C-":
        if values["-CR_R-"]:
            config["classic_retail"] = "retail"
        elif values["-CR_C-"]:
            config["classic_retail"] = "classic"
        # save config
        with open('config.pickle', 'wb') as file:
            pickle.dump(config, file)

    # if event == change in the addon-list
    elif event == "-ALU-":
        update_config_with_addonlist(w_values=values)

    # if event == Update All Addons
    elif event == "-UAA-":
        update_config_with_addonlist(w_values=values)

        window['-LOGBOX-' + Sg.WRITE_ONLY_KEY].print(f"{get_time()}   Updating all AddOns:")
        addon_updater()
        with open('config.pickle', 'wb') as file:
            pickle.dump(config, file)
        window['-LOGBOX-' + Sg.WRITE_ONLY_KEY].print(f"\n{get_time()}   All AddOns have been updated\n")

    # End program if user closes window or presses the OK button
    elif event == "-CLOSE-" or event == Sg.WIN_CLOSED:
        # save config
        with open('config.pickle', 'wb') as file:
            pickle.dump(config, file)

        # end programm
        break


window.close()

# py -3.8-64 -m PyInstaller --onefile --paths="C:\Users\nicla\Owncloud2\Documents\GitHub\MSc Addon Manager\venv\Lib\site-packages" "MSc Addon Manager.py"
