import PySimpleGUI as sg
import configparser
from threading import Thread
import time


from helper_func import config_writer
from reports import new_worklist


def _menu():
    """
    Top menu of the gui
    :return: The top menu
    :rtype: list
    """
    menu_top_def = [
        # ["&File", ["&Open    Ctrl-O", "&Save    Ctrl-S", "---", '&Properties', "&Exit", ]],
        ["&Listening", ["Folder", ["In", "Out", ], "E-mail", "reset"], ],
        ["&Help", ["Info", "About"]],
        ["Reports", ["Transfer", "Create Worklist", "Setup"]] #todo make these work
    ]
    layout = [[sg.Menu(menu_top_def)]]
    return layout


def _gui_main_layout():
    """
    The main layout for the gui
    :return: The main layout for the gui
    :rtype: list
    """
    tool_tip_plate_number = "This is the amount of Destination plates for the transfer"

    main = sg.Frame("Listening", [[
        sg.Column([
            [sg.ProgressBar(100, key="-BAR-", size=(25, 5), expand_x=True), sg.Checkbox("KILL", visible=False, key="-KILL-")],
            [sg.Button("Analyse", key="-ANALYSE-", expand_x=True,
                       tooltip="Choose a folder with all transfer files, and generates a report"),
             sg.Button("Listen", key="-LISTEN-", expand_x=True,
                       tooltip="starts the program that listen to the folder for files"),
             sg.Button("Kill", key="-KILL_BUTTON-", expand_x=True,
                       tooltip="stops the program that listen to the folder for files"),
             sg.Button("Close", key="-CLOSE-", expand_x=True,
                       tooltip="Closes the whole program")],
            [sg.Text("Plates:"),
             sg.Input("", key="-PLATE_NUMBER-", size=3,
                      tooltip=tool_tip_plate_number),
             sg.Text("Counter", key="-PLATE_COUNTER-", visible=True, tooltip="Plate analysed"),
             sg.Text("Error", key="-ERROR_PLATE_COUNTER-", visible=True, tooltip="Plates failed"),
             sg.Checkbox("Show Plate", key="-SHOW_PLATE_LIST-", enable_events=True,
                         tooltip="Will show a list of all the plates that have been transferred so far")],
            [sg.Checkbox("Transfer", key="-ADD_TRANSFER_REPORT_TAB-", visible=False),
             sg.Text(key="-TIME_TEXT-", visible=False), sg.Text(key="-INIT_TIME_TEXT-", visible=False)],

        ]),
        sg.VerticalSeparator(),
        sg.Column([
            [sg.Multiline(key="-TEXT_FIELD-", visible=False)],
            [sg.Checkbox("E-Mail Report", visible=False, key="-E_MAIL_REPORT-"),
             sg.Checkbox("Send E-mail", visible=False, key="-SEND_E_MAIL-")]
        ])
    ]])

    layout = [[main]]

    return layout


def main_layout():
    """
    The main setup for the layout for the gui
    :return: The setup and layout for the gui
    :rtype: sg.Window
    """

    # sg.theme()
    top_menu = _menu()

    layout = [[
        top_menu,
        _gui_main_layout()
    ]]

    return sg.Window("Echo Data", layout, finalize=True, resizable=True)


def _gui_popup_email_list(data, headings):
    """
    Layout for a popup menu
    :param data: The data that needs to be displayed
    :type data: list
    :param headings: The headings of the table where the data is displayed
    :type headings: list
    :return: The popup window
    :rtype: sg.Window
    """
    # headings = ["Source Plate", "Source Well", "Volume Needed", "Volume left", "Counters", "New Well"]

    col = sg.Frame("Table", [[
        sg.Column([
            [sg.Table(headings=headings, values=data, key="-TABLE-")],
            [sg.Button("Save", key="-TABLE_SAVE-"), sg.Button("Add", key="-TABLE_ADD-"),
             sg.Button("close", key="-TABLE_CLOSE-")]
        ])
    ]])

    layout = [[col]]

    return sg.Window("Table", layout, finalize=True, resizable=True)


def popup_email_list_controller(data, config, headings):
    """
    A popup menu
    :param data: The data that needs to be displayed
    :type data: list
    :param config: The config handler, with all the default information in the config file.
    :type config: configparser.ConfigParser
    :param headings: The headings of the table where the data is displayed
    :type headings: list
    :return:
    """
    window = _gui_popup_email_list(data, headings)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "-TABLE_CLOSE-":
            window.close()
            break

        if event == "-TABLE_SAVE-":
            config_data = {}
            table_data = window["-TABLE-"].get()
            for list in table_data:
                for data_index, data in enumerate(list):
                    if data_index == 0:
                        temp_name = data
                    if data_index == 1:
                        temp_email = data
                config_data[temp_name] = temp_email

            heading = "Email_list"
            config_writer(config, heading, config_data)

        if event == "-TABLE_ADD-":
            name = sg.PopupGetText("Name")
            email = sg.PopupGetText("Email")
            if name and email:
                table_data = window["-TABLE-"].get()
                table_data.append([name, email])
                window["-TABLE-"].update(values=table_data)


def _gui_popup_workilist(config):

    col = sg.Frame("Worklist setup", [[
        sg.Column([
            [sg.FolderBrowse("Plate Layout:", key="-WORKLIST_PLATE_LAYOUT_FOLDER-",
                             target="-WORKLIST_PLATE_LAYOUT_FOLDER_TARGET-",
                             tooltip="Choose a folder with Excel sheet describing what compound is in each well"),
             sg.Text(key="-WORKLIST_PLATE_LAYOUT_FOLDER_TARGET-")],
            [sg.FileBrowse("Transfer File:", key="-WORKLIST_TRANS_FILE-",
                           target="-WORKLIST_TRANS_FILE_TARGET-",
                           tooltip="Choose an excel file with all transfers for min one full set"),
             sg.Text(key="-WORKLIST_TRANS_FILE_TARGET-")],
            [sg.FolderBrowse("Surveys", key="-WORKLIST_SURVEY_FOLDER-",
                             target="-WORKLIST_SURVEY_FOLDER_TARGET-",
                             tooltip="Choose the folder with surveys for the source plate used to the worklist"),
             sg.Text(key="-WORKLIST_SURVEY_FOLDER_TARGET-")],
            [sg.Text("Set Amount:"),
             sg.Input(key="-WORKLIST_SET_AMOUNT-", tooltip="The amount of sets that worklist should include", size=5)],
            [sg.Text("Starting Set:"),
             sg.Input(1, key="-WORKLIST_STARTING_SET-", size=5,
                      tooltip="If you need a worklist starting from a higher number than 1")],
            [sg.Text("File Name:"),
             sg.Input(key="-WORKLIST_FILE_NAME-", tooltip="The name of the final excel file with the worklist")],
            [sg.Text("Dead Volume:"),
             sg.Text("LDV"), sg.Input(config["Dead_vol"]["ldv"], key="-WORKLIST_DEAD_VOL_LDV-", size=5,
                                      tooltip="dead volumen for the specific plate type (LDV) in uL"),
             sg.Text("PP"), sg.Input(config["Dead_vol"]["pp"], key="-WORKLIST_DEAD_VOL_PP-", size=5,
                                     tooltip="dead volumen for the specific plate type (PP) in uL")],
            [sg.Text("Specific Transfers"), sg.Button(key="-WORKLIST_SPECIFIC_TRANS-")],
            [sg.ProgressBar(max_value=100, key="-WORKLIST_PROGRESSBAR-", expand_x=True, orientation="horizontal")],
            [sg.Checkbox("Kill Progressbar", key="-WORKLIST_KILL-", visible=False)],
            [sg.Button("Generate", expand_x=True, key="-WORKLIST_GENERATE-"),
             sg.Button("Close", expand_x=True, key="-WORKLIST_CLOSE-")]

        ])
    ]])

    layout = [[col]]

    return sg.Window("Worklist", layout, finalize=True, resizable=True)


def popup_worklist_controller(config):

    window = _gui_popup_workilist(config)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "-WORKLIST_CLOSE-":
            window.close()
            break

        if event == "-WORKLIST_GENERATE-":
            if not values["-WORKLIST_PLATE_LAYOUT_FOLDER-"]:
                plate_layout_folder = sg.PopupGetFolder("Please select a folder with layout for all the source plates "
                                                        "included in the worklist")
            else:
                plate_layout_folder = values["-WORKLIST_PLATE_LAYOUT_FOLDER-"]

            if not values["-WORKLIST_TRANS_FILE-"]:
                trans_file = sg.FileBrowse("Choose an excel file with all transfers for min one full set")
            else:
                trans_file = values["-WORKLIST_TRANS_FILE-"]

            if not values["-WORKLIST_SET_AMOUNT-"]:
                try:
                    set_amount = int(sg.popup_get_text("How many sets should the worklist include?"))
                except ValueError:
                    set_amount = int(sg.popup_get_text("Please provide a number"))
            else:
                set_amount = int(values["-WORKLIST_SET_AMOUNT-"])

            if values["-WORKLIST_STARTING_SET-"]:
                starting_set = int(values["-WORKLIST_STARTING_SET-"])
            else:
                starting_set = None

            if not values["-WORKLIST_FILE_NAME-"]:
                file_name = sg.PopupGetText("Please provide a save file name")
            else:
                file_name = values["-WORKLIST_FILE_NAME-"]

            if not values["-WORKLIST_SURVEY_FOLDER-"]:
                survey_folder = sg.FolderBrowse("Choose the folder with surveys for the source plate used to the "
                                                "worklist")
            else:
                survey_folder = values["-WORKLIST_SURVEY_FOLDER-"]
            dead_vol = {"LDV": float(values["-WORKLIST_DEAD_VOL_LDV-"]), "PP": float(values["-WORKLIST_DEAD_VOL_PP-"])}

            specific_transfers = None
            save_location = config["Folder"]["worklist"]

            t1 = Thread(target=new_worklist, args=(survey_folder, plate_layout_folder, trans_file, set_amount,
                                                   dead_vol, save_location, file_name, specific_transfers,
                                                   starting_set, window), daemon=True)

            t2 = Thread(target=worklist_progressbar, args=(True, window,), daemon=False)
            t1.start()
            t2.start()


def worklist_progressbar(run, window):
    # Define minimum and maximum timer values
    min_timer = 0
    max_timer = 100

    # Initialize counter to 0
    counter = 0

    # Loop until run is False
    while run:

        # Check if counter has reached minimum or maximum timer values
        if counter == min_timer:
            runner = "pos"
        elif counter == max_timer:
            runner = "neg"

        # Increase or decrease counter based on the value of runner
        if runner == "pos":
            counter += 10
        elif runner == "neg":
            counter -= 10

        # Update the progress bar
        window["-WORKLIST_PROGRESSBAR-"].update(counter)

        # Wait for 100ms
        time.sleep(0.1)

        # Check if the worklist kill switch is activated
        if window["-WORKLIST_KILL-"].get():
            run = False


def _gui_popup_settings(config):

    config.read("config.ini")

    col_1 = sg.Column([
        [sg.Text("Time limit for no plate counter:"),
         sg.Input(default_text=config["Time"]["time_limit_no_plate_counter"],
                  key="-SETTINGS_TIME_LIMIT_NO_PLATE_COUNTER-", size=5)],
        [sg.Text("Time limit for plate counter:"),
         sg.Input(default_text=config["Time"]["time_limit_plate_counter"],
                  key="-SETTINGS_TIME_LIMIT_PLATE_COUNTER-", size=5)],
        [sg.Text("Dead vol LDV:", size=5),
         sg.Input(default_text=config["Dead_vol"]["ldv"], key="-SETTINGS_DEAD_VOL_LDV-", size=5),
         sg.Text("MAX vol LDV:", size=5), sg.Input(default_text=config["Max_vol"]["pp"],
                                                   key="-SETTINGS_MAX_VOL_LDV-", size=5)],
        [sg.Text("Dead vol PP:", size=5),
         sg.Input(default_text=config["Dead_vol"]["pp"], key="-SETTINGS_DEAD_VOL_PP-", size=5),
         sg.Text("MAX vol LDV:", size=5), sg.Input(default_text=config["Max_vol"]["ldv"],
                                                   key="-SETTINGS_MAX_VOL_PP-", size=5)],
        [sg.Button("Echo Error list", key="-SETTINGS_ECHO_ERROR_LIST_BUTTON-")],
        [sg.Button("Save", expand_x=True, key="-SETTINGS_SAVE-"),
         sg.Button("Close", expand_x=True, key="-SETTINGS_CLOSE-")]
    ])

    layout = [[sg.Frame("Settings", [[col_1]])]]

    return sg.Window("Table", layout, finalize=True, resizable=True)


def popup_settings_controller(config):
    window = _gui_popup_settings(config)

    while True:
        event, values = window.read()
        if event == "-SETTINGS_CLOSE-":
            break
        elif event == "-SETTINGS_SAVE-":
            try:
                time_limit_no_plate_counter = float(values["-SETTINGS_TIME_LIMIT_NO_PLATE_COUNTER-"])
                time_limit_plate_counter = float(values["-SETTINGS_TIME_LIMIT_PLATE_COUNTER-"])
                dead_vol_ldv = float(values["-SETTINGS_DEAD_VOL_LDV-"])
                dead_vol_pp = float(values["-SETTINGS_DEAD_VOL_PP-"])
                max_vol_ldv = float(values["-SETTINGS_MAX_VOL_LDV-"])
                max_vol_pp = float(values["-SETTINGS_MAX_VOL_PP-"])
            except ValueError:
                sg.popup("Please enter a valid float number")
                continue

            config.set("Time", "time_limit_no_plate_counter", str(time_limit_no_plate_counter))
            config.set("Time", "time_limit_plate_counter", str(time_limit_plate_counter))
            config.set("Dead_vol", "ldv", str(dead_vol_ldv))
            config.set("Dead_vol", "pp", str(dead_vol_pp))
            config.set("Max_vol", "ldv", str(max_vol_ldv))
            config.set("Max_vol", "pp", str(max_vol_pp))

            with open("config.ini", "w") as configfile:
                config.write(configfile)
            break

    window.close()
