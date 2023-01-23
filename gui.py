import PySimpleGUI as sg
from helper_func import config_writer


def _menu():
    """
    Top menu of the gui
    :return: The top menu
    :rtype: list
    """
    menu_top_def = [
        # ["&File", ["&Open    Ctrl-O", "&Save    Ctrl-S", "---", '&Properties', "&Exit", ]],
        ["&Listening", ["Folder", ["In", "Out", ], "E-mail"], ],
        ["&Help", ["Info", "About"]],
        ["Reports", ["Transfer", "setup"]]
    ]
    layout = [[sg.Menu(menu_top_def)]]
    return layout


def _gui_main_layout():
    """
    The main layout for the gui
    :return: The main layout for the gui
    :rtype: list
    """
    main = sg.Frame("Listening", [[
        sg.Column([
            [sg.ProgressBar(100, key="-BAR-", size=(25, 5)), sg.Checkbox("KILL", visible=False, key="-KILL-")],
            [sg.Button("Analyse", key="-ANALYSE-"), sg.Button("Listen", key="-LISTEN-"),
             sg.Button("Kill", key="-KILL_BUTTON-"), sg.Button("Close", key="-CLOSE-")],
            [sg.Checkbox("Transfer", key="-ADD_TRANSFER_REPORT_TAB-", visible=False),
             sg.Text(key="-TIME_TEXT-", visible=False), sg.Text(key="-INIT_TIME_TEXT-", visible=False)],
            [sg.Text("Plate before sending the report"), sg.Checkbox("E-Mail Report", visible=False, key="-E_MAIL_REPORT-"),
             sg.Input("", key="-PLATE_NUMBER-", size=3), sg.Text("Counter", key="-PLATE_COUNTER-", visible=True)]
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


def _gui_popup_table(data, headings):
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


def popup_controller(data, config, headings):
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
    window = _gui_popup_table(data, headings)

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




