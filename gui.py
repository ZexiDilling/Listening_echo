import PySimpleGUI as sg
from helper_func import config_writer


def menu():
    menu_top_def = [
        # ["&File", ["&Open    Ctrl-O", "&Save    Ctrl-S", "---", '&Properties', "&Exit", ]],
        ["&Listening", ["Folder", ["In", "Out", ], "E-mail"], ],
        ["&Help", ["Info", "About"]],
    ]
    layout = [[sg.Menu(menu_top_def)]]
    return layout


def gui_main_layout():
    main = sg.Frame("Listening", [[
        sg.Column([
            [sg.ProgressBar(100, key="-BAR-", size=(25, 5)), sg.Checkbox("KILL", visible=False, key="-KILL-")],
            [sg.Button("Analyse", key="-ANALYSE-"), sg.Button("Listen", key="-LISTEN-"),
             sg.Button("Kill", key="-KILL_BUTTON-"), sg.Button("Close", key="-CLOSE-")]
        ])
    ]])

    layout = [[main]]

    return layout


def main_layout():

    # sg.theme()
    top_menu = menu()

    layout = [[
        top_menu,
        gui_main_layout()
    ]]

    return sg.Window("Echo Data", layout, finalize=True, resizable=True)


def gui_popup_table(data, headings):
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
    window = gui_popup_table(data, headings)

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




