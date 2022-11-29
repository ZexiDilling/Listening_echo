import PySimpleGUI as sg
import configparser
import threading

from gui import main_layout, popup_controller
from excle_handler_echo_data import excel_controller
from helper_func import config_writer, config_header_to_list, progressbar
from e_mail import listening_controller


def main(config):

    window = main_layout()

    while True:

        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "-CLOSE-":
            break

        if event == "-ANALYSE-":
            data_location = sg.popup_get_folder("Where is the data located?")
            if data_location:
                file_name = sg.popup_get_text("File Name")
                if file_name:
                    save_location = sg.popup_get_folder("Save Location")
                    if save_location:
                        excel_controller(data_location, file_name, save_location)
                        sg.popup("Done")

        if event == "-LISTEN-":
            window["-KILL-"].update(value=False)
            threading.Thread(target=listening_controller, args=(config, True, window,), daemon=True).start()
            threading.Thread(target=progressbar, args=(True, window,), daemon=True).start()

        if event == "-KILL_BUTTON-":
            window["-KILL-"].update(value=True)

        if event == "In":
            config_heading = "Folder"
            sub_heading = "in"
            new_folder = sg.PopupGetFolder(f"Current folder: {config[config_heading][sub_heading]}", "Data Folder")
            if new_folder:
                data_dict = {sub_heading: new_folder}
                config_writer(config, config_heading, data_dict)

        if event == "Out":
            config_heading = "Folder"
            sub_heading = "out"
            new_folder = sg.PopupGetFolder(f"Current folder: {config[config_heading][sub_heading]}", "Data Folder")
            if new_folder:
                data_dict = {sub_heading: new_folder}
                config_writer(config, config_heading, data_dict)

        if event == "E-mail":
            config_header = "Email_list"
            table_data = config_header_to_list(config, config_header)

            headings = ["Name", "E-mail"]

            popup_controller(table_data, config, headings)

        if event == "Info":
            sg.Popup("To be written")

        if event == "About":
            sg.Popup("Echo Data Listening and analyses. Made By Charlie")


if __name__ == "__main__":
    config = configparser.ConfigParser()
    config.read("config.ini")
    main(config)
