import PySimpleGUI as sg
import configparser
import threading
from os import path
import time
from datetime import date

from gui import main_layout, popup_controller
from reports import skipped_well_controller
from helper_func import config_writer, config_header_to_list, clear_file
from e_mail import listening_controller, mail_report_sender


def main(config):
    """
    The main GUI setup and control for the whole program
    The while loop, is listening for button presses (Events) and will call different functions depending on
    what have been pushed.
    :param config: The config handler, with all the default information in the config file.
    :type config: configparser.ConfigParser
    :return:
    """

    window = main_layout()

    while True:

        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "-CLOSE-":
            break

        if event == "-ANALYSE-":
            data_location = sg.popup_get_folder("Where is the data located?")

            if not path.exists(config["Folder"]["out"]):
                folder_out = sg.PopupGetFolder("Please select the folder where your reports ends up")
                config_heading = "Folder"
                sub_heading = "out"
                data_dict = {sub_heading: folder_out}
                config_writer(config, config_heading, data_dict)

            save_location = config["Folder"]["out"]
            file_name = f"Report_{date.today()}"
            temp_counter = 2
            full_path = f"{save_location}/{file_name}.xlsx"
            while path.exists(full_path):
                temp_file_name = f"{file_name}_{temp_counter}"
                temp_counter += 1
                full_path = f"{save_location}/{temp_file_name}.xlsx"

            status = skipped_well_controller(data_location, full_path)
            sg.popup(status)

        if event == "-LISTEN-":
            window["-KILL-"].update(value=False)
            window["-PLATE_COUNTER-"].update(value=0)
            window["-E_MAIL_REPORT-"].update(value=True)

            if not window["-PLATE_NUMBER-"].get():
                window["-PLATE_NUMBER-"].update(value=0)

            # Clears out the temp_list file. to make sure only new data is in the file.
            clear_file("trans_list", config)

            if not path.exists(config["Folder"]["in"]):
                folder_in = sg.PopupGetFolder("Please select the folder you would like to listen to")
                config_heading = "Folder"
                sub_heading = "in"
                data_dict = {sub_heading: folder_in}
                config_writer(config, config_heading, data_dict)

            if not path.exists(config["Folder"]["out"]):
                folder_out = sg.PopupGetFolder("Please select the folder where your reports ends up")
                config_heading = "Folder"
                sub_heading = "out"
                data_dict = {sub_heading: folder_out}
                config_writer(config, config_heading, data_dict)

            threading.Thread(target=listening_controller, args=(config, True, window,), daemon=True).start()
            threading.Thread(target=progressbar, args=(config, True, window,), daemon=True).start()

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
            with open("README.txt") as file:
                info = file.read()

            sg.Popup(info)

        if event == "About":
            sg.Popup("Echo Data Listening and analyses. Made By Charlie")


def progressbar(config, run, window):
    """
    The progress bar, that shows the program working
    :param run: If the bar needs to be running or not
    :type run: bool
    :param window: Where the bar is displayed
    :type window: PySimpleGUI.PySimpleGUI.Window
    :return:
    """
    min = 0
    max = 100
    counter = 0

    # Timer for when too sent a report. if there are no files created for the periode of time, a report will be sent.
    time_limit = int(config["Time"]["timer_set"])

    while run:
        if counter == min:
            runner = "pos"
        elif counter == max:
            runner = "neg"
            # This is a setup to send a E-mail with a full report over all failed wells.
            # It is set up for time. #ToDo make it work with number
            if window["-E_MAIL_REPORT-"].get():
                current_time = time.time()
                try:
                    last_e_mail_time = float(window["-TIME_TEXT-"].get())
                except ValueError:
                    last_e_mail_time = time.time()

                if current_time - last_e_mail_time > time_limit and window["-E_MAIL_REPORT-"].get():
                    temp_file_name = "trans_list"
                    mail_report_sender(temp_file_name, window, config)
                    window["-E_MAIL_REPORT-"].update(value=False)


        if runner == "pos":
            counter += 10
        elif runner == "neg":
            counter -= 10

        window["-BAR-"].update(counter)

        time.sleep(0.1)
        if window["-KILL-"].get():
            run = False


if __name__ == "__main__":
    config = configparser.ConfigParser()
    config.read("config.ini")
    main(config)
