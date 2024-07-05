import PySimpleGUI as sg

import threading
from os import path
import time
from datetime import date

from gui import main_layout, popup_email_list_controller, popup_settings_controller, popup_worklist_controller
from reports import skipped_well_controller
from helper_func import config_writer, config_header_to_list, clear_file
from e_mail import listening_controller, mail_report_sender, mail_estimated_time


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

            overview_data = skipped_well_controller(data_location, full_path, config)
            popup_answer = sg.PopupYesNo("send the report to E-mail list?")
            if popup_answer.casefold() == "yes":
                mail_report_sender(full_path, window, config, overview_data)

            sg.popup(overview_data)

        if event == "-LISTEN-":
            window["-KILL-"].update(value=False)
            window["-PLATE_COUNTER-"].update(value=0)
            window["-ERROR_PLATE_COUNTER-"].update(value=0)
            window["-E_MAIL_REPORT-"].update(value=True)
            window["-TEXT_FIELD-"].update(value="")

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

        if event == "-SHOW_PLATE_LIST-":
            window["-TEXT_FIELD-"].update(visible=values["-SHOW_PLATE_LIST-"])

        if event == "reset":
            window["-PLATE_COUNTER-"].update(value=0)
            window["-ERROR_PLATE_COUNTER-"].update(value=0)
            window["-TIME_TEXT-"].update(value="")
            window["-INIT_TIME_TEXT-"].update(value="")
            window["-ADD_TRANSFER_REPORT_TAB-"].update(value=False)
            window["-TEXT_FIELD-"].update(value="")
            window["-E_MAIL_REPORT-"].update(value=False)
            window["-SEND_E_MAIL-"].update(value=False)

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

            popup_email_list_controller(table_data, config, headings)

        if event == "Info":
            with open("README.txt") as file:
                info = file.read()

            sg.Popup(info)

        if event == "About":
            sg.Popup("Echo Data Listening and analyses. Programmed By Charlie for DTU SCore")

        if event == "Transfer":
            sg.Popup("Not working... and not sure what it should do :D ")

        if event == "Setup":
            popup_settings_controller(config)

        if event == "Create Worklist":
            popup_worklist_controller(config)


def progressbar(config, run, window):
    """
    The progress bar, that shows the program working
    :param run: If the bar needs to be running or not
    :type run: bool
    :param window: Where the bar is displayed
    :type window: PySimpleGUI.PySimpleGUI.Window
    :return:
    """
    min_timer = 0
    max_timer = 100
    counter = 0

    # Timer for when too sent a report. if there are no files created for the period of time, a report will be sent.
    # set one for runs where there is not set a plate counter, or if the platform fails.
    # set one for if plate counter is used. To avoid sending multiple report files, one for each source plate
    time_limit_no_plate_counter = float(config["Time"]["time_limit_no_plate_counter"])
    time_limit_plate_counter = float(config["Time"]["time_limit_plate_counter"])

    temp_file_name = "trans_list"
    total_plates = int(window["-PLATE_NUMBER-"].get())
    procent_splitter = [
        round(total_plates / 100 * 10),
        round(total_plates / 100 * 25),
        round(total_plates / 100 * 50),
        round(total_plates / 100 * 75)
    ]
    time_estimates_send = []

    while run:
        current_time = time.time()
        if counter == min_timer:
            runner = "pos"
        elif counter == max_timer:
            runner = "neg"
            # This is a setup to send a E-mail with a full report over all failed wells.
            # It is set up for time.
            if window["-E_MAIL_REPORT-"].get():
                try:
                    last_e_mail_time = float(window["-TIME_TEXT-"].get())
                except ValueError:
                    last_e_mail_time = time.time()

                if current_time - last_e_mail_time > time_limit_no_plate_counter and window["-E_MAIL_REPORT-"].get():
                    mail_report_sender(temp_file_name, window, config)
                    window["-E_MAIL_REPORT-"].update(value=False)

            if window["-SEND_E_MAIL-"].get():
                if current_time - last_e_mail_time > time_limit_plate_counter and window["-E_MAIL_REPORT-"].get():
                    mail_report_sender(temp_file_name, window, config)
                    window["-E_MAIL_REPORT-"].update(value=False)

            if total_plates >= int(config["Plate_setup"]["limit"]):
                current_plate = int(window["-PLATE_COUNTER-"].get())

                if current_plate in procent_splitter:
                    if current_plate not in time_estimates_send:
                        time_estimates_send.append(current_plate)
                        elapsed_time = current_time - float(window["-INIT_TIME_TEXT-"].get())
                        mail_estimated_time(config, total_plates, current_plate, elapsed_time)

        if runner == "pos":
            counter += 10
        elif runner == "neg":
            counter -= 10

        window["-BAR-"].update(counter)

        time.sleep(0.1)
        if window["-KILL-"].get():
            run = False


if __name__ == "__main__":
    import configparser
    config = configparser.ConfigParser()
    config.read("config.ini")
    main(config)
