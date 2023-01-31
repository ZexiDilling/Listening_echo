import configparser
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
import smtplib
from email.message import EmailMessage
import os
from os import path
from datetime import date

from get_data import get_xml_trans_data_skipping_wells
from helper_func import write_temp_list_file, read_temp_list_file
from reports import skipped_well_controller


class MyEventHandler(FileSystemEventHandler):
    def __str__(self):
        """This is a standard class for watchdog.
        This is the class that is listening for files being created, moved or deleted.
        ATM the system only react to newly created files"""

    def __init__(self, window):
        self.config = configparser.ConfigParser()
        self.config.read("config.ini")
        self.window = window

    def on_created(self, event):
        """
        This event is triggered when a new file appears in the target folder
        It checks the file in the event for missing transferees, if there are any, it sends an E-mail.
        :param event: The full event, including the path to the file that have been created
        """
        # The list for all the trans files
        temp_file_name = "trans_list"

        # plate list:
        all_plates = self.window["-TEXT_FIELD-"].get()
        if all_plates:
            plate_list = all_plates.split(",")
        else:
            plate_list = []
        current_plate = len(plate_list)

        # checks if path is a directory
        if os.path.isfile(event.src_path):
            temp_file = event.src_path

            if "transfer" in temp_file.casefold():
                # Gets time-code for when file was created.
                # Is used to send a full report after x-amount of time
                self.window["-TIME_TEXT-"].update(value=time.time())
                # sets time-code for the first E-mail:
                if not self.window["-INIT_TIME_TEXT-"].get():
                    self.window["-INIT_TIME_TEXT-"].update(value=time.time())

                # Writes the file name to a list, to be used for creating a report for all the trans-files
                write_temp_list_file(temp_file_name, temp_file, self.config)

                # a counter for setting max time between sending E-mails.
                counter = 0
                sending_mail = True
                while sending_mail:
                    # Set timer to sleep while echo is finishing writing data to the files
                    time.sleep(2)
                    try:
                        all_data, skipped_wells, skip_well_counter, _, _, _, _, destination_plate = \
                            get_xml_trans_data_skipping_wells(temp_file)
                    except:
                        counter += 1
                        if counter == 30:
                            sending_mail = False
                    else:
                        # Check if there are any well being skipped in the file
                        if skipped_wells:

                            # set-up the E-mail
                            msg_subject = f"Liquid Transferee failed for {skip_well_counter} wells"
                            data = all_data
                            e_mail_type = "error"

                            # send an E-mail with information from the trans file
                            mail_setup(msg_subject, data, self.config, e_mail_type)

                        # count destination plates
                        if destination_plate not in plate_list:
                            plate_list.append(destination_plate)

                            # Is used for sending a report after x-amount.
                            # Should fit with the numbers of destination plates in the run.
                            current_plate = len(plate_list)
                            self.window["-PLATE_COUNTER-"].update(value=current_plate)

                            # Update the list of plates
                            all_plates += f",{destination_plate}"
                            self.window["-TEXT_FIELD-"].update(value=all_plates)

                        sending_mail = False

                    # Check plate amount. If it reach set amount, it will create a report over all the files and send it.
                    if current_plate == int(self.window["-PLATE_NUMBER-"].get()):
                        self.window["-SEND_E_MAIL-"].update(values=True)
                        # mail_report_sender(temp_file_name, self.window, self.config)
                        # self.window["-E_MAIL_REPORT-"].update(value=False)

        else:
            print(event.src_path)
            print("folder is created")


    # def on_deleted(self, event):
    #     """
    #     This event is triggered when a file is removed from the folder, either by deletion or moved.
    #     :param event:
    #     :return:
    #     """
    #     print("delet")
    #     print(event)

    # def on_modified(self, event):
    #     """
    #     This event is triggered when a file is modified.
    #     :param event:
    #     :return:
    #     """
    #     print("mod")
    #     print(event)


def _mail_error(all_data, config):
    """
    Function to send email notification when errors occur during data transfer.
    :param all_data: List containing information about the missing wells
    :type all_data: list
    :param config: The config handler, with all the default information in the config file.
    :type config: configparser.ConfigParser
    :return: A string containing the error message to be sent as an email
    :rtype: str
    """
    # Amount of missing wells
    missing_wells = int(all_data[0])
    trans_string = ""

    # Going through every single missing well and writes the error msg
    for count in range(missing_wells):
        # Get the transferee name and error message
        transferee = all_data[4 + (count * 2)]
        error = all_data[5 + (count * 2)]

        # Extract error code from the error message
        error_code = error[9:18]

        # Build the error message string for each missing well
        trans_string += f"Transferee: {transferee} - Error: {error}"

        # Check if the error code is present in the config dict
        error_description = config.get("Echo_error", {}).get(error_code)
        if error_description:
            trans_string += f" - {error_description}\n"
        else:
            trans_string += " - New error YAY!!!\n"

    # combine all details into one body of text
    body = f"Missing {all_data[0]} Wells \n" \
           f"Source plate {all_data[2].split(',')[-1]}\n" \
           f"Destination plate: {all_data[3].split(',')[-1]}\n" \
           f"{trans_string}"

    return body


def _mail_final_report(overview_data, config):
    """
    Writes the body of the E-mail
    :param overview_data: An overview of all the data generated.
    :type overview_data: dict
    :param config: The config handler, with all the default information in the config file.
    :type config: configparser.ConfigParser
    :return:
    """

    body = \
        f"Hey SCore people!\n" \
        f"This is the report of {overview_data['plate_amount']}-plates\n" \
        f"{overview_data['amount_complete_plates']} have completed with zero failed transferes:\n" \
        f"{overview_data['amount_failed_plates']} plates have {overview_data['failed_trans']} failed transferes\n" \
        f"There are {overview_data['failed_wells']} failed wells on {overview_data['amount_source_plates']} " \
        f"source plates \n" \
        f"Time taken from first transfer to last is: {overview_data['time_for_all_trans']}.\n\n" \
        f"The best wishes\n" \
        f"The Echo :D"

    return body


def mail_report_sender(temp_file_name, window, config):
    """
    This function sends the final report of the transfer operation.

    :param temp_file_name: The name of the temporary file where all transfer data is stored.
    :type temp_file_name: str
    :param window: The GUI window
    :type window: PySimpleGUI.PySimpleGUI.Window
    :param config: The config handler, with all the default information in the config file.
    :type config: configparser.ConfigParser
    :return:
    """
    # Reads the temp_file where all the trans file have been written to

    file_list = read_temp_list_file(temp_file_name, config)

    # Setup the report
    report_name = f"Report_{date.today()}"
    save_location = config["Folder"]["out"]
    temp_counter = 2
    full_path = f"{save_location}/{report_name}.xlsx"
    while path.exists(full_path):
        temp_report_name = f"{report_name}_{temp_counter}"
        temp_counter += 1
        full_path = f"{save_location}/{temp_report_name}.xlsx"

    # Create the report file, and saves it.
    overview_data = skipped_well_controller(file_list, full_path, config)

    # Sleep for 10 seconds, to make sure that the report have been created before trying to send it.
    time.sleep(10)

    # Get elapse time for the transfers completion
    last_e_mail_time = float(window["-TIME_TEXT-"].get())
    first_e_mail_time = float(window["-INIT_TIME_TEXT-"].get())
    elapsed = last_e_mail_time - first_e_mail_time
    # Change it in to HMS (hour minute seconds) formate and store it
    elapsed_time = time.strftime("%Hh%Mm%Ss", time.gmtime(elapsed))
    overview_data["time_for_all_trans"] = elapsed_time

    # sends an E-mail, with the report included
    msg_subject = f"Final report for transfer: {date.today()}"
    e_mail_type = "final_report"
    mail_setup(msg_subject, overview_data, config, e_mail_type)
    print("sent final report")


def mail_setup(msg_subject, all_data, config, e_mail_type):
    """
    Sends an E-mail to user specified in the config file.
    :param msg_subject: error msg
    :type msg_subject: str
    :param all_data: All the data from the file.
    :type all_data: dict
    :param config: The configparser.
    :type config: configparser.ConfigParser
    :param e_mail_type: What kind of E-mail to send.
        "error" - sends an E-mail with all the fail transfers from the echo
        "final_report" - sends an E-mail with an overview of all the transfers, and a report for the complete transfer
    :type e_mail_type: str
    :return:
    """

    # Basic setup for sending.
    # The server - DTU internal server - Pulling from the config file
    # Sender - The E-mail that sends the msg - Pulling from the config file
    # Receivers -  List of people that will get the E-mail. - Pulling from the config file
    # File_data is for attachment
    file_data = None
    msg = EmailMessage()
    dtu_server = config["Email_settings"]["server"]
    sender = config["Email_settings"]["sender"]
    equipment_name = "Echo"
    receiver = []
    for people in config["Email_list"]:
        receiver.append(config["Email_list"][people])

    # Sends different E-mails depending on e-mail type.
    # Error E-mails, is sent if there is an error on Echo transfers
    # Final Report is sent when the full run is done. or if the system crash depending on
    if e_mail_type == "error":
        body = _mail_error(all_data, config)
    elif e_mail_type == "final_report":
        overview_data = all_data
        body = _mail_final_report(overview_data, config)
        filename = overview_data["path"]
        with open(filename, 'rb') as f:
            file_data = f.read()
        subtype = filename.split(".")[-1]
        filename = filename.split("/")[-1]

    # Setting up the e-mail
    msg["Subject"] = f"{msg_subject}"
    msg["from"] = f"{equipment_name} <{sender}>"
    msg["To"] = ", ".join(receiver)
    msg.set_content(body)
    if file_data:
        msg.add_attachment(file_data, maintype="application", subtype=subtype, filename=filename)
    # msg.attach(MIMEText(body))

    # Sending the E-mail.
    server = smtplib.SMTP(dtu_server, port=25)
    server.send_message(msg)
    # server.sendmail(msg["from"], msg["to"], msg.as_string())
    server.quit()


def listening_controller(config, run, window):
    """
    main controller for listening for files.
    :param config: The config handler, with all the default information in the config file.
    :type config: configparser.ConfigParser
    :param run: A state to tell if the listening is active or not
    :type run: bool
    :param window: The window where the activation of the listening is.
    :type window: PySimpleGUI.PySimpleGUI.Window
    :return:
    """

    path = config["Folder"]["in"]

    event_handler = MyEventHandler(window)

    observer = Observer()
    observer.schedule(event_handler, path, recursive=True)
    observer.start()

    try:
        while run:
            time.sleep(1)
            if window["-KILL-"].get():
                run = False

    finally:
        observer.stop()
        observer.join()
        print("done")


if __name__ == "__main__":

    msg_subject = "testing attachment"
    all_data = {}
    config = configparser.ConfigParser()
    config.read("config.ini")
    e_mail_type = "final_report"

    mail_setup(msg_subject, all_data, config, e_mail_type)