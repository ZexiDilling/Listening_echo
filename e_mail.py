import configparser
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
import smtplib
from email.message import EmailMessage
import os

from get_data import get_xml_trans_data_skipping_wells


class MyEventHandler(FileSystemEventHandler):
    def __str__(self):
        """This is a standard class for watchdog.
        This is the class that is listening for files being created, moved or deleted.
        ATM the system only react to newly created files"""

    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config.read("config.ini")

    # @staticmethod
    # def data_skipping_zero(path):
    #     """
    #     Looking through a transferee XML-file from the ECHO. and looking to see if there are any skipped wells for the
    #     transferee. If there are skipped wells.
    #     Skipped wells, are wells that have not been transferred due to different reasons.
    #     :param path: A path to the file.
    #     :type path: str
    #     :returns: all_data, skipped_wells, well_counter, well_vol, transferees
    #     all_data - All the data from the missing transferees
    #     skipped_wells - What wells are skipped
    #     well_counter - Count how many times a single well is skipped
    #     well_vol - counts how much volume is missing per well
    #     transferees - Makes a dict over all transferees
    #     :rtype all_data: dict
    #     :rtype skipped_wells: dict
    #     :rtype well_counter: dict
    #     :rtype transferees: dict
    #     """
    #     skipped_wells = {}
    #     well_counter = {}
    #     well_vol = {}
    #     transferees = {}
    #     all_data = []
    #
    #     doc = ET.parse(path)
    #     root = doc.getroot()
    #
    #     # find amount of well that is skipped
    #     for wells in root.iter("skippedwells"):
    #         wells_skipped = wells.get("total")
    #         if int(wells_skipped) != 0:
    #             all_data.append(wells_skipped)
    #
    #             # finds barcode for source and destination
    #             for dates in root.iter("transfer"):
    #                 date = dates.get("date")
    #                 all_data.append(date)
    #
    #             # finds barcode for source and destination
    #             for plates in root.iter("plate"):
    #                 barcode = plates.get("barcode")
    #                 source_destination = plates.get("type")
    #                 all_data.append(source_destination + ", " + barcode)
    #
    #                 # Gets the barcode for the source plate
    #                 if source_destination == "source":
    #                     temp_barcode = barcode
    #                     try:
    #                         skipped_wells[barcode]
    #                     except KeyError:
    #                         skipped_wells[barcode] = []
    #
    #                 # Gets the barcode for the destination plate
    #                 if source_destination == "destination":
    #                     temp_dest_barcode = barcode
    #                     try:
    #                         transferees[barcode]
    #                     except KeyError:
    #                         transferees[barcode] = {}
    #
    #                     try:
    #                         transferees[barcode][temp_barcode]
    #                     except KeyError:
    #                         transferees[barcode][temp_barcode] = []
    #
    #             # finds destination and source wells data
    #             for z in range(int(wells_skipped)):
    #                 temp_trans = []
    #                 destination_well = wells[z].get("dn")
    #                 source_well = wells[z].get("n")
    #                 reason = wells[z].get("reason")
    #                 trans_volume = wells[z].get("vt")
    #                 all_data.append("SW: " + source_well + " DW: " + destination_well + " vol: " + trans_volume)
    #                 all_data.append(" reason: " + reason)
    #                 temp_trans.append(source_well)
    #                 temp_trans.append(destination_well)
    #                 temp_trans.append(float(trans_volume))
    #                 try:
    #                     well_counter[source_well] += 1
    #                     well_vol[source_well] += float(trans_volume)
    #                 except KeyError:
    #                     well_counter[source_well] = 1
    #                     well_vol[source_well] = float(trans_volume)
    #                 if source_well not in skipped_wells[temp_barcode]:
    #                     skipped_wells[temp_barcode].append(source_well)
    #                 transferees[temp_dest_barcode][temp_barcode].append(temp_trans)
    #
    #     return all_data, skipped_wells, well_counter, well_vol, transferees

    def on_created(self, event):
        """
        This event is triggered when a new file appears in the target folder
        It checks the file in the event for missing transferees, if there are any, it sends an E-mail.
        :param event: The full event, including the path to the file that have been created
        """
        print(type(event))
        print(event)
        # checks if path is a directory

        if os.path.isfile(event.src_path):
            counter = 0
            sending_mail = True
            while sending_mail:
                # Set timer to sleep while echo is finishing writing data to the files
                time.sleep(2)
                try:
                    all_data, skipped_wells, skip_well_counter, _, _, _ = get_xml_trans_data_skipping_wells(event.src_path)
                except:
                    counter += 1
                    if counter == 30:
                        sending_mail = False
                else:
                    if skipped_wells:
                        _mail_setup(f"Liquid Transferee failed for {skip_well_counter} wells", all_data, self.config)
                    sending_mail = False
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


def _mail_setup(error_msg, all_data, config):
    """
    Sends an E-mail to user specified in the config file.
    :param error_msg: error msg
    :type error_msg: str
    :param all_data: All the data from the file.
    :type all_data: dict
    :param config: The configparser.
    :type config: configparser.ConfigParser
    :return:
    """

    # Basic setup for sending.
    # The server - DTU internal server
    # Sender - The E-mail that sends the msg
    # Receivers -  List of people that will get the E-mail.
    dtu_server = config["Email_settings"]["server"]
    sender = config["Email_settings"]["sender"]
    equipment_name = "Echo"
    receiver = []
    for people in config["Email_list"]:
        receiver.append(config["Email_list"][people])

    # amount of missing wells
    missing_wells = int(all_data[0])
    trans_string = ""

    # going through ever single missing well, and writes the error msg.
    for counts in range(missing_wells):
        transferee = all_data[4+(counts*2)]
        error = all_data[5+(counts*2)]
        error_code = error[0:9]
        trans_string += f"Transferee: {transferee} - Error: {error}"

        # Looking for known error messages
        try:
            config["Echo_error"][error_code]
        except KeyError:
            trans_string += " - New error YAY!!!\n"
        else:
            trans_string += f" - {config['Echo_error'][error_code]}\n"

    # combine all details into one body of text
    body = f"Missing {all_data[0]} Wells \n" \
           f"Source plate {all_data[2].split(',')[-1]}\n" \
           f"Destination plate: {all_data[3].split(',')[-1]}\n" \
           f"{trans_string}"

    # Setting up the e-mail
    msg = EmailMessage()
    msg["Subject"] = f"{error_msg}"
    msg["from"] = f"{equipment_name} <{sender}>"
    msg["To"] = ", ".join(receiver)
    msg.set_content(body)

    # Sending the E-mail.
    server = smtplib.SMTP(dtu_server, port=25)
    server.send_message(msg)
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

    event_handler = MyEventHandler()

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

    mail_setup("hey", "plate_2", "hej")