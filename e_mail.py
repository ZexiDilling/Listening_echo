import configparser
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
import smtplib


import xml.etree.ElementTree as ET


class MyEventHandler(FileSystemEventHandler):
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config.read("config.ini")


    def data_skipping_zero(self, path):
        skippped_wells = {}
        well_counter = {}
        well_vol = {}
        transferees = {}
        all_data = []
        counter = 0

        # path = self.file_names(self.main_folder)

        doc = ET.parse(path)
        root = doc.getroot()

        # find amount of well that is skipped
        for wells in root.iter("skippedwells"):
            wells_skipped = wells.get("total")
            if int(wells_skipped) != 0:
                all_data.append(wells_skipped)

                # finds barcode for source and destination
                for dates in root.iter("transfer"):
                    date = dates.get("date")
                    all_data.append(date)

                # finds barcode for source and destination
                for plates in root.iter("plate"):
                    barcode = plates.get("barcode")
                    source_destination = plates.get("type")
                    all_data.append(source_destination + ", " + barcode)

                    if source_destination == "source":
                        temp_barcode = barcode
                        try:
                            skippped_wells[barcode]
                        except KeyError:
                            skippped_wells[barcode] = []
                    if source_destination == "destination":
                        temp_dest_barcode = barcode
                        try:
                            transferees[barcode]
                        except KeyError:
                            transferees[barcode] = {}

                        try:
                            transferees[barcode][temp_barcode]
                        except KeyError:
                            transferees[barcode][temp_barcode] = []

                # finds destination and source wells data
                for z in range(int(wells_skipped)):
                    temp_trans = []
                    dn = wells[z].get("dn")
                    n = wells[z].get("n")
                    reason = wells[z].get("reason")
                    vt = wells[z].get("vt")
                    all_data.append("SW: " + n + " DW: " + dn + " vol: " + vt)
                    all_data.append(" reason: " + reason)
                    temp_trans.append(n)
                    temp_trans.append(dn)
                    temp_trans.append(float(vt))
                    try:
                        well_counter[n] += 1
                        well_vol[n] += float(vt)
                    except KeyError:
                        well_counter[n] = 1
                        well_vol[n] = float(vt)
                    if n not in skippped_wells[temp_barcode]:
                        skippped_wells[temp_barcode].append(n)
                    transferees[temp_dest_barcode][temp_barcode].append(temp_trans)


        return all_data, skippped_wells, well_counter, well_vol, transferees

    def on_created(self, event):

        time.sleep(1)
        all_data, skippped_wells, well_counter, well_vol, transferees = self.data_skipping_zero(event.src_path)
        if skippped_wells:
            mail_setup("JEG ER FEJLet", all_data, self.config)

    def on_deleted(self, event):
        print("delet")
        print(event)

    # def on_modified(self, event):
    #     print("mod")
    #     print(event)


def mail_setup(error, all_data, config):
    dtu_server = config["Email_settings"]["server"]
    sender = config["Email_settings"]["sender"]
    receiver = []
    for people in config["Email_list"]:
        receiver.append(config["Email_list"][people])

    missing_wells = int(all_data[0])
    trans_string = ""

    for counts in range(missing_wells):
        trans_string += f"Transferee: {all_data[4+(counts*2)]} - Error: {all_data[5+(counts*2)]}\n"

    subject = f"Echo Data Error - {error}"
    body = f"Missing {all_data[0]} Wells \n" \
           f"Source plate {all_data[2].split(',')[-1]}\n" \
           f"Destination plate: {all_data[3].split(',')[-1]}\n" \
           f"{trans_string}"


    message = """
    From: %s\r\nTo: %s\r\nSubject: %s\r\n\

    %s
    """ % (sender, ", ".join(receiver), subject, body)

    server = smtplib.SMTP(dtu_server, port=25)
    server.sendmail(sender, receiver, message)
    server.quit()


def listening_controller(config, run, window):

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