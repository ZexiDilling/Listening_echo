import configparser
import time
import smtplib
from email.message import EmailMessage
from openpyxl import Workbook


def tester(config, error):
    test_list = []
    if test_list:
        print("Fuck")
    current_time = time.time()
    print(current_time)
    # time.sleep(2)
    time_2_hrs = time.time()

    elapsed = time_2_hrs-current_time

    check = time.strftime("%Hh%Mm%Ss", time.gmtime(elapsed))
    print(check)


    overview_data = {"plate_amount": 0,
                     "amount_complete_plates": 0,
                     "amount_failed_plates": 0,       #DONE
                     "failed_wells": 0,                   #DONE
                     "failed_trans": 0,                   #DONE
                     "amount_source_plates": 0,           #DONE
                     "time_for_all_trans": 0,                 #DONE
                     "path": ""}                              #DONE
    test = "hej"
    overview_data["path"] = test

if __name__ == "__main__":
    config = configparser.ConfigParser()
    config.read("config.ini")
    # print(type(config))
    error = "HJÆLP "
    tester(config, error)