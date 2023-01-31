import configparser
import time
import smtplib
from email.message import EmailMessage
from openpyxl import Workbook


def tester(config, error):
    set_amount = 10
    starting_set = 0

    specific_transfers = {"13-plate-C": {"PP": False, "LDV": True}, "13-plate-D": {"PP": False, "LDV": True}, "13-plate-A": {"PP": True, "LDV": True}}

    if specific_transfers:
        plate_range = specific_transfers

    else:
        plate_range = range(set_amount + 1 - starting_set)

    for test in plate_range:
        if not plate_range[test]["PP"]:
            continue
        print(test)

if __name__ == "__main__":
    config = configparser.ConfigParser()
    config.read("config.ini")
    # print(type(config))
    error = "HJÆLP "
    tester(config, error)