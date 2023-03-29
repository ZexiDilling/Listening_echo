import configparser
import time
import smtplib
from email.message import EmailMessage
from openpyxl import Workbook
from datetime import datetime, timedelta


def tester(config, error):
    specific_transfers = None
    ending_set = 75
    starting_set = 70
    plate_letter = ["A", "B", "C", "D"]

    if specific_transfers:
        plate_range = specific_transfers
    else:
        plate_range = range(set_amount + 1 - starting_set)
        print(plate_range)

    for sets in plate_range:
        for letters in plate_letter:

            destination_plate = f"{sets + starting_set}-{letters}"
            print(destination_plate)

if __name__ == "__main__":
    config = configparser.ConfigParser()
    config.read("config.ini")
    # print(type(config))
    error = "HJÆLP "
    tester(config, error)