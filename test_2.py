import configparser
import time
import smtplib
from email.message import EmailMessage
from openpyxl import Workbook
from datetime import datetime, timedelta


def tester(config, error):

    current_plate = 4
    procent_splitter = [5, 12, 25, 38]
    time_estimates_send = [5]
    current_plate = 5

    if current_plate in time_estimates_send and not time_estimates_send:

        print("test")
    print(time_estimates_send)

if __name__ == "__main__":
    config = configparser.ConfigParser()
    config.read("config.ini")
    # print(type(config))
    error = "HJÆLP "
    tester(config, error)