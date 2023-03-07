import configparser
import time
import smtplib
from email.message import EmailMessage
from openpyxl import Workbook
from datetime import datetime, timedelta


def tester(config, error):

    procent_splitter = [1, 2, 5, 8]
    time_estimates_send = [1]
    current_plate = 2


    if current_plate in procent_splitter:
        if current_plate not in time_estimates_send:

            print("test")
    print(time_estimates_send)

if __name__ == "__main__":
    config = configparser.ConfigParser()
    config.read("config.ini")
    # print(type(config))
    error = "HJÆLP "
    tester(config, error)