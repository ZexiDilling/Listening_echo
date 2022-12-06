import configparser
import smtplib
from email.message import EmailMessage
from openpyxl import Workbook


def tester(config, error):
    print(type(Workbook()))




if __name__ == "__main__":
    config = configparser.ConfigParser()
    config.read("config.ini")
    # print(type(config))
    error = "HJÆLP "
    tester(config, error)