import configparser

from controller import main

config = configparser.ConfigParser()
config.read("config.ini")
main(config)


