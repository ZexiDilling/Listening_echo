import os
import time


def folder_to_files(folder_path):
    """
    Gets all files in a folder in a list
    :param folder_path: the path to the folder
    :type folder_path: str
    :return: A list of all the files in the folder
    :rtype: list
    """
    file_list = []

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            file_list.append(str(os.path.join(root, file)))

    return file_list


def config_writer(config, heading, data_dict):
    """
    Writes data to the config file
    :param config: The config handler, with all the default information in the config file.
    :type config: configparser.ConfigParser
    :param heading: The heading of the specific configuration
    :type heading: str
    :param data_dict: The data that needs to be added to the dict
    :type data_dict: dict
    :return:
    """

    for data in data_dict:
        config.set(heading, data, data_dict[data])

    with open("config.ini", "w") as config_file:
        config.write(config_file)


def config_header_to_list(config, header):
    """
    Gets all the data from a category in the config file
    :param config: The config handler, with all the default information in the config file.
    :type config: configparser.ConfigParser
    :param header: name for witch category the data needs to be fetched
    :type header: str
    :return: all the data from a category in a list of list for display on a PySimpleGui-table
    :rtype: list
    """
    table_data = []
    for data in config[header]:
        temp_data = [data, config[header][data]]
        table_data.append(temp_data)

    return table_data


def progressbar(run, window):
    """
    The progress bar, that shows the program working
    :param run: If the bar needs to be running or not
    :type run: bool
    :param window: Where the bar is displayed
    :type window: PySimpleGUI.PySimpleGUI.Window
    :return:
    """
    min = 0
    max = 100
    counter = 0
    while run:
        if counter == min:
            runner = "pos"
        elif counter == max:
            runner = "neg"
        if runner == "pos":
            counter += 10
        elif runner == "neg":
            counter -= 10

        window["-BAR-"].update(counter)

        time.sleep(0.1)
        if window["-KILL-"].get():
            run = False

