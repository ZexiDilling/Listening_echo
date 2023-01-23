
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


def clear_file(file, config):
    """
    clears the file for data
    :param file: the files that needs to be cleared. reffed in config
    :type file: str
    :param config:
    :return:
    """
    with open(config["Temp_files"][file], "w") as f:
        f.close()


def write_temp_list_file(temp_file_name, file, config):
    """
    Writes file names to a list
    :param file: the file that needs to be written
    :type file: str
    :param config: the config file
    :return:
    """
    trans_list_file = config["Temp_files"][temp_file_name]

    with open(trans_list_file, "a") as f:
        f.write(file)
        f.write(",")


def read_temp_list_file(temp_file_name, config):
    """
    Reads the txt-file with all the trans files in it.
    :param config:
    :return: a list of all the trans files
    :rtype: list
    """

    with open(config["Temp_files"][temp_file_name], "r") as f:
        lines = f.read()
        lines = lines.removesuffix(",")
        file_list = lines.split(",")

    return file_list

if __name__ == "__main__":
    import configparser
    config = configparser.ConfigParser()
    config.read("config.ini")
    file = "trans_list"
    write_temp_list_file(file, config)
    read_temp_list_file(config)

    # sg.main_get_debug_data()