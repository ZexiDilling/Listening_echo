import os
import time


def folder_to_files(folder_path):
    file_list = []

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            file_list.append(str(os.path.join(root, file)))

    return file_list


def config_writer(config, heading, data_dict):

    for data in data_dict:
        config.set(heading, data, data_dict[data])

    with open("config.ini", "w") as config_file:
        config.write(config_file)

def config_header_to_list(config, header):
    table_data = []
    for data in config[header]:
        temp_data = [data, config[header][data]]
        table_data.append(temp_data)

    return table_data


def progressbar(run, window):
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

