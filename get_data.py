from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl import Workbook, load_workbook
import csv
import xml.etree.ElementTree as ET

from helper_func import folder_to_files


def well_compound_list(folder):
    """
    Takes excel file with wells in clm 1 and compound name in clm 2
    :param file:
    :return:
    """
    file_list = folder_to_files(folder)

    compound_data = {}
    for files in file_list:
        plate_name = files.removesuffix(".xlsx").split("\\")[-1]
        # compound_data_org = {}
        compound_data[plate_name] = {}

        wb = load_workbook(filename=files)
        ws = wb.active
        for row, data in enumerate(ws):
            if row != 0:

                for col, cells in enumerate(data):

                    if col == 0:
                        temp_well = cells.value
                    if col == 1:
                        temp_compound = cells.value
                        try:
                            compound_data[plate_name][temp_well]
                        except KeyError:
                            compound_data[plate_name][temp_well] = temp_compound

    return compound_data


def get_all_trans_data(file_trans):
    """
    Takes excel data. all_transferees.
    :param file_trans:
    :return:
    """

    all_plate_trans = {}
    is_first_set = True
    first_set = None
    single_set = {}

    wb = load_workbook(filename=file_trans)
    ws = wb.active

    for row, data in enumerate(ws):
        if row != 0:

            for col, cells in enumerate(data):

                if col == 1:
                    temp_plate = cells.value
                    if row == 1:
                        first_set = temp_plate[:-2]

                    if temp_plate[:-2] != first_set:
                        is_first_set = False

                    try:
                        all_plate_trans[temp_plate[:-2]]
                    except KeyError:
                        all_plate_trans[temp_plate[:-2]] = {}

                    try:
                        all_plate_trans[temp_plate]
                    except KeyError:
                        all_plate_trans[temp_plate] = {}

                    if is_first_set:
                        try:
                            single_set[temp_plate[:-2]]
                        except KeyError:
                            single_set[temp_plate[:-2]] = {}

                # if col == 2:
                #     temp_well = cells.value

                if col == 3:
                    temp_compound = cells.value

                if col == 4:
                    temp_vol = float(cells.value)

                if col == 6:
                    temp_source = cells.value

                    try:
                        all_plate_trans[temp_plate[:-2]][temp_source]
                    except KeyError:
                        all_plate_trans[temp_plate[:-2]][temp_source] = {}

                    try:
                        all_plate_trans[temp_plate][temp_source]
                    except KeyError:
                        all_plate_trans[temp_plate][temp_source] = {}

                    if is_first_set:
                        try:
                            single_set[temp_plate[:-2]][temp_source]
                        except KeyError:
                            single_set[temp_plate[:-2]][temp_source] = {}

            try:
                all_plate_trans[temp_plate][temp_source][temp_compound]
            except KeyError:
                all_plate_trans[temp_plate][temp_source][temp_compound] = {}
            try:
                all_plate_trans[temp_plate][temp_source][temp_compound]["vol_needed"] += temp_vol
            except KeyError:
                all_plate_trans[temp_plate][temp_source][temp_compound]["vol_needed"] = 0

            try:
                all_plate_trans[temp_plate[:-2]][temp_source][temp_compound]
            except KeyError:
                all_plate_trans[temp_plate[:-2]][temp_source][temp_compound] = {}
            try:
                all_plate_trans[temp_plate[:-2]][temp_source][temp_compound]["vol_needed"] += temp_vol
            except KeyError:
                all_plate_trans[temp_plate[:-2]][temp_source][temp_compound]["vol_needed"] = 0

            if is_first_set:
                try:
                    single_set[temp_plate[:-2]][temp_source][temp_compound]
                except KeyError:
                    single_set[temp_plate[:-2]][temp_source][temp_compound] = {}
                try:
                    single_set[temp_plate[:-2]][temp_source][temp_compound]["vol_needed"] += temp_vol
                except KeyError:
                    single_set[temp_plate[:-2]][temp_source][temp_compound]["vol_needed"] = 0

    return all_plate_trans, single_set


def get_survey_csv_data(path):
    """
    get survey data from CSV files
    :param path:
    :return:
    """
    survey_data = {}

    file_list = folder_to_files(path)

    for file in file_list:
        plate_name = file.split("\\")[-1].split(".")[0]
        if file.endswith(".xml") and plate_name.startswith("Survey"):

            barcode = "_".join(plate_name.split("_")[1:])
            # print(plate_name)
            # print(barcode)
            try:
                survey_data[barcode]
            except KeyError:
                survey_data[barcode] = {}

            try:
                survey_data[barcode][plate_name]
            except KeyError:
                survey_data[barcode][plate_name] = {}

            with open(file, newline="\n") as csv_file:
                data = csv.reader(csv_file, delimiter=",")
                for row_index, row in enumerate(data):
                    if row_index > 3:
                        for index, data in enumerate(row):
                            if index == 0:
                                col_letter = data
                            else:
                                temp_well = f"{col_letter}{index}"
                                if data:
                                    survey_data[barcode][plate_name][temp_well] = data
        return survey_data


def get_xml_trans_data_skipping_wells(path):
    # random data
    skipped_wells = {}
    working_list = {}
    all_data = []
    skip_well_counter = 0

    # counting Transferees
    trans_plate_counter = []
    counter_plates = []

    # counting transferees all data
    all_trans_counter = {}

    file_list = folder_to_files(path)

    for files in file_list:

        if files.split("\\")[-1].startswith("Transfer"):
            # path = self.file_names(self.main_folder)
            doc = ET.parse(files)
            root = doc.getroot()

            # for counting plates and transferees
            for plates in root.iter("plate"):
                barcode = plates.get("barcode")
                source_destination = plates.get("type")

                if source_destination == "destination":
                    counter_plates.append(barcode)
                    temp_d_barcode = barcode
                if source_destination == "source":
                    temp_s_barcode = barcode

            try:
                all_trans_counter[temp_d_barcode].append(temp_s_barcode)
            except KeyError:
                all_trans_counter[temp_d_barcode] = [temp_s_barcode]

            # find amount of well that is skipped
            for wells in root.iter("skippedwells"):
                wells_skipped = wells.get("total")
                if int(wells_skipped) != 0:
                    all_data.append(wells_skipped)
                    skip_well_counter += int(wells_skipped)

                    # finds barcode for source and destination
                    for dates in root.iter("transfer"):
                        date = dates.get("date")
                        all_data.append(date)

                    # finds barcode for source and destination
                    for plates in root.iter("plate"):
                        barcode = plates.get("barcode")
                        source_destination = plates.get("type")
                        all_data.append(source_destination + ", " + barcode)

                        if source_destination == "source":
                            temp_barcode = barcode
                            try:
                                skipped_wells[barcode]
                            except KeyError:
                                skipped_wells[barcode] = {}
                        if source_destination == "destination":
                            temp_dest_barcode = barcode
                            try:
                                working_list[barcode]
                            except KeyError:
                                working_list[barcode] = {}

                            try:
                                working_list[barcode][temp_barcode]
                            except KeyError:
                                working_list[barcode][temp_barcode] = []

                    # finds destination and source wells data
                    for z in range(int(wells_skipped)):
                        temp_trans = []
                        dn = wells[z].get("dn")
                        n = wells[z].get("n")
                        reason = wells[z].get("reason")
                        vt = wells[z].get("vt")
                        all_data.append("SW: " + n + " DW: " + dn + " vol: " + vt)
                        all_data.append(" reason: " + reason)
                        temp_trans.append(n)
                        temp_trans.append(float(vt))
                        temp_trans.append(dn)
                        try:
                            skipped_wells[temp_barcode][n]["counter"] += 1
                            skipped_wells[temp_barcode][n]["vol"] += float(vt)
                        except KeyError:
                            skipped_wells[temp_barcode][n] = {"counter": 1, "vol": float(vt)}

                        working_list[temp_dest_barcode][temp_barcode].append(temp_trans)

    # counting plates
    # counts the number of repeated barcodes and makes a list with the barcode and amount of
    # instance with the same name

    for plates in counter_plates:
        number = counter_plates.count(plates)
        trans_plate_counter.append(str(plates) + "," + str(number))

    # remove duplicates from the list
    trans_plate_counter = list(dict.fromkeys(trans_plate_counter))

    return all_data, skipped_wells, skip_well_counter, working_list, trans_plate_counter, \
           all_trans_counter


def get_xml_trans_data_printing_wells(path):

    all_trans_data = {}

    file_list = folder_to_files(path)

    for files in file_list:
        if files.split("\\")[-1].startswith("Transfer"):
            trans_name = files.split("\\")[-1].removesuffix(".xml")
            # path = self.file_names(self.main_folder)
            doc = ET.parse(files)
            root = doc.getroot()

            # find amount of well that is skipped
            for wells in root.iter("printmap"):
                wells_printed = wells.get("total")
                if int(wells_printed) != 0:

                    all_trans_data[trans_name] = {"destination_plate": "", "source_plate": "", "transferees": {},
                                                  "date": ""}

                    for plates in root.iter("plate"):

                        barcode = plates.get("barcode")
                        source_destination = plates.get("type")

                        if source_destination == "source":
                            all_trans_data[trans_name]["source_plate"] = barcode


                        if source_destination == "destination":
                            all_trans_data[trans_name]["destination_plate"] = barcode

                    for dates in root.iter("transfer"):
                        all_trans_data[trans_name]["date"] = dates.get("date")

                    # finds destination and source wells data
                    for z in range(int(wells_printed)):
                        dest_well = wells[z].get("dn")
                        source_well = wells[z].get("n")
                        volume = wells[z].get("vt")
                        all_trans_data[trans_name]["transferees"] = {"destination_well": dest_well, "source_well":
                            source_well, "volume": volume}

    return all_trans_data
