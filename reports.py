from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl import Workbook

from get_data import *


def _write_to_excel_plate_transferees_list_of_plates(wb, data):

    ws = wb.create_sheet('Plate_trans_list')

    # write headers
    ws.cell(row=1, column=1, value="destination plate").font = Font(bold=True)
    ws.merge_cells("B1:F1")
    ws["B1"].value = "source plates"
    ws["B1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["B1"].font = Font(bold=True)

    col = 1
    row = 2

    for destination in data:
        ws.cell(row=row, column=col, value=destination)
        for index, source in enumerate(data[destination]):
            ws.cell(row=row, column=col + 1 + index, value=source)

        row += 1


def _write_to_excel_plate_transferees(wb, data):

    ws = wb.create_sheet('Plate_trans_counter')

    # split the data up by ","
    temp = []
    for i in enumerate(data):
        temp.append(i[1].split(","))

    # write headers
    ws.cell(row=1, column=1, value="destination plate").font = Font(bold=True)
    ws.cell(row=1, column=2, value="counter").font = Font(bold=True)

    # write in the data, barcodes for destination plates and the number of transferees.
    col = 1
    row = 2
    for plate, counts in temp:
        ws.cell(row=row, column=col, value=plate)
        ws.cell(row=row, column=col + 1, value=counts)
        row += 1


def _write_to_excel_error_report(wb, data, type):
    # Error_Report
    # all data
    ws = wb.create_sheet(f"{type}")
    row = 1
    col = 1
    i = 0
    while i < len(data):
        # write to excel worksheet
        ws.cell(row=row, column=col, value=data[i+1]).font = Font(bold=True)    # date
        ws.cell(row=row, column=col + 1, value=data[i + 2]).font = Font(bold=True)  # source barcode
        ws.cell(row=row, column=col + 2, value=data[i + 3]).font = Font(bold=True)  # destination barcode
        # sets n = to wells skipped
        n = data[i]
        ws.cell(row=row, column=col + 3, value=n).font = Font(bold=True)    # writes amount of wells skipped
        # sets counter for next loop
        temp = i + 4
        i = temp
        q = 0

        # loop for writing out wells
        for k in range(temp, (int(n)*2)+temp, 2):
            # makes sure that the Excel file is only 5-row wide
            if q == 5:
                row = row + 2
                q = 0
            # writes SW, DW and vol for col 1, writes reason for skipped in col 2
            ws.cell(row=row + 1, column=col + q, value=data[k])
            ws.cell(row=row + 2, column=col + q, value=data[k + 1])
            q += 1

        # counter sets for next iteration
        row = row + 3
        col = col
        i = int(n)*2+i


def _write_work_list(wb, data, type):
    ws = wb.create_sheet(f'{type}_worklist')

    row = 1
    col = 1

    # write headlines
    ws.cell(row=row, column=col, value="source_plate")
    ws.cell(row=row, column=col + 1, value="source_well")
    ws.cell(row=row, column=col + 2, value="trans_vol")
    ws.cell(row=row, column=col + 3, value="destination_well")
    ws.cell(row=row, column=col + 4, value="destination_plate")
    row += 1

    for destination in data:
        for source in data[destination]:
            for index, trans in enumerate(data[destination][source]):
                # source plate
                ws.cell(row=row, column=col, value=source)

                for off_set, info in enumerate(data[destination][source][index]):
                    # source well trans vol, destination will
                    ws.cell(row=row, column=col + off_set + 1, value=str(info))
                # destination plate
                ws.cell(row=row, column=col + 4, value=destination)
                row += 1


def _skip_report(wb, skipped_wells, skip_well_counter, working_list):
    ws = wb.active
    ws.title = "Overview_Report"
    row = 1
    col = 1

    source_plate_amount = 0
    destination_plate_amount = 0
    source_plate_list = []
    destination_plate_list = []
    for destination_plates in working_list:
        destination_plate_amount += 1
        destination_plate_list.append(destination_plates)
        for source_plate in working_list[destination_plates]:

            if source_plate not in source_plate_list:
                source_plate_list.append(source_plate)
                source_plate_amount += 1

    # overview
    ws.cell(row=row, column=col, value="Skipped Wells counter").font = Font(bold=True)
    ws.cell(row=row, column=col + 1, value=skip_well_counter)
    row += 1
    ws.cell(row=row, column=col, value="Source Plates counter").font = Font(bold=True)
    ws.cell(row=row, column=col + 1, value=source_plate_amount)
    row += 1
    ws.cell(row=row, column=col, value="Destination plate counter").font = Font(bold=True)
    ws.cell(row=row, column=col + 1, value=destination_plate_amount)
    row += 2
    ws.cell(row=row, column=col, value="Destination Plates").font = Font(bold=True)
    ws.cell(row=row, column=col + 1, value="Source Plates").font = Font(bold=True)
    row += 1

    temp_row = row
    for destination in destination_plate_list:
        ws.cell(row=temp_row, column=col, value=destination)
        temp_row += 1

    temp_row = row
    for source in source_plate_list:
        ws.cell(row=temp_row, column=col + 1, value=source)
        temp_row += 1

    temp_row = row
    temp_col = col
    for source_plate in skipped_wells:
        ws.cell(row=temp_row - 1, column=temp_col + 3, value="Source Plates").font = Font(bold=True)
        ws.cell(row=temp_row - 1, column=temp_col + 4, value="Source Wells").font = Font(bold=True)
        ws.cell(row=temp_row - 1, column=temp_col + 5, value="Well Instance").font = Font(bold=True)
        ws.cell(row=temp_row - 1, column=temp_col + 6, value="Well Volume").font = Font(bold=True)

        ws.cell(row=temp_row, column=temp_col + 3, value=source_plate)
        for index, source_well in enumerate(skipped_wells[source_plate]):
            ws.cell(row=temp_row + index, column=temp_col + 4, value=source_well)
            ws.cell(row=temp_row + index, column=temp_col + 5, value=skipped_wells[source_plate][source_well]["counter"])
            ws.cell(row=temp_row + index, column=temp_col + 6, value=skipped_wells[source_plate][source_well]["vol"])

        temp_col += 4


def skipped_well_controller(data_location, file_name, save_location):
    wb = Workbook()
    all_data, skipped_wells, skip_well_counter, working_list, trans_plate_counter, all_trans_counter = \
        get_xml_trans_data_skipping_wells(data_location)
    _skip_report(wb, skipped_wells, skip_well_counter, working_list)
    _write_to_excel_plate_transferees_list_of_plates(wb, all_trans_counter)
    _write_to_excel_plate_transferees(wb, trans_plate_counter)
    _write_to_excel_error_report(wb, all_data, "Error_Report")
    _write_work_list(wb, working_list, "Old")

    wb.save(f"{save_location}/{file_name}.xlsx")
    return "Done"


def _rename_source_plates(trans_data, prefix_dict):


    for trans in trans_data:
        temp_source_plate = trans_data[trans]["source_plate"]
        temp_date = trans_data[trans]["date"]
        for prefix in prefix_dict:
            if prefix_dict[prefix]["start"] <= temp_date <= prefix_dict[prefix]["end"]:
                trans_data[trans]["source_plate"] = f"{prefix}_{temp_source_plate}"


def _write_trans_report(trans_data, compound_data, save_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    row = 1
    col = 1

    # Headers
    ws.cell(row=row, column=col + 0, value="Destination Plates").font = Font(bold=True)
    ws.cell(row=row, column=col + 1, value="Destination Well").font = Font(bold=True)
    ws.cell(row=row, column=col + 2, value="Compound").font = Font(bold=True)
    ws.cell(row=row, column=col + 3, value="Volume").font = Font(bold=True)
    ws.cell(row=row, column=col + 4, value="Comments").font = Font(bold=True)
    ws.cell(row=row, column=col + 5, value="Source Plates").font = Font(bold=True)
    ws.cell(row=row, column=col + 6, value="Source Well").font = Font(bold=True)

    row += 1

    for row_index, trans in enumerate(trans_data):
        destination_plate = trans_data[trans]["destination_plate"]
        source_plate = trans_data[trans]["source_plate"]
        volume = trans_data[trans]["transferees"]["volume"]
        destination_well = trans_data[trans]["transferees"]["destination_well"]
        source_well = trans_data[trans]["transferees"]["source_well"]
        compound = compound_data[source_plate][source_well]

        ws.cell(row=row + row_index, column=col + 0, value=destination_plate)   #.font = Font(bold=True)
        ws.cell(row=row + row_index, column=col + 1, value=destination_well)    #.font = Font(bold=True)
        ws.cell(row=row + row_index, column=col + 2, value=compound)    #.font = Font(bold=True)
        ws.cell(row=row + row_index, column=col + 3, value=volume)  #.font = Font(bold=True)
        ws.cell(row=row + row_index, column=col + 4, value="Comments")  #.font = Font(bold=True)
        ws.cell(row=row + row_index, column=col + 5, value=source_plate)    #.font = Font(bold=True)
        ws.cell(row=row + row_index, column=col + 6, value=source_well) #.font = Font(bold=True)

    wb.save(save_file)


def trans_report_controller(trans_data_folder, plate_layout_folder, all_trans_file, data_location, file_name, save_location):
    trans_data = get_xml_trans_data_printing_wells(trans_data_folder)
    compound_data = well_compound_list(plate_layout_folder)
    save_file = f"{save_location}/{file_name}.xlsx"

    prefix_on = True
    prefix_dict = {"OLD": {"start": "2022-11-22", "end": "2022-12-01"},
              "NEW": {"start": "2022-10-01", "end": "2022-11-21"}}
    if prefix_on:
        rename_source_plates(trans_data, prefix_dict)

    # get_comments(all_trans_file) # TODO find out what to do about comments

    write_report(trans_data, compound_data, save_file)


    return "Done"


def well_report(trans_file, save_file):

    wb = load_workbook(trans_file)
    ws = wb.active

    all_wells = {}

    for row_index, row in enumerate(ws):
        if row_index > 0:
            for clm, data in enumerate(row):
                if clm == 0:
                    comment = data.value
                elif clm == 1:
                    destination_plate = data.value
                elif clm == 2:
                    destination_well = data.value
                elif clm == 3:
                    compound = data.value
                elif clm == 4:
                    volume = data.value
                elif clm == 5:
                    source_well = data.value
                elif clm == 6:
                    source_plate = data.value
                else:
                    source_plate_type = data.value

                    try:
                        all_wells[source_plate]
                    except KeyError:
                        all_wells[source_plate] = {}

                    try:
                        all_wells[source_plate][source_well]
                    except KeyError:
                        all_wells[source_plate][source_well] = {"counter": 1, "volume": float(volume),
                                                                "compound": compound}
                    else:
                        all_wells[source_plate][source_well]["counter"] += 1
                        all_wells[source_plate][source_well]["volume"] += float(volume)

    wb = Workbook()
    ws = wb.active
    ws.title = "Well Report"
    row = 1
    col = 1

    # Headers
    ws.cell(row=row, column=col + 0, value="Source Plates").font = Font(bold=True)
    ws.cell(row=row, column=col + 1, value="Source Well").font = Font(bold=True)
    ws.cell(row=row, column=col + 2, value="count").font = Font(bold=True)
    ws.cell(row=row, column=col + 3, value="volume").font = Font(bold=True)
    ws.cell(row=row, column=col + 4, value="compound").font = Font(bold=True)

    row += 1

    for plates in all_wells:
        for wells in all_wells[plates]:
            ws.cell(row=row, column=col + 0, value=plates)
            ws.cell(row=row, column=col + 1, value=wells)
            ws.cell(row=row, column=col + 2, value=all_wells[plates][wells]["counter"])
            ws.cell(row=row, column=col + 3, value=all_wells[plates][wells]["volume"])
            ws.cell(row=row, column=col + 4, value=all_wells[plates][wells]["compound"])
            row += 1

    wb.save(save_file)
    print("done")


def _compound_to_survey(plate_layout, survey_data):

    survey_layout = {}

    for plates in survey_data:

        try:
            survey_layout[plates]
        except KeyError:
            survey_layout[plates] = {}

        for plate_names in survey_data[plates]:

            for well in survey_data[plates][plate_names]:

                if survey_data[plates][plate_names][well] != 0:

                    try:
                        temp_compound = plate_layout[plate_names][well]
                    except KeyError:
                        temp_compound = "No compound match found"
                    temp_vol = survey_data[plates][plate_names][well]

                    try:
                        survey_layout[plates][temp_compound]
                    except KeyError:
                        survey_layout[plates][temp_compound] = {}

                    try:
                        survey_layout[plates][temp_compound][plate_names]
                    except KeyError:
                        survey_layout[plates][temp_compound][plate_names] = {}

                    try:
                        survey_layout[plates][temp_compound][plate_names][well]
                    except KeyError:
                        survey_layout[plates][temp_compound][plate_names][well] = ""

                    survey_layout[plates][temp_compound][plate_names][well] = float(temp_vol)

    return survey_layout


def _write_new_worklist(set_data, survey_layout, dead_vol_ul, set_amount, save_file):
    wb = Workbook()
    ws = wb.active
    breaking = False
    row = 1
    col = 1

    error_missing_survey_data = "Missing survey Data"
    error_missing_liquid = "Not enough liquid"

    # headers:
    ws.cell(row=row, column=col + 0, value="source_plates").font = Font(bold=True)
    ws.cell(row=row, column=col + 1, value="source_well").font = Font(bold=True)
    ws.cell(row=row, column=col + 2, value="volume").font = Font(bold=True)
    ws.cell(row=row, column=col + 3, value="destination_well").font = Font(bold=True)
    ws.cell(row=row, column=col + 4, value="destination_plates").font = Font(bold=True)
    ws.cell(row=row, column=col + 5, value="compound").font = Font(bold=True)
    ws.cell(row=row, column=col + 6, value="comments").font = Font(bold=True)

    row += 1

    set_number = 1
    for sets in range(set_amount):
        for rows in set_data:
            destination_plate = f"{set_number + sets}-{set_data[rows]['destination_plate']}"
            destination_well = set_data[rows]["destination_well"]
            compound = set_data[rows]["compound"]
            volume_nl = float(set_data[rows]["volume_nl"])
            volume_ul = volume_nl/1000
            sample_comment = set_data[rows]["sample_comment"]
            source_plate_origin = set_data[rows]["source_plate"]
            plate_type = set_data[rows]["plate_type"]
            try:
                survey_layout[source_plate_origin][compound]
            except KeyError:
                source_plate = error_missing_survey_data
                source_well = error_missing_survey_data
            else:
                for temp_plates in survey_layout[source_plate_origin][compound]:
                    for wells in survey_layout[source_plate_origin][compound][temp_plates]:
                        if survey_layout[source_plate_origin][compound][temp_plates][wells] >= volume_ul + dead_vol_ul[plate_type]:
                            source_plate = temp_plates
                            source_well = wells
                            survey_layout[source_plate_origin][compound][temp_plates][wells] -= volume_ul
                            breaking = True
                            break
                        else:
                            source_plate = error_missing_liquid
                            source_well = error_missing_liquid

                    if breaking:
                        breaking = False
                        break
            if source_plate == error_missing_survey_data or source_plate == error_missing_liquid:
                ws.cell(row=row, column=col + 0, value=source_plate).fill = PatternFill(start_color='B284BE',
                                                                                        end_color='B284BE',
                                                                                        fill_type='solid')
                ws.cell(row=row, column=col + 1, value=source_well).fill = PatternFill(start_color='B284BE',
                                                                                       end_color='B284BE',
                                                                                       fill_type='solid')
                ws.cell(row=1, column=1).fill = PatternFill(start_color='B284BE', end_color='B284BE',
                                                                         fill_type='solid')
                ws.cell(row=1, column=2).fill = PatternFill(start_color='B284BE', end_color='B284BE',
                                                            fill_type='solid')

            else:
                ws.cell(row=row, column=col + 0, value=source_plate)
                ws.cell(row=row, column=col + 1, value=source_well)
            ws.cell(row=row, column=col + 2, value=volume_nl)
            ws.cell(row=row, column=col + 3, value=destination_well)
            ws.cell(row=row, column=col + 4, value=destination_plate)
            ws.cell(row=row, column=col + 5, value=compound)
            ws.cell(row=row, column=col + 6, value=sample_comment)
            row += 1

    wb.save(save_file)


def new_worklist(survey_folder, plate_layout_folder, file_trans, set_amount, dead_vol_ul, save_location, save_file_name):
    save_file = f"{save_location}/{save_file_name}.xlsx"
    survey_data = get_survey_csv_data(survey_folder)
    plate_layout = well_compound_list(plate_layout_folder)
    _, _, set_compound_data = get_all_trans_data(file_trans)

    survey_layout = _compound_to_survey(plate_layout, survey_data)

    _write_new_worklist(set_compound_data, survey_layout, dead_vol_ul, set_amount, save_file)

    print("done")

if __name__ == "__main__":
    trans_data_folder = "C:/Users/phch/Desktop/more_data_files/2022-11-22"
    plate_layout_folder = "C:/Users/phch/Desktop/more_data_files/plate_layout"
    data_location = "C:/Users/phch/Desktop/more_data_files/2022-11-22"
    file_name = "test_trans_report"
    save_location = "C:/Users/phch/Desktop/more_data_files/"
    all_trans_file = "C:/Users/phch/Desktop/more_data_files/all_trans.xlsx"
    path = "C:/Users/phch/Desktop/echo_data"

    trans_report_controller(trans_data_folder, plate_layout_folder, all_trans_file, data_location, file_name, save_location)
    # well_report(all_trans_file)

    # print(file_names(path))

