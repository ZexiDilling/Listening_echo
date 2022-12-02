from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl import Workbook


from get_data import well_compound_list, get_survey_csv_data, get_all_trans_data


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
    survey_folder = "C:/Users/phch/Desktop/more_data_files/surveys"
    path = "C:/Users/phch/Desktop/echo_data"
    save_file_name = "test_new_worklist"

    dead_vol_ul = {"LDV": 2.5, "PP": 15.0}

    set_amount = 1

    new_worklist(survey_folder, plate_layout_folder, all_trans_file, set_amount, dead_vol_ul, save_location, save_file_name)
    # excel_controller(trans_data_folder, plate_layout_folder, all_trans_file, data_location, file_name, save_location)
    # get_comments(all_trans_file)

    # print(file_names(path))

