from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl import Workbook, load_workbook

from get_data import get_xml_trans_data_skipping_wells, get_xml_trans_data_printing_wells, well_compound_list,\
    get_survey_csv_data, get_all_trans_data


def _write_to_excel_plate_transferees_list_of_plates(wb, data, title):
    """
    Writes witch source plates goes to witch destination plate
    :param wb: excel Workbook
    :type wb: openpyxl.workbook.workbook.Workbook
    :param data: Data with all plate transferees
    :type data: dict
    :return:
    """

    # Create the sheets and name it
    ws = wb.create_sheet(title)

    # write headers
    ws.cell(row=1, column=1, value="destination plate").font = Font(bold=True)
    ws.merge_cells("B1:F1")
    ws["B1"].value = "source plates"
    ws["B1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["B1"].font = Font(bold=True)

    col = 1
    row = 2

    #writes the data
    for destination in data:
        ws.cell(row=row, column=col, value=destination)
        for index, source in enumerate(data[destination]):
            ws.cell(row=row, column=col + 1 + index, value=source)

        row += 1


def _write_to_excel_plate_transferees(wb, data, title):    # TODO description
    """

    :param wb: excel Workbook
    :type wb: openpyxl.workbook.workbook.Workbook
    :param data:
    :return:
    """
    ws = wb.create_sheet(title)

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


def _write_to_excel_error_report(wb, data, title): # TODO description
    """

    :param wb: excel Workbook
    :type wb: openpyxl.workbook.workbook.Workbook
    :param data:
    :param type:
    :return:
    """
    # Error_Report
    # all data
    ws = wb.create_sheet(title)
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


def _write_work_list(wb, data, title): # TODO description
    """

    :param wb: excel Workbook
    :type wb: openpyxl.workbook.workbook.Workbook
    :param data:
    :param type:
    :return:
    """
    ws = wb.create_sheet(title)

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


def _skip_report(wb, skipped_wells, skip_well_counter, working_list, overview_data, title): # TODO description
    """

    :param wb: excel Workbook
    :type wb: openpyxl.workbook.workbook.Workbook
    :param skipped_wells:
    :param skip_well_counter:
    :param working_list:
    :return:
    """
    ws = wb.active
    ws.title = title
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
    failed_source_wells = 0
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
            failed_source_wells += 1
        temp_col += 4

    overview_data["amount_failed_plates"] = destination_plate_amount
    overview_data["failed_trans"] = skip_well_counter
    overview_data["amount_source_plates"] = source_plate_amount
    overview_data["failed_wells"] = failed_source_wells


def _write_completed_plates(wb, zero_error_trans_plate, overview_data, title):
    ws = wb.create_sheet(title)
    row = 1
    col = 1

    # Headlines:
    ws.cell(row=row, column=col, value="Complete plates").font = Font(bold=True)

    row += 1
    complete_plate_counter = 0
    if zero_error_trans_plate:
        for plate_index, plates in enumerate(zero_error_trans_plate):
            ws.cell(row=row + plate_index, column=col, value=plates)
            complete_plate_counter += 1
    else:
        ws.cell(row=row, column=col, value="All plates failed somewhere").font = Font(bold=True)

    overview_data["amount_complete_plates"] = complete_plate_counter


def skipped_well_controller(data_location, full_path): # TODO description
    """

    :param data_location:
    :param report_name:
    :param save_location:
    :return:
    """
    # Getting data for wells being skipped
    all_data, skipped_wells, skip_well_counter, working_list, trans_plate_counter, all_trans_counter, \
    zero_error_trans_plate = get_xml_trans_data_skipping_wells(data_location)
    # Gets data from all transfers
    trans_data = get_xml_trans_data_printing_wells(data_location)

    # Create a dict with data to get a quick overview
    overview_data = {"plate_amount": 0,
                     "amount_complete_plates": 0,
                     "amount_failed_plates": 0,
                     "failed_wells": 0,
                     "failed_trans": 0,
                     "amount_source_plates": 0,
                     "time_for_all_trans": 0,
                     "path": ""}

    # Create the workbook
    wb = Workbook()

    # Add different reports to their own worksheet in an excel ark, to create a full report over transferes and skipped wells
    _skip_report(wb, skipped_wells, skip_well_counter, working_list, overview_data, title="Overview_Report")
    _write_to_excel_plate_transferees_list_of_plates(wb, all_trans_counter, title="Plate_trans_list")
    _write_to_excel_plate_transferees(wb, trans_plate_counter, title="Plate_trans_counter")
    _write_to_excel_error_report(wb, all_data, title="Error_Report")
    _write_work_list(wb, working_list, title="Old_Worklist")
    _write_completed_plates(wb, zero_error_trans_plate, overview_data, title="Completed_plates")
    _write_trans_report(wb, trans_data, title="Trans_Report", compound_data=None, report=True)

    wb.save(full_path)

    overview_data["plate_amount"] = overview_data["amount_complete_plates"] + overview_data["amount_failed_plates"]

    return overview_data


def _rename_source_plates(trans_data, prefix_dict): # TODO description


    for trans in trans_data:
        temp_source_plate = trans_data[trans]["source_plate"]
        temp_date = trans_data[trans]["date"]
        for prefix in prefix_dict:
            if prefix_dict[prefix]["start"] <= temp_date <= prefix_dict[prefix]["end"]:
                trans_data[trans]["source_plate"] = f"{prefix}_{temp_source_plate}"


def _write_trans_report(wb, trans_data, compound_data, title, report): # TODO description

    # sets row and coloumn counter for headlines
    row = 1
    col = 1

    #Check what the function is being used for. If it is for the analysis, then there is no reason to add compound data
    #It needs to create a new worksheet if it is for the report.
    #Spacer is to make space for the 'compound' data if it is included
    if report:
        ws = wb.create_sheet(title)
        spacer = 0
    else:
        ws = wb.active
        ws.title = title
        spacer = 1
        ws.cell(row=row, column=col + 2, value="Compound").font = Font(bold=True)

    # Headers
    ws.cell(row=row, column=col + 0, value="Destination Plates").font = Font(bold=True)
    ws.cell(row=row, column=col + 1, value="Destination Well").font = Font(bold=True)
    ws.cell(row=row, column=col + spacer + 2, value="Volume").font = Font(bold=True)
    ws.cell(row=row, column=col + spacer + 3, value="Source Plates").font = Font(bold=True)
    ws.cell(row=row, column=col + spacer + 4, value="Source Well").font = Font(bold=True)

    row += 1

    for row_index, trans in enumerate(trans_data):
        destination_plate = trans_data[trans]["destination_plate"]
        source_plate = trans_data[trans]["source_plate"]
        volume = trans_data[trans]["transferees"]["volume"]
        destination_well = trans_data[trans]["transferees"]["destination_well"]
        source_well = trans_data[trans]["transferees"]["source_well"]
        if compound_data:
            compound = compound_data[source_plate][source_well]

        ws.cell(row=row + row_index, column=col + 0, value=destination_plate)   #.font = Font(bold=True)
        ws.cell(row=row + row_index, column=col + 1, value=destination_well)    #.font = Font(bold=True)
        ws.cell(row=row + row_index, column=col + spacer + 2, value=volume)  #.font = Font(bold=True)
        ws.cell(row=row + row_index, column=col + spacer + 3, value=source_plate)    #.font = Font(bold=True)
        ws.cell(row=row + row_index, column=col + spacer + 4, value=source_well) #.font = Font(bold=True)
        if not report and compound_data:
            ws.cell(row=row + row_index, column=col + 2, value=compound)  # .font = Font(bold=True)


def trans_report_controller(trans_data_folder, plate_layout_folder, file_name, save_location): # TODO description
    trans_data = get_xml_trans_data_printing_wells(trans_data_folder)
    compound_data = well_compound_list(plate_layout_folder)
    save_file = f"{save_location}/{file_name}.xlsx"

    prefix_on = True
    prefix_dict = {"OLD": {"start": "2022-11-22", "end": "2022-12-01"},
              "NEW": {"start": "2022-10-01", "end": "2022-11-21"}}
    if prefix_on:
        _rename_source_plates(trans_data, prefix_dict)

    wb = Workbook()

    _write_trans_report(wb, trans_data, compound_data, title="Trans_report", report=False)

    wb.save(save_file)

    return "Done"


def well_report(trans_file, save_file): # TODO description

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


def _compound_to_survey(plate_layout, survey_data): # TODO description

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


def _write_new_worklist(set_compound_data, survey_layout, dead_vol_ul, set_amount, save_file, starting_set): # TODO description
    """
    writes two worklist for plate_printing.
    1 for LDV trans
    1 for PP trans
    and sets them in one excel sheet
    :param set_compound_data: Data over compounds
    :type set_compound_data: dict
    :param survey_layout: combination of survey data for a plate, and the layout of the plate. To know how much is left
        of each compound
    :type survey_layout: dict
    :param dead_vol_ul: dead volume for plates #TODO Get this data from the config file
    :type dead_vol_ul: dict
    :param set_amount: Amount of sets to produce. This do not take into account the starting set. so set it to 50 and
        starting set at 40, and it will only produce 10 sets.
    :type set_amount: int
    :param save_file: Where to save the data and what name to call it
    :type save_file: str
    :param starting_set: The number of the first set. To make it possible to start from a different number than 1
    :type starting_set: int
    :return: an excel file with data

    """
    wb = Workbook()
    # Create a sheet for each transferee, as we can't run LDV and PP trans at the same time
    ws0 = wb.active
    ws1 = wb.create_sheet("LDV_trans")
    ws2 = wb.create_sheet("PP_trans")

    # Failed plates, are plates not enough liquid for a transferee. name is depending on plate name and compound
    failed_plate = {}

    # Compound dict is a dict of compounds for specific plates, where dead volume is set lower, due to missing liquids.
    # An Echo will mostly work under "dead volume".
    compound_list = {"P23_LDV": {"Buparlisib": 2.0, "BGB-11417": 2.2}, "P23_LDV2_I2": {"Copanlisib": 2.0}}
    # compound_list = {}

    # Breaking is used, to make sure that all wells are used. instead of all transfers being using the last well in
    # the list
    breaking = False

    # Setup rows for the different worksheets
    row_ldv = 1
    row_pp = 1
    row = 1
    col = 1

    # setup error msg.
    error_missing_survey_data = "Missing survey Data"
    error_missing_liquid = "Not enough liquid"

    # headers_LDV:
    ws1.cell(row=row, column=col + 0, value="source_plates").font = Font(bold=True)
    ws1.cell(row=row, column=col + 1, value="source_well").font = Font(bold=True)
    ws1.cell(row=row, column=col + 2, value="volume").font = Font(bold=True)
    ws1.cell(row=row, column=col + 3, value="destination_well").font = Font(bold=True)
    ws1.cell(row=row, column=col + 4, value="destination_plates").font = Font(bold=True)
    ws1.cell(row=row, column=col + 5, value="compound").font = Font(bold=True)
    ws1.cell(row=row, column=col + 6, value="comments").font = Font(bold=True)

    # headers_PP:
    ws2.cell(row=row, column=col + 0, value="source_plates").font = Font(bold=True)
    ws2.cell(row=row, column=col + 1, value="source_well").font = Font(bold=True)
    ws2.cell(row=row, column=col + 2, value="volume").font = Font(bold=True)
    ws2.cell(row=row, column=col + 3, value="destination_well").font = Font(bold=True)
    ws2.cell(row=row, column=col + 4, value="destination_plates").font = Font(bold=True)
    ws2.cell(row=row, column=col + 5, value="compound").font = Font(bold=True)
    ws2.cell(row=row, column=col + 6, value="comments").font = Font(bold=True)

    if not starting_set:
        starting_set = 1

    for sets in range(starting_set, set_amount+1):
        for rows in set_compound_data:
            destination_plate = f"{sets}-{set_compound_data[rows]['destination_plate']}"
            destination_well = set_compound_data[rows]["destination_well"]
            compound = set_compound_data[rows]["compound"]
            volume_nl = float(set_compound_data[rows]["volume_nl"])
            volume_ul = volume_nl/1000
            sample_comment = set_compound_data[rows]["sample_comment"]
            source_plate_origin = set_compound_data[rows]["source_plate"]
            plate_type = set_compound_data[rows]["plate_type"]

            # Change between LDV and PP trans, depending on the source plate type
            if "pp" in source_plate_origin.casefold():
                ws = ws2
                row_pp += 1
                row = row_pp
            else:
                ws = ws1
                row_ldv += 1
                row = row_ldv

            # Check if there is survey data for the plate, and for the compound in the plate
            try:
                survey_layout[source_plate_origin][compound]
            except KeyError:
                source_plate = source_plate_origin
                source_well = error_missing_survey_data
            else:

                # Break is here to break out of the loop, to avoid looping over all the wells with the same compound in
                # them, for all the plate where the compound is.
                for temp_plates in survey_layout[source_plate_origin][compound]:
                    for wells in survey_layout[source_plate_origin][compound][temp_plates]:
                        try:
                            compound_list[source_plate_origin]
                        except KeyError:
                            dead_vol = dead_vol_ul[plate_type]
                        else:
                            if compound in compound_list[source_plate_origin]:
                                dead_vol = compound_list[source_plate_origin][compound]
                            else:
                                dead_vol = dead_vol_ul[plate_type]

                        if survey_layout[source_plate_origin][compound][temp_plates][wells] >= volume_ul + dead_vol:
                            source_plate = temp_plates
                            source_well = wells
                            survey_layout[source_plate_origin][compound][temp_plates][wells] -= volume_ul
                            breaking = True
                            break
                        else:
                            source_plate = f"{source_plate_origin}"
                            source_well = error_missing_liquid

                    if breaking:
                        breaking = False
                        break

            # This will colour cells where there is an error, and colour the headlines, to ensure notability.
            if source_well == error_missing_survey_data or source_well == error_missing_liquid:
                ws.cell(row=row, column=col + 0, value=f"{source_plate}_{compound}").fill = PatternFill(start_color='B284BE',
                                                                                        end_color='B284BE',
                                                                                        fill_type='solid')
                ws.cell(row=row, column=col + 1, value=source_well).fill = PatternFill(start_color='B284BE',
                                                                                       end_color='B284BE',
                                                                                       fill_type='solid')
                ws.cell(row=1, column=1).fill = PatternFill(start_color='B284BE', end_color='B284BE',
                                                                         fill_type='solid')
                ws.cell(row=1, column=2).fill = PatternFill(start_color='B284BE', end_color='B284BE',
                                                            fill_type='solid')
                failed_plate_compound = f"{source_plate}_{compound}"
                try:
                    failed_plate[failed_plate_compound]
                except KeyError:
                    failed_plate[failed_plate_compound] = {"plate": source_plate, "vol_ul": 0.0, "compound": compound, "trans_counter": 0}

                failed_plate[failed_plate_compound]["vol_ul"] += volume_ul
                failed_plate[failed_plate_compound]["trans_counter"] += 1

            # Writes the data for each transfer.
            else:
                ws.cell(row=row, column=col + 0, value=source_plate)
                ws.cell(row=row, column=col + 1, value=source_well)
            ws.cell(row=row, column=col + 2, value=volume_nl)
            ws.cell(row=row, column=col + 3, value=destination_well)
            ws.cell(row=row, column=col + 4, value=destination_plate)
            ws.cell(row=row, column=col + 5, value=compound)
            ws.cell(row=row, column=col + 6, value=sample_comment)
            row += 1

    ws = ws0
    headlines_report = ["plates", "vol_ul", "compound", "trans_counter"]
    row = 1
    col = 1
    col_ldv = 1
    col_pp = col_ldv + len(headlines_report)
    row_ldv = 2
    row_pp = row_ldv

    # Headlines
    for counter in range(2):
        for headlines in headlines_report:
            ws.cell(row=row, column=col, value=headlines)
            col += 1

    for index, plates in enumerate(failed_plate):
        if "LDV" in plates:
            row = row_ldv
            col = col_ldv
            row_ldv += 1

        if "PP" in plates:
            row = row_pp
            col = col_pp
            row_pp += 1

        ws.cell(row=row, column=col, value=failed_plate[plates]["plate"])
        ws.cell(row=row, column=col + 1, value=failed_plate[plates]["vol_ul"])
        ws.cell(row=row, column=col + 2, value=failed_plate[plates]["compound"])
        ws.cell(row=row, column=col + 3, value=failed_plate[plates]["trans_counter"])



    wb.save(save_file)


def new_worklist(survey_folder, plate_layout_folder, file_trans, set_amount, dead_vol_ul, save_location, save_file_name, starting_set=None): # TODO description
    save_file = f"{save_location}/{save_file_name}.xlsx"
    survey_data = get_survey_csv_data(survey_folder)
    plate_layout = well_compound_list(plate_layout_folder)
    _, _, set_compound_data = get_all_trans_data(file_trans)

    survey_layout = _compound_to_survey(plate_layout, survey_data)

    _write_new_worklist(set_compound_data, survey_layout, dead_vol_ul, set_amount, save_file, starting_set)

    print("done")

if __name__ == "__main__":
    trans_data_folder = "C:/Users/phch/Desktop/more_data_files/2022-11-22"
    plate_layout_folder = "D:/plate_layout"
    data_location = "C:/Users/phch/Desktop/more_data_files/2022-11-22"
    file_name = "test_trans_report"
    save_location = "C:/Users/Openscreen/Desktop/"
    all_trans_file = "C:/Users/Openscreen/Desktop/more_data_files/all_trans.xlsx"
    path = "C:/Users/phch/Desktop/echo_data"

    file_trans = "D:/all_trans.xlsx"
    set_amount = 60
    starting_set = 17
    dead_vol_ul = {"LDV": 2.5, "PP": 15}
    save_file_name = "test_set_trans"

    survey_folder = "C:/Users/Openscreen/Desktop/surveys/170123"

    new_worklist(survey_folder, plate_layout_folder, file_trans, set_amount, dead_vol_ul, save_location, save_file_name)



    # trans_report_controller(trans_data_folder, plate_layout_folder, all_trans_file, data_location, file_name, save_location)
    # well_report(all_trans_file)

    # print(file_names(path))

