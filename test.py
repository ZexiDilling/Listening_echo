from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl import Workbook, load_workbook
import xml.etree.ElementTree as ET

from get_data import get_xml_trans_data_printing_wells, well_compound_list


def rename_source_plates(trans_data, prefix_dict):


    for trans in trans_data:
        temp_source_plate = trans_data[trans]["source_plate"]
        temp_date = trans_data[trans]["date"]
        for prefix in prefix_dict:
            if prefix_dict[prefix]["start"] < temp_date < prefix_dict[prefix]["end"]:
                trans_data[trans]["source_plate"] = f"{prefix}_{temp_source_plate}"


def write_report(trans_data, compound_data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    row = 1
    col = 1

    # Headers
    ws.cell(row=row, column=col + 0, value="Source Plates").font = Font(bold=True)
    ws.cell(row=row, column=col + 1, value="Source Well").font = Font(bold=True)
    ws.cell(row=row, column=col + 2, value="Volume").font = Font(bold=True)
    ws.cell(row=row, column=col + 3, value="Destination Well").font = Font(bold=True)
    ws.cell(row=row, column=col + 4, value="Destination Plates").font = Font(bold=True)
    ws.cell(row=row, column=col + 5, value="Compound").font = Font(bold=True)
    ws.cell(row=row, column=col + 6, value="Comments").font = Font(bold=True)

    for trans in trans_data:




def excel_controller(trans_data_folder, plate_layout_folder, data_location, file_name, save_location):
    trans_data = get_xml_trans_data_printing_wells(trans_data_folder)
    compound_data = well_compound_list(plate_layout_folder)

    prefix_on = True
    prefix_dict = {"OLD": {"start": "2022-11-22", "end": "2022-12-01"},
              "NEW": {"start": "2022-10-01", "end": "2022-10-22"}}
    if prefix_on:
        rename_source_plates(trans_data, prefix_dict)

    write_report(trans_data, compound_data)





    return "Done"


if __name__ == "__main__":
    trans_data_folder = "C:/Users/phch/Desktop/more_data_files/2022-11-22"
    plate_layout_folder = "C:/Users/phch/Desktop/more_data_files/plate_layout"
    data_location = "C:/Users/phch/Desktop/more_data_files/2022-11-22"
    file_name = "test_trans_report"
    save_location = "C:/Users/phch/Desktop/more_data_files/"

    path = "C:/Users/phch/Desktop/echo_data"

    excel_controller(trans_data_folder, plate_layout_folder, data_location, file_name, save_location)


    # print(file_names(path))

