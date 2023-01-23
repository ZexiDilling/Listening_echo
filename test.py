def mail_report_sender(temp_file_name, config):
    # Reads the temp_file where all the trans file have been written to

    file_list = read_temp_list_file(temp_file_name, config)

    # Setup the report
    report_name = f"Report_{date.today()}"
    save_location = config["Folder"]["out"]
    temp_counter = 2
    full_path = f"{save_location}/{report_name}.xslx"
    while path.exists(full_path):
        report_name = f"{report_name}_{temp_counter}"
        temp_counter += 1
        full_path = f"{save_location}/{report_name}.xslx"

    # Create the report file, and saves it.
    skipped_well_controller(file_list, report_name, save_location)

    # sends an E-mail, with the report included
    msg_subject = f"Final report for transfer: {date.today()}"
    data = full_path
    e_mail_type = "final_report"
    mail_setup(msg_subject, data, config, e_mail_type)