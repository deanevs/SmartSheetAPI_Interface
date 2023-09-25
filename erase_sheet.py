from smartsheet_interface import ss_interface, settings

# sheet_to_erase_rows = settings.SH_CNL_ID

def erase_sheet_data(sheet):
    sh = ss_interface.Sheet(settings.smartsheet_token, sheet)
    sh.delete_all_rows()

# erase_sheet_data(sheet_to_erase_rows)