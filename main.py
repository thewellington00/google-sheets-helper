from google_sheets import Spreadsheet

if __name__ == '__main__':
    spreadsheet_id = '<insert here>'
    spreadsheet = Spreadsheet(spreadsheet_id)
    worksheet = spreadsheet.get_worksheet('Sheet1')
    worksheet.cross_join_ranges_to_clipboard('b26:b30', 'c26:c33')