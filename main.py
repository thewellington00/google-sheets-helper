from google_sheets import Spreadsheet

if __name__ == '__main__':
    spreadsheet = Spreadsheet('12-OzoD7ujbC73MCBK5kwQ47UMxIuIrW3ZMi_pX58Hbk')
    worksheet = spreadsheet.get_worksheet('raw_data')
    named_ranges = worksheet.create_named_ranges_from_headers()