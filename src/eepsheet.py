class EepSheet:
    """EEP Excel sheet class.
    """
    colpos = {
        'region': 0,
        'location': 1,
        'school': 2,
        'student_name': 3,
        'sex': 4,
        'graduation_year': 5,
        'student_donor_id': 6,
        'student_donor_name': 7,
        'student_donor_donation_amount_local': 8,
        'comment': 9,
        'import_order_number': 10,
        'auto_student_number': 11,
        'auto_donor_student_count_number': 12,
        'school_name_length': 13,
    }
    sheet = None

    def __init__(self, sheet):
        self.sheet = sheet

    def max_rows(self):
        return self.sheet.nrows

    def get_sheet_row_hi(self):
        """Return sheet's 'last row' based on whether the first 3 columns' values.
        """
        excel_row_lo = 0
        excel_row_hi = 0

        try:
            sheet = self.sheet
            for rownum in range(excel_row_lo, self.max_rows()):
                try:
                    if (
                        not sheet.cell(rownum, 0).value and
                        not sheet.cell(rownum + 1, 0).value and
                        not sheet.cell_value(rownum, 1) and
                        not sheet.cell_value(rownum + 1, 1) and
                        not sheet.cell_value(rownum, 2) and
                        not sheet.cell_value(rownum + 1, 2)
                    ):
                        excel_row_hi = rownum
                        break;
                except:
                    print 'Error occured trying to get sheet_row_hi'
                    break;
        except:
            excel_row_hi = rownum

        print 'Last Excel row for sheet {}: {}'.format(
            sheet.name.encode('utf-8'),
            excel_row_hi
        )

        return excel_row_hi

    def cell_value(self, row, col):
        """Returns cell value and trim white spaces if the value is a string.
        """
        try:
            val = self.sheet.cell_value(row, col)
            if hasattr(val, 'strip'):
                return val.strip()
            else:
                return val
        except:
            return ''

    def col_values(self, col_pos, start_row, end_row):
        """Get all values for a column from start_row to end_row.
        """
        return self.sheet.col_values(col_pos, start_row, end_row)

    def get_region(self, row):
        return self.cell_value(row, self.colpos['region'])

    def get_location(self, row):
        return self.cell_value(row, self.colpos['location'])

    def get_school(self, row):
        return self.cell_value(row, self.colpos['school'])

    def get_student_name(self, row):
        return self.cell_value(row, self.colpos['student_name'])

    def get_graduation_year(self, row):
        return self.cell_value(row, self.colpos['graduation_year'])
    
    def get_donor_id(self, row):
        return self.cell_value(row, self.colpos['student_donor_id'])
