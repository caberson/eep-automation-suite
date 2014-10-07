class EepSheet:
    COL_REGION = 0
    COL_LOCATION = 1
    COL_SCHOOL = 2
    COL_STUDENT_NAME = 3
    COL_SEX = 4
    COL_GRADUATION_YEAR = 5
    COL_STUDENT_DONOR_ID = 6
    COL_STUDENT_DONOR_NAME = 7
    COL_STUDENT_DONOR_DONATION_AMOUNT_LOCAL = 8
    COL_COMMENT = 9
    COL_IMPORT_ORDER_NUMBER = 10
    COL_AUTO_STUDENT_NUMBER = 11
    COL_AUTO_DONOR_STUDENT_COUNT_NUMBER = 12
    COL_SCHOOL_NAME_LENGTH = 13

    sheet = None

    def __init__(self, sheet):
        self.sheet = sheet

    def max_rows(self):
        return self.sheet.nrows

    def cell_value(self, row, col):
        try:
            val = self.sheet.cell_value(row, col)
            if hasattr(val, 'strip'):
                return val.strip()
            else:
                return val
        except:
            return ''

    def get_region(self, row):
        return self.cell_value(row, self.COL_REGION)

    def get_location(self, row):
        return self.cell_value(row, self.COL_LOCATION)

    def get_school(self, row):
        return self.cell_value(row, self.COL_SCHOOL)
