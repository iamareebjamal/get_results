import re
import xlrd

fac_no_pattern = re.compile('[0-9]{2}[PLCMEK][EK][B][0-9]{3}')


def is_fac_no(fac_no):
    return fac_no_pattern.match(fac_no) != None


class Collector(object):
    __worksheet = None

    def __init__(self, file_path):
        workbook = xlrd.open_workbook(file_path)
        self.__worksheet = workbook.sheet_by_index(0)

    def get_start_index(self):
        sheet = self.__worksheet
        for row in xrange(sheet.nrows):
            for col in xrange(sheet.ncols):
                value = sheet.cell(row, col).value
                if not value or value == '':
                    pass
                elif isinstance(value, basestring) and is_fac_no(value):
                    return (row, col)

    def collect(self):
        sheet = self.__worksheet
        row, col = self.get_start_index()
        keys = ['fac_no', 'en_no', 'name']
        return [dict(zip(keys, [str(cell.value) for cell in sheet.row(row_index)[col:]])) for row_index in xrange(row, sheet.nrows)]
