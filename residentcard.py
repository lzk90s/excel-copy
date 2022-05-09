from core import excel
import datetime
from core.util import *

# 序号列索引
SERIAL_NO_COLUMN_INDEX = 2
SERIAL_NO_COLUMN_NAME = '序号'
# 名字列索引
NAME_COLUMN_INDEX = 3
NAME_COLUMN_NAME = '姓名'


def validate_xlsx(files):
    for file in files:
        assert excel.check_cell_value(file, 1, SERIAL_NO_COLUMN_INDEX, SERIAL_NO_COLUMN_NAME) is True, file
        assert excel.check_cell_value(file, 1, NAME_COLUMN_INDEX, NAME_COLUMN_NAME) is True, file

        d1 = excel.get_row_detail(file, SERIAL_NO_COLUMN_INDEX)
        d2 = excel.get_row_detail(file, NAME_COLUMN_INDEX)
        assert len(d1) == len(d2), file

        for i in range(len(d1)):
            k1 = d1[i]
            k2 = d2[i]
            assert k1[1] == k2[1], k1[0]


def sort_xlsx(files):
    def sheet_sort_key(v: str):
        x = v.title
        date = x[0:x.index('-')]
        d = datetime.datetime.strptime(date, "%m.%d")
        return d

    for file in files:
        excel.sort_by_title(file, lambda x: sheet_sort_key(x))


def generate_summary_file(files: str, out_file):
    sheet_datas = []
    for file in files:
        rows = excel.get_row_detail(file, SERIAL_NO_COLUMN_INDEX)
        if not rows:
            continue
        rows.append(['', ''])
        rows.append(['总计', excel.get_total_rows(file, SERIAL_NO_COLUMN_INDEX)])
        file_name = file if file.rindex('/') < 0 else file[file.rindex('/') + 1:]
        s = {
            'title': file_name,
            'info': ['点位', '数量'],
            'data': rows
        }
        sheet_datas.append(s)
    excel.write_workbook(out_file, sheet_datas)


def print_xlsx_rows(files: str):
    for file in files:
        total_line = excel.get_total_rows(file, SERIAL_NO_COLUMN_INDEX)
        print(file + ', 共有: ' + str(total_line))


if __name__ == "__main__":
    d = '/Users/kun/Desktop/市民卡业务带图'

    files = list_dir_files(d, ['.xlsx'], ['汇总.xlsx'])

    validate_xlsx(files)
    sort_xlsx(files)
    generate_summary_file(files, d + os.sep + '汇总.xlsx')
    print_xlsx_rows(files)
