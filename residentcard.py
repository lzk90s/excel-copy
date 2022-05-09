from core import excel
import datetime
from core.util import *

# 序号列索引
SERIAL_NO_COLUMN_INDEX = 2
SERIAL_NO_COLUMN_NAME = '序号'
# 名字列索引
NAME_COLUMN_INDEX = 3
NAME_COLUMN_NAME = '姓名'


def parse_sheet_name(sheet_name: str):
    ss = sheet_name.split('-')
    date = datetime.datetime.strptime(ss[0], "%m.%d")
    address = ss[1]
    operator = '未知' if len(ss) < 3 else ss[2]

    return date, address, operator


def validate_xlsx(files):
    for file in files:
        assert excel.check_cell_value(file, 1, SERIAL_NO_COLUMN_INDEX, SERIAL_NO_COLUMN_NAME) is True, file
        assert excel.check_cell_value(file, 1, NAME_COLUMN_INDEX, NAME_COLUMN_NAME) is True, file

        d1 = excel.parse_sheets(file, lambda ws: [ws.title, excel.get_max_row(ws, SERIAL_NO_COLUMN_NAME)])
        d2 = excel.parse_sheets(file, lambda ws: [ws.title, excel.get_max_row(ws, NAME_COLUMN_NAME)])
        assert len(d1) == len(d2), file

        for i in range(len(d1)):
            k1 = d1[i]
            k2 = d2[i]
            assert k1[1] == k2[1], k1[0]


def default_sort_key(title):
    date, addr, operator = parse_sheet_name(title)
    return date


def sort_xlsx(files):
    for file in files:
        excel.sort_by_title(file, lambda x: default_sort_key(x.title))


def generate_summary_file(files: str, out_file):
    sheet_datas = []
    for file in files:
        rows = excel.parse_sheets(file, lambda ws: [ws.title, excel.get_max_row(ws, SERIAL_NO_COLUMN_NAME)])
        if not rows:
            continue

        rows.append(['', ''])
        rows.append(['总计', excel.get_total_rows(file, SERIAL_NO_COLUMN_NAME)])
        file_name = file if file.rindex('/') < 0 else file[file.rindex('/') + 1:]
        s = {
            'sheet_name': file_name,
            'head': ['点位', '数量'],
            'data': rows
        }
        sheet_datas.append(s)
    excel.write_workbook(out_file, sheet_datas)


def generate_personal_performance(files: str, out_file):
    sheet_datas = []
    operator_map = {}
    for file in files:
        rows = excel.parse_sheets(file, lambda ws: [ws.title, excel.get_max_row(ws, SERIAL_NO_COLUMN_NAME)])
        if not rows:
            continue

        tmp_rows = []
        for r in rows:
            sheet_name = r[0]
            total_num = r[1]
            date, address, operator = parse_sheet_name(sheet_name)
            tmp_row = [sheet_name, total_num]
            if operator in operator_map.keys():
                operator_map.get(operator).append(tmp_row)
            else:
                operator_map[operator] = [tmp_row]
            tmp_rows.append(tmp_row)

    for k, v in operator_map.items():
        v.sort(key=lambda x: default_sort_key(x[0]))
        s = {
            'sheet_name': k,
            'head': ['点位', '数量'],
            'data': v
        }
        sheet_datas.append(s)
    excel.write_workbook(out_file, sheet_datas)


def print_xlsx_rows(files: str):
    for file in files:
        total_line = excel.get_total_rows(file, SERIAL_NO_COLUMN_NAME)
        print(file + ', 共有: ' + str(total_line))


if __name__ == "__main__":
    d = '/Users/kun/Desktop/市民卡业务带图'

    files = list_dir_files(d, ['.xlsx'], ['汇总.xlsx', '绩效.xlsx'])

    validate_xlsx(files)
    sort_xlsx(files)
    print_xlsx_rows(files)
    generate_summary_file(files, d + os.sep + '汇总.xlsx')
    generate_personal_performance(files, d + os.sep + '绩效.xlsx')
