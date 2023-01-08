import datetime

from excel import *


class TableColumnDefine:
    def __init__(self, idx, name):
        self.idx = idx
        self.name = name


SERIAL_NO_COLUMN = TableColumnDefine(1, '序号')
NAME_COLUMN = TableColumnDefine(2, '姓名')
PHONE_COLUMN = TableColumnDefine(3, '电话')
REMARK_COLUMN = TableColumnDefine(4, '备注')
SUCCEED_COLUMN = TableColumnDefine(5, '是否领卡')
FAILED_COLUMN = TableColumnDefine(6, '是否失败')

SUMMARY_SHEET_NAME = '汇总'


def load_all_name(file):
    wb = load_workbook(file)
    all_ws = get_all_worksheets(wb)

    count = 0
    name_addr_mapping = {}

    for ws in all_ws:
        assert isinstance(ws, Worksheet)
        if ws.title == SUMMARY_SHEET_NAME:
            continue

        for i in range(2, ws.max_row+1):
            idx = get_cell_value(ws, i, SERIAL_NO_COLUMN.idx)
            name = get_cell_value(ws, i, NAME_COLUMN.idx)

            if not idx or not name:
                break
            count = count + 1
            if name not in name_addr_mapping:
                name_addr_mapping[name] = []
            name_addr_mapping[name].append(ws.title)

    print(count)
    return name_addr_mapping


def match_excel(mapping, file):
    wb = load_workbook(file)
    ws = get_all_worksheets(wb)[0]
    for i in range(2, ws.max_row):
        name = get_cell_value(ws, i, name_col_idx)
        match_type = None
        match_detail = None
        if name in mapping:
            val = mapping[name]
            if len(mapping[name]) == 1:
                match_type = '唯一匹配'
                match_detail = val[0]
            elif len(mapping[name]) > 1:
                match_type = '多个匹配'
                match_detail = ", ".join(val)
        else:
            match_type = '不匹配'

        set_cell_value(ws, i, match_type_idx, match_type)
        set_cell_value(ws, i, match_detail_idx, match_detail)

    save_workbook(wb, file)


if __name__ == "__main__":
    src_file = '/Users/kun/Desktop/云文档/市民卡业务-邵祥/市民卡-延安路支行.xlsx'
    mapping = load_all_name(src_file)

    name_col_idx = 1
    match_type_idx = 2
    match_detail_idx = 3

    dst_file = '/Users/kun/Desktop/云文档/延安路支行汇总数据.xlsx'
    match_excel(mapping, dst_file)


