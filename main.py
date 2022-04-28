import click as click
import copy
import openpyxl
import os
import json
from excel import *

SUMMARY_XLSX_FILE = '汇总.xlsx'


def check_file_ext(file_name, file_ext):
    if not str(file_name).endswith(file_ext):
        raise ValueError("invalid file extend, file is " + file_name + ", ext is " + file_ext)


def list_dir_files(dir_path: str, file_ext):
    if not os.path.isdir(dir_path):
        print("create directory " + target)
        os.mkdir(target)

    res = []
    files = os.listdir(dir_path)
    for file in files:
        if os.path.isdir(file) or not str(file).endswith(file_ext) or file == SUMMARY_XLSX_FILE:
            continue
        res.append(file)
    return res


def copy_file(src, target, file_ext='xlsx', copy_image=True):
    check_file_ext(src, file_ext)
    check_file_ext(target, file_ext)
    copy_xlsx(src, target, copy_image)


def copy_dir(src, target, file_ext='xlsx', copy_image=True):
    files = list_dir_files(src, file_ext)
    for file in files:
        src_file = src + os.sep + file
        dst_file = target + os.sep + file
        copy_file(src_file, dst_file, file_ext, copy_image)


def generate_summary_file(dir_path: str, file_ext='xlsx'):
    sheet_datas = []
    files = list_dir_files(dir_path, file_ext)
    for file in files:
        src_file = dir_path + os.sep + file
        rows = get_xlsx_row_detail(src_file)
        if not rows:
            continue
        rows.append(['', ''])
        rows.append(['总计', get_xlsx_total_rows(src_file)])
        s = {
            'title': file,
            'info': ['点位', '数量'],
            'data': rows
        }
        sheet_datas.append(s)
    write_xlsx(dir_path + os.sep + SUMMARY_XLSX_FILE, sheet_datas)


def statistic_xlsx_rows(dir_path: str, file_ext='xlsx'):
    files = list_dir_files(dir_path, file_ext)
    for file in files:
        src_file = dir_path + os.sep + file
        total_line = get_xlsx_total_rows(src_file)
        print(file + ', 共有: ' + str(total_line))


if __name__ == "__main__":
    file_ext = 'xlsx'
    d='/Users/kun/Desktop/市民卡业务带图'
    copy_dir(d, d, file_ext)
    generate_summary_file(d, file_ext)
    statistic_xlsx_rows(d, file_ext)
