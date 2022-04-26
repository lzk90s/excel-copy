import click as click
import copy
import openpyxl
import os
from openpyxl.utils import get_column_letter


def _get_excel_total_row(path: str):
    wb = openpyxl.load_workbook(path)
    sheet_names = wb.sheetnames
    sheet_names.sort()
    total_row = 0
    for sheet_name in sheet_names:
        sheet = wb[sheet_name]
        row = sheet.max_row
        total_row += row - 1
    return total_row


def _copy_xlsx(path: str, save_path: str, copy_image=True):
    wb = openpyxl.load_workbook(path)
    wb2 = openpyxl.Workbook()

    sheet_names = wb.sheetnames
    sheet_names.sort()
    for sheet_name in sheet_names:
        print(sheet_name)
        sheet = wb[sheet_name]
        sheet2 = wb2.create_sheet(sheet_name)

        # tab颜色
        sheet2.sheet_properties.tabColor = sheet.sheet_properties.tabColor

        # 开始处理合并单元格形式为“(<CellRange A1：A4>,)，替换掉(<CellRange 和 >,)' 找到合并单元格
        wm = list(sheet.merged_cells)
        if len(wm) > 0:
            for i in range(0, len(wm)):
                cell2 = str(wm[i]).replace('(<CellRange ', '').replace('>,)', '')
                sheet2.merge_cells(cell2)

        for i, row in enumerate(sheet.iter_rows()):
            sheet2.row_dimensions[i + 1].height = sheet.row_dimensions[i + 1].height
            for j, cell in enumerate(row):
                sheet2.column_dimensions[get_column_letter(j + 1)].width = sheet.column_dimensions[
                    get_column_letter(j + 1)].width
                sheet2.cell(row=i + 1, column=j + 1, value=cell.value)

                # 设置单元格格式
                source_cell = sheet.cell(i + 1, j + 1)
                target_cell = sheet2.cell(i + 1, j + 1)
                target_cell.fill = copy.copy(source_cell.fill)
                if source_cell.has_style:
                    target_cell._style = copy.copy(source_cell._style)
                    target_cell.font = copy.copy(source_cell.font)
                    target_cell.border = copy.copy(source_cell.border)
                    target_cell.fill = copy.copy(source_cell.fill)
                    target_cell.number_format = copy.copy(source_cell.number_format)
                    target_cell.protection = copy.copy(source_cell.protection)
                    target_cell.alignment = copy.copy(source_cell.alignment)

        if copy_image:
            for image in sheet._images:
                sheet2.add_image(image)

    if 'Sheet' in wb2.sheetnames:
        del wb2['Sheet']
    wb2.save(save_path)

    wb.close()
    wb2.close()

    print('Done.')


def check_file_ext_filterend(file_name, file_ext_filter):
    if not str(file_name).endswith(file_ext_filter):
        raise ValueError("invalid file extend, file is " + file_name + ", ext is " + file_ext_filter)


def copy_file(src, target, file_ext_filter='xlsx', copy_image=True):
    check_file_ext_filterend(src, file_ext_filter)
    check_file_ext_filterend(target, file_ext_filter)
    _copy_xlsx(src, target, copy_image)


def copy_dir(src, target, file_ext_filter='xlsx', copy_image=True):
    if not os.path.isdir(target):
        print("create directory " + target)
        os.mkdir(target)

    files = os.listdir(src)
    for file in files:
        if os.path.isdir(file) or not str(file).endswith(file_ext_filter):
            continue
        src_file = src + os.sep + file
        dst_file = target + os.sep + file
        copy_file(src_file, dst_file, file_ext_filter, copy_image)


def get_excel_row_in_dir(dir_path: str, file_ext_filter='xlsx'):
    if not os.path.isdir(dir_path):
        print("create directory " + target)
        os.mkdir(target)

    files = os.listdir(dir_path)
    for file in files:
        if os.path.isdir(file) or not str(file).endswith(file_ext_filter):
            continue
        src_file = dir_path + os.sep + file
        total_line = _get_excel_total_row(src_file)
        print(file + ', 共有: ' + str(total_line))


@click.command()
@click.option('--mode', '-m', help='mode', default='dir')
@click.option('--src', "-s", help='source file', required=True)
@click.option('--target', "-t", help='target file', required=True)
@click.option('--file_ext_filter', "-e", help='file extend', default='xlsx')
@click.option('--copy_image', '-i', help='copy image', default=True)
def main(mode, src, dst, file_ext_filter, copy_image):
    copy_handler = copy_dir if 'dir' == mode else copy_file
    copy_handler(src, dst, file_ext_filter, copy_image)


if __name__ == "__main__":
    #copy_dir('/Users/kun/Desktop/市民卡业务带图', '/Users/kun/Desktop/市民卡业务带图1')
    get_excel_row_in_dir('/Users/kun/Desktop/市民卡业务带图')
