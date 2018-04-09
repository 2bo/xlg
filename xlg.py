# -*- coding: utf-8 -*-
import sys
import re
import openpyxl
import argparse
import datetime


def main():

    parser = argparse.ArgumentParser(
        description='Search and replace with regex in xlsx/xlsm')
    parser.add_argument('src_file', help='target file')
    parser.add_argument('search_text', help='search text with regrex')
    parser.add_argument('-r', '--replace',
                        help='replace text with regrex', metavar='replace_text', nargs=1)
    parser.add_argument(
        '-s', '--saveas', help='save as replaced file', metavar='save_as', nargs=1)

    args = parser.parse_args()
    print(args)

    if args.replace is None:
        try:
            xlsx_search(args.src_file, args.search_text)
        except Exception as exception:
            print(exception)
            sys.exit(1)
    else:
        try:
            if args.saveas is None:
                xlsx_replace(args.src_file, args.search_text, args.replace[0])
            else:
                xlsx_replace(args.src_file, args.search_text,
                             args.replace[0], args.saveas[0])
        except Exception as exception:
            print(exception)
            sys.exit(1)


def xlsx_search(filename, regrex_str):
    try:
        workbook = openpyxl.load_workbook(filename, read_only=True)
    except:
        raise

    regrex = re.compile(regrex_str)

    # Loop Sheets
    for sheet in workbook.worksheets:
        # Loop all cells
        for row in sheet.rows:
            for cell in row:
                if cell.value is not None:
                    # regrex search
                    value_str = str(cell.value)
                    match_obj = regrex.search(value_str)
                    if match_obj is not None:
                        # Output result
                        print('{sheet}\t{coordinate}\t{value_str}'.format(
                            sheet=sheet.title, coordinate=cell.coordinate, value_str=value_str))


def xlsx_replace(filename, regrex_str, replace_str, saveas=None):
    try:
        workbook = openpyxl.load_workbook(filename)
    except:
        raise

    for sheet in workbook.worksheets:
        for row in sheet.rows:
            for cell in row:
                if cell.value is not None:
                    value_str = str(cell.value)
                    number_format = cell.number_format
                    print(type(cell.value))
                    replaced_str = re.sub(regrex_str, replace_str, value_str)
                    
                    if value_str != replaced_str:
                        print('{sheet}\t{coordinate}\t{value_str}\t->\t{replaced_str}'.format(
                            sheet=sheet.title, coordinate=cell.coordinate, value_str=value_str, replaced_str=replaced_str))
                        if isinstance(cell.value,str):
                            cell.value = replaced_str
                        elif is_int(replaced_str):
                            cell.value = int(replaced_str)
                        elif is_float(replace_str):
                            cell.value = float(replaced_str)
                            
                        cell.value = replaced_value
                        cell.number_format = number_format
    if saveas is None:
        saveas = filename
    workbook.save(saveas)

def is_int(target_str):
    if re.match(r'^[-+]?[0-9]+$', target_str):
        return True
    else:
        return False

def is_float(target_str):
    if re.match(r'[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?',target_str):
        return True
    else:
        return False

def is_datetime(target_str):
    try:
        datetime.datetime.strptime(target_str, '%Y-%m%d %H:%M:%S')
        return True
    except ValueError:
        return False

if __name__ == '__main__':
    main()
