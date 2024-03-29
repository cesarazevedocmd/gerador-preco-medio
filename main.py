import copy
import datetime
import json
import os.path as path
from datetime import timedelta
from os import listdir

import xlrd
import xlwt

import my_asset_class as Asset

operations = []
fii = []
assets = []
assets_per_year = {}


def calculate_average_price(op, new_op):
    new_qty = op.qty + new_op.qty
    if new_qty != 0:
        return ((op.qty * op.price) + (new_op.qty * new_op.price)) / new_qty
    return 0


def asset(op):
    return op.asset


def sort_assets(assets_list):
    assets_list.sort(key=asset)


def validate_new_asset_name(asset_name):
    for dic_name in config_json["new_names"]:
        if asset_name in dic_name:
            return dic_name[asset_name]

    return asset_name


def add_new_operation(new_op, operation_type):

    new_op.asset = validate_new_asset_name(new_op.asset)

    for op in operations:
        if new_op.asset == op.asset:
            if operation_type == 'C':
                new_qty = op.qty + new_op.qty
                new_price = calculate_average_price(op, new_op)
                op.qty = new_qty
                op.price = float(int(new_price * 1000) / 1000)
                return
            elif operation_type == 'V':
                new_qty = op.qty - new_op.qty
                op.qty = new_qty
                if new_qty == 0:
                    op.price = 0.0
                return
    if operation_type == 'V':
        new_op.qty = (new_op.qty * -1)
    operations.append(new_op)


def read_cei_excel_file(file):
    file_b3 = xlrd.open_workbook(file)
    sheet = file_b3.sheet_by_index(0)

    found_buy_title = False
    found_empty_line = False
    starting_asset_read = False

    for row in range(0, sheet.nrows):
        value = str(sheet.cell(row, 3).value)

        if value == 'INFORMAÇÕES DE NEGOCIAÇÃO DE ATIVOS':
            found_buy_title = True
            continue

        if (value == '' and found_buy_title) and not found_empty_line:
            found_empty_line = True
            continue

        if value == 'Cód' and found_buy_title and found_empty_line:
            starting_asset_read = True
            continue

        if starting_asset_read:
            if value != '':
                asset = str(sheet.cell(row, 3).value).strip()
                data_compra = str(sheet.cell(row, 10))
                qty_buy = int(sheet.cell(row, 18).value)
                qty_sell = int(sheet.cell(row, 24).value)
                average_buy_price = str(sheet.cell(row, 34).value)
                average_sell_price = str(sheet.cell(row, 43).value)
                qnt_liquida = str(sheet.cell(row, 49).value)
                operation_type = str(sheet.cell(row, 54).value)
                operation_type = 'V' if operation_type == 'VENDIDA' else 'C'

                if asset[-1] == 'F':
                    asset = asset[0:len(asset) - 1]

                qty = qty_sell if operation_type == 'V' else qty_buy
                price = average_sell_price if operation_type == 'V' else average_buy_price

                op = Asset.MyAsset(asset=asset, qty=qty, price=price)
                add_new_operation(op, operation_type)
            else:
                break


def read_inter_excel_file(file):
    file_b3 = xlrd.open_workbook(file)
    sheet = file_b3.sheet_by_index(0)

    starting_asset_read = False

    for row in range(0, sheet.nrows):
        if starting_asset_read:
            value = str(sheet.cell(row, 3).value)
        else:
            value = str(sheet.cell(row, 0).value)

        if value == 'PRAÇA':
            starting_asset_read = True
            continue

        if starting_asset_read:
            if value == 'SUBTOTAL:':
                continue

            if value != '' and value != 'None':
                operation_type = str(sheet.cell(row, 1).value)
                asset = str(sheet.cell(row, 3).value).split(' ')[0]
                qnt = int(sheet.cell(row, 5).value)
                price = str(sheet.cell(row, 6).value)

                if asset[-1] == 'F':
                    asset = asset[0:len(asset) - 1]

                op = Asset.MyAsset(asset=asset, qty=qnt, price=price)
                add_new_operation(op, operation_type)
            else:
                starting_asset_read = False


def create_spreadsheet_with_asset_per_year(xls):
    sheet_per_year = xls.add_sheet('ASSETS_PER_YEAR')

    collum_a = 0
    collum_b = 1
    collum_c = 2
    collum_d = 3
    for year_key in assets_per_year.keys():
        line = 0
        sheet_per_year.write(line, collum_a, str(year_key))

        line += 1
        sheet_per_year.write(line, collum_a, 'ASSET')
        sheet_per_year.write(line, collum_b, 'QTY')
        sheet_per_year.write(line, collum_c, 'PRICE')
        sheet_per_year.write(line, collum_d, 'TOTAL')

        for operation in assets_per_year.get(year_key):
            line += 1
            sheet_per_year.write(line, collum_a, operation.asset)
            sheet_per_year.write(line, collum_b, operation.qty)
            sheet_per_year.write(line, collum_c, operation.price)
            sheet_per_year.write(line, collum_d, (operation.price * operation.qty))

        jump_collum = 5
        collum_a += jump_collum
        collum_b += jump_collum
        collum_c += jump_collum
        collum_d += jump_collum


def create_spreadsheet_with_assets(xls):
    sheet = xls.add_sheet('ASSETS')

    collum_zero = 0
    collum_one = 1
    collum_two = 2
    collum_five = 5
    collum_six = 6
    collum_seven = 7

    line = 0
    sheet.write(line, collum_zero, 'ASSET')
    sheet.write(line, collum_one, 'QTY')
    sheet.write(line, collum_two, 'PRICE')

    for operation in assets:
        line += 1
        sheet.write(line, collum_zero, operation.asset)
        sheet.write(line, collum_one, operation.qty)
        sheet.write(line, collum_two, operation.price)

    line = 0
    sheet.write(line, collum_five, 'FII')
    sheet.write(line, collum_six, 'QTY')
    sheet.write(line, collum_seven, 'PRICE')

    for operation in fii:
        line += 1
        sheet.write(line, collum_five, operation.asset)
        sheet.write(line, collum_six, operation.qty)
        sheet.write(line, collum_seven, operation.price)


def write_excel_file(name_result_file):
    xls = xlwt.Workbook()

    create_spreadsheet_with_assets(xls)

    create_spreadsheet_with_asset_per_year(xls)

    xls.save(name_result_file)


def remove_asset_zero_qty(assets_list):
    new_assets_with_zero_qty = []
    for op in assets_list:
        if op.qty > 0:
            new_assets_with_zero_qty.append(op)
    return new_assets_with_zero_qty


def part_assets():
    for operation in remove_asset_zero_qty(operations):
        if str(operation.asset).endswith('11'):
            fii.append(operation)
        else:
            assets.append(operation)


def perform_from_cei():
    print('READING XLS')
    for file in listdir('files_cei'):
        file_path = 'files_cei/' + file
        if path.isfile(file_path):
            read_cei_excel_file(file_path)

    print('SORTING')
    sort_assets(operations)

    print('REMOVING ASSET WITH ZERO AND PARTING ASSETS')
    part_assets()

    print('WRITING DATA')
    write_excel_file('cei_result.xls')

    print('FINISHED WITH SUCCESS')


def check_developments(date):
    developments = config_json['developments']

    for dev in developments:
        if dev['date'] == date:
            for op in operations:
                if op.asset == dev['asset']:
                    op.qty = op.qty * (dev['from'] * dev['to'])
                    op.price = op.price / (dev['from'] * dev['to'])
                    break


def add_operations_to_year(year):
    if assets_per_year.get(year) is None:
        assets_copy = copy.deepcopy(operations)
        sort_assets(assets_copy)
        assets_per_year[year] = remove_asset_zero_qty(assets_copy)


def perform_from_inter():
    print('READING XLS')

    config = config_json['inter_file']

    if config['cod_cli'] is None or config['cod_cli'] == '':
        print('MISSING KEY \"cod_cli\", CHECK configs.json')
    else:
        if config['read_file_by_file']:
            for file in listdir('files_inter'):
                file_path = 'files_inter/' + file
                if path.isfile(file_path):
                    read_inter_excel_file(file_path)
        else:
            for year in range(int(config['initial_year']), int(config['final_year']) + 1):
                my_datetime = datetime.datetime(year, 1, 1)
                while my_datetime.strftime('%Y') == str(year):
                    date = my_datetime.strftime('%m-%d-%Y')
                    check_developments(date)
                    str_date = my_datetime.strftime('%d%m%Y')
                    name_file = '_NotaCor_' + str_date + '_' + config['cod_cli'] + '.xls'
                    file_path = 'files_inter/' + name_file
                    if path.exists(file_path) and path.isfile(file_path):
                        read_inter_excel_file(file_path)
                    else:
                        file_path = file_path + 'x'
                        if path.exists(file_path) and path.isfile(file_path):
                            read_inter_excel_file(file_path)
                    my_datetime = my_datetime + timedelta(days=1)
                add_operations_to_year(year)

        print('SORTING')
        sort_assets(operations)

        print('REMOVING ASSET WITH ZERO AND PARTING ASSETS')
        part_assets()

        print('WRITING DATA')
        write_excel_file('inter_result.xls')

        print('FINISHED WITH SUCCESS')


def perform():
    if config_json['perform_inter']:
        perform_from_inter()

    if config_json['perform_cei']:
        perform_from_cei()


if __name__ == '__main__':
    file_json = open('configs.json')
    config_json = json.load(file_json)
    perform()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
