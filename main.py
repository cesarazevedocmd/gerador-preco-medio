import xlrd
import xlwt
import json
import datetime
from datetime import timedelta
import my_asset_class as Asset
import os.path as path
from os import listdir

operations = []
operations_without_zero_qty = []


def calculate_average_price(op, new_op):
    new_qty = op.qty + new_op.qty
    if new_qty != 0:
        return ((op.qty * op.price) + (new_op.qty * new_op.price)) / new_qty
    return 0


def add_new_operation(new_op, operation_type):
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

            if value != '':
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


def asset(op):
    return op.asset


def write_excel_file(name_result_file):
    xls = xlwt.Workbook()
    sheet = xls.add_sheet('ASSETS')

    index = 0
    sheet.write(index, 0, 'ASSET')
    sheet.write(index, 1, 'QTY')
    sheet.write(index, 2, 'PRICE')

    for operation in operations_without_zero_qty:
        index += 1
        sheet.write(index, 0, operation.asset)
        sheet.write(index, 1, operation.qty)
        sheet.write(index, 2, operation.price)

    xls.save(name_result_file)


def remove_asset_zero_qty():
    for op in operations:
        if op.qty > 0:
            operations_without_zero_qty.append(op)


def perform_from_cei():
    print('READING XLS')
    for file in listdir('files_cei'):
        file_path = 'files_cei/' + file
        if path.isfile(file_path):
            read_cei_excel_file(file_path)

    print('SORTING')
    operations.sort(key=asset)

    print('REMOVING ASSET WITH ZERO')
    remove_asset_zero_qty()

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
            for ano in range(int(config['initial_year']), int(config['final_year']) + 1):
                my_datetime = datetime.datetime(ano, 1, 1)
                while my_datetime.strftime('%Y') == str(ano):
                    date = my_datetime.strftime('%m-%d-%Y')
                    check_developments(date)
                    str_date = my_datetime.strftime('%d%m%Y')
                    name_file = '_NotaCor_' + str_date + '_'+config['cod_cli']+'.xls'
                    file_path = 'files_inter/' + name_file
                    if path.exists(file_path) and path.isfile(file_path):
                        read_inter_excel_file(file_path)
                    my_datetime = my_datetime + timedelta(days=1)

        print('SORTING')
        operations.sort(key=asset)

        print('REMOVING ASSET WITH ZERO')
        remove_asset_zero_qty()

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
