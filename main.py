import xlrd
import my_asset_class as Asset

operations = []


def add_new_operation(new_op, operation_type):
    for op in operations:
        if new_op.asset == op.asset:
            if operation_type == 'COMPRADA':
                new_qnt = op.quant + new_op.quant
                new_price = ((op.quant * op.price) + (new_op.quant * new_op.price)) / new_qnt
                op.quant = new_qnt
                op.price = float(int(new_price*1000)/1000)
                return
            elif operation_type == 'VENDIDA':
                new_qnt = op.quant - new_op.quant
                op.quant = new_qnt
                return
    operations.append(new_op)


def perform():
    file_b3 = xlrd.open_workbook('files/compras_b3.xls')
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
                asset = str(sheet.cell(row, 3).value)
                data_compra = str(sheet.cell(row, 10))
                qnt_buy = int(sheet.cell(row, 18).value)
                qnt_sell = int(sheet.cell(row, 24).value)
                average_buy_prince = str(sheet.cell(row, 34).value)
                average_sell_prince = str(sheet.cell(row, 43).value)
                qnt_liquida = str(sheet.cell(row, 49).value)
                operation_type = str(sheet.cell(row, 54).value)

                qnt = qnt_sell if operation_type == 'VENDIDA' else qnt_buy
                prince = average_sell_prince if operation_type == 'VENDIDA' else average_buy_prince

                op = Asset.MyAsset(asset=asset, quant=qnt, price=prince)
                add_new_operation(op, operation_type)
            else:
                break


def asset(op):
    return op.asset


if __name__ == '__main__':
    perform()
    operations.sort(key=asset)
    for operation in operations:
        operation.print()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
