import argparse
import json
import re

import xlrd
import xlsxwriter

PRODUCT_COLUMN_NAME="My Products: Products"
COLUMNS="ABCDEFGHIJKLMNOPQRSTUVWXYZ"

def parse_products(s):
    items = []
    total = None
    for chunk in s.split('\n'):
        t = re.search('^Total: (\S+)$', chunk.strip())
        if t is not None:
            total = t.group(1)
            continue
        m = re.search('^(?P<name>.+) \(Amount: (?P<unit_price>\S+) USD, Quantity: (?P<quantity>[0-9]+)\)$', chunk.strip())
        items.append({ 'name' : m.group('name'),
                       'unit_price' : float(m.group('unit_price')),
                       'quantity' : int(m.group('quantity')) })
    return items, total

def load(src):
    book = xlrd.open_workbook(src)
    sheet = book.sheet_by_index(0)
    columns = []
    for x in xrange(sheet.ncols):
        columns.append(sheet.cell_value(rowx=0, colx=x))
    rows = []
    for y in xrange(1, sheet.nrows):
        row = {}
        for x in xrange(sheet.ncols):
            row[columns[x]] = sheet.cell_value(rowx=y, colx=x)
        rows.append(row)

    products = []
    for row in rows:
        items, total = parse_products(row[PRODUCT_COLUMN_NAME])
        for item in items:
            if item['name'] not in products:
                products.append(item['name'])

    pcol = columns.index(PRODUCT_COLUMN_NAME)
    cols_out = columns[:pcol] + products + ['Total'] + columns[pcol:]

    for row in rows:
        items, total = parse_products(row[PRODUCT_COLUMN_NAME])
        row['Total'] = total
        for item in items:
           row[item['name']] = item['quantity']

    return cols_out, rows, pcol

def save(outfile, cols, rows, pcol):
    workbook = xlsxwriter.Workbook(outfile)
    worksheet = workbook.add_worksheet()

    # write header row
    header_format = workbook.add_format()
    header_format.set_bold()
    header_format.set_bg_color('#90EE90')
    header_format.set_pattern(1)
    header_format.set_text_wrap()
    header_format.set_align('center')
    header_format.set_align('vcenter')
    header_format.set_border(1)
    cidx = 0
    for col in cols:
        worksheet.write(0, cidx, col, header_format)
        cidx += 1

    money = workbook.add_format()
    money.set_num_format(0x07)
    money.set_border(1)

    default_fmt = workbook.add_format()
    default_fmt.set_border(1)
    
    # write data rows
    ridx = 1
    for row in rows:
        cidx = 0
        for col in cols:
            if col in row:
                if col == 'Total':
                    worksheet.write(ridx, cidx, float(row[col]), money)
                else:
                    worksheet.write(ridx, cidx, row[col], default_fmt)
            else:
                worksheet.write(ridx, cidx, 0, default_fmt)
            cidx += 1
        ridx += 1

    # write product item totals
    bold = workbook.add_format()
    bold.set_bold()
    bold.set_border(1)
    worksheet.write(ridx, pcol-1, 'Total', bold)

    num_products = cols.index('Total') - pcol
    for i in xrange(num_products):
        worksheet.write(ridx, pcol+i, "=SUM(%s%d:%s%d)" %
                        (COLUMNS[pcol+i], 2, COLUMNS[pcol+i], len(rows)+1),
                        bold)
    ridx += 1

    lem_price_fmt = workbook.add_format()
    lem_price_fmt.set_num_format(0x07)
    lem_price_fmt.set_pattern(1)
    lem_price_fmt.set_bg_color('yellow')
    lem_price_fmt.set_border(1)
    
    worksheet.write(ridx, pcol-1, 'Lem Prices', bold)
    for i in xrange(num_products):
        worksheet.write(ridx, pcol+i, 0, lem_price_fmt)
    ridx += 1

    worksheet.write(ridx, pcol-1, 'Total Cost', default_fmt)
    for i in xrange(num_products):
        worksheet.write(ridx, pcol+i, "=%s%d*%s%d" %
                        (COLUMNS[pcol+i], ridx-1, COLUMNS[pcol+i], ridx),
                        money)
    ridx += 1

    check_fmt = workbook.add_format()
    check_fmt.set_num_format(0x07)
    check_fmt.set_pattern(1)
    check_fmt.set_bg_color('#90EE90')
    check_fmt.set_border(1)

    worksheet.write(ridx, pcol-1, 'Check for Lem:', default_fmt)
    worksheet.write(ridx, pcol, "=SUM(%s%d:%s%d)" %
                    (COLUMNS[pcol], ridx, COLUMNS[pcol+num_products-1], ridx),
                    check_fmt)
    
    workbook.close()

def main():
    parser = argparse.ArgumentParser(description='Convert JotForm spreadsheet')
    parser.add_argument('srcfile', help='.xlsx from JotForm')
    parser.add_argument('dstfile', help='.xlsx to generate')
    args = parser.parse_args()
    cols, rows, pcol = load(args.srcfile)
    save(args.dstfile, cols, rows, pcol)
    
if __name__ == "__main__":
    main()
