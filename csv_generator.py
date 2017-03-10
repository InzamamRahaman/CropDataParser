# Name: Inzamam Kaleem Rahaman
# Date: 27th February 2017
# File description: command line application for the generation of csv files
#                   containing the data contained in the supplied NAMDEVCO xls files


import argparse
import glob
import os
from pyexcel_xls import get_data
from datetime import datetime
from collections import defaultdict
from dateutil.parser import parse

parser = argparse.ArgumentParser()
parser.add_argument('--input', help='The input directory to search')
parser.add_argument('--output', help='The output directory to search')
args = parser.parse_args()


def get_stating_point(sheet):
    """
    Finds the point in the sheet where the data actually begins
    :param sheet: the sheet as a list of lists
    :return: the index where the data actually begins
    """
    data = sheet
    start_point = 0
    for row_id, row_data in enumerate(sheet):
        if len(row_data) > 0 and data[row_id][0] == 'Commodity':
            return start_point + 1
        start_point += 1
    return -1


def is_all_upper(string):
    """
    Determines if all the characters in a string are uppercase
    :param string: the string to be considered
    :return: True if all characters are uppercase, False otherwise
    """
    for s in string:
        if s.islower():
            return False
    return True

def clean_headings(sheet):
    """
    Removes blank rows and headings such as the rows with "ROOT CROPS"
    :param sheet: the sheet with the formatting removed
    :return: rows of the sheet as lists with heading rows removed
    """
    return filter(lambda x: len(x) > 6 and not is_all_upper(x[0]), sheet)


def clean_sheet(sheet):
    """
    Cleans the sheet to yield a list of lists containing only crop data
    :param sheet: the sheet to be cleaned
    :return: the cleaned sheet as a list of lists
    """
    starting_point = get_stating_point(sheet)
    if starting_point == -1:
        print('Error!')
        return None
    temp = sheet[starting_point:]
    return list(clean_headings(temp))

def get_price_data_from_row(row):
    crop = row[0]
    unit = row[1]
    price = row[6]
    price_row = [crop, unit, price]
    return price_row

def get_volume_data_from_row(row):
    crop = row[0]
    unit = row[1]
    volume = row[3]
    volume_row = [crop, unit, volume]
    return volume_row

def get_prices(rows):
    prices = map(get_price_data_from_row, rows)
    return prices

def get_volumes(rows):
    volumes = map(get_volume_data_from_row, rows)
    return volumes

def generate_csv_string(headings, rows):
    contents = ','.join(headings)
    row_contents = map(lambda x: ','.join(map(str, x)), rows)
    body = '\n'.join(row_contents)
    contents = contents + '\n' + body
    return contents

def generate_prices_csv_content(prices):
    return generate_csv_string(['Crop', 'Unit', 'Price'], prices)

def generate_volume_csv_content(volumes):
    return generate_csv_string(['Crop', 'Unit', 'Volume'], volumes)


def sheet_date(sheet):
    starting_point = get_stating_point(sheet)
    data_row = sheet[starting_point - 1]
    date_string = data_row[3]
    date = date_string
    if type(date_string) == type(''):
        date_format1 = '%d/%m/%Y'
        date_format2 = "%d/%m/%y"
        try:
            date = datetime.strptime(date_string, date_format1)
        except:
            date = datetime.strptime(date_string, date_format2)
    return date



file_extension = 'xls'
file_regex = '*.{0}'.format(file_extension)
xls_files = glob.glob(os.path.join(args.input, file_regex))

crop_table = defaultdict(list)

def ensure_datetime(d):
    """
    Takes a date or a datetime as input, outputs a datetime
    """
    if isinstance(d, datetime):
        return d
    return datetime(d.year, d.month, d.day)


def update_crop_table(crop_table, date, rows):
    price_data_loc = 6
    volume_data_loc = 3
    crop_name_loc = 0
    unit_loc = 1
    for row in rows:
        crop = row[crop_name_loc]
        unit = row[unit_loc]
        price = row[price_data_loc]
        volume = row[volume_data_loc]
        date = ensure_datetime(date)
        crop_data = [date, unit, price, volume]
        crop_table[crop].append(crop_data)
    return crop_table


prices_dir = os.path.join(args.output, 'prices')
volume_dir = os.path.join(args.output, 'volume')
crop_dir = os.path.join(args.output, 'crops')
os.makedirs(prices_dir, exist_ok=True)
os.makedirs(volume_dir, exist_ok=True)
os.makedirs(crop_dir, exist_ok=True)

for xls_file in xls_files:
    filename = os.path.basename(xls_file)
    print('Processing ', filename)
    withoutextension = os.path.splitext(filename)[0]
    output_filename = '{0}.csv'.format(withoutextension)
    prices_filename = os.path.join(prices_dir, output_filename)
    volume_filename = os.path.join(volume_dir, output_filename)
    spreadsheet = get_data(xls_file)
    sheet = spreadsheet['Sheet1']
    cleaned_rows = clean_sheet(sheet)
    date = sheet_date(sheet)
    crop_table = update_crop_table(crop_table, date, cleaned_rows)
    prices = get_prices(cleaned_rows)
    volumes = get_volumes(cleaned_rows)
    price_contents = generate_prices_csv_content(prices)
    volume_contents = generate_volume_csv_content(volumes)
    with open(prices_filename, 'w') as outf:
        outf.write(price_contents)
    with open(volume_filename, 'w') as outf:
        outf.write(volume_contents)


for crop, table in crop_table.items():
    filename = '{0}.csv'.format(crop)
    filepath = os.path.join(crop_dir, filename)
    table = sorted(table, key=lambda x: x[0])
    representation = generate_csv_string(['date', 'unit', 'price', 'volume'], table)
    with open(filepath, 'w') as outf:
        outf.write(representation)








