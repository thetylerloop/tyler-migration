#!/usr/bin/env python

from glob import glob
import os
from zipfile import ZipFile

from agate import csv
import xlrd


TEXAS_STATE_FIPS = '48'
SMITH_COUNTY_FIPS = '423'

FORMAT2_DIR = 'data/format2'
FORMAT3_DIR = 'data/format3'

OUTPUT_PATH = 'output/all_years.csv'
OUTPUT_FIELDNAMES = [
    'year2',
    'year1_state_fips',
    'year1_county_fips',
    'year1_state',
    'year1_county',
    'returns',
    'exemptions',
    'agi'
]


def main():
    output = []

    for path in sorted(glob('{}/*.zip'.format(FORMAT2_DIR))):
        output.extend(parse_format2(path))

    for path in sorted(glob('{}/*.csv'.format(FORMAT3_DIR))):
        output.extend(parse_format3(path))

    output = sorted(output, key=lambda row: (row['year2'], row['year1_state_fips'], row['year1_county_fips']))

    with open(OUTPUT_PATH, 'w') as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_FIELDNAMES)
        writer.writeheader()
        writer.writerows(output)

    


def parse_format2(path):
    year = '20{}'.format(path[-6:-4])

    print('Parsing', year)

    output = []

    archive = ZipFile(path, 'r')

    slug = path[-8:-4]

    filename_formats = [
        'co{}iTx.xls'.format(slug),
        'co{}TXi.xls'.format(slug),
        'co{}Txi.xls'.format(slug),
        'co{}iTX.xls'.format(slug),
        'co{}itx.xls'.format(slug)
    ]

    archive_filename = None

    for filename in filename_formats:
        if filename in archive.namelist():
            archive_filename = filename

    with archive.open(archive_filename, 'r') as f:
        book = xlrd.open_workbook(file_contents=f.read())

        sheet = book.sheet_by_index(0)

        columns = []

        for i in range(sheet.ncols):
            columns.append(sheet.col_values(i)[8:])

        output = []

        for row in zip(*columns):
            output.append({
                'year2': year,
                'year1_state_fips': str(row[2]),
                'year1_county_fips': str(row[3]),
                'year1_state': row[4],
                'year1_county': row[5],
                'returns': str(row[6]),
                'exemptions': str(row[7]),
                'agi': str(row[8])
            })

    return output


def parse_format3(path):
    year = '20{}'.format(path[-6:-4])

    print('Parsing', year)

    output = []

    with open(path, encoding='latin1') as f:
        reader = csv.DictReader(f)
        
        for row in reader:
            if row['y2_statefips'] == TEXAS_STATE_FIPS and row['y2_countyfips'] == SMITH_COUNTY_FIPS:
                output.append({
                    'year2': year,
                    'year1_state_fips': int(row['y1_statefips']),
                    'year1_county_fips': int(row['y1_countyfips']),
                    'year1_state': row['y1_state'],
                    'year1_county': row['y1_countyname'],
                    'returns': row['n1'],
                    'exemptions': row['n2'],
                    'agi': row['agi']
                })

    return output


if __name__ == '__main__':
    main()