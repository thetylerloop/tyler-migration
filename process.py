#!/usr/bin/env python

from glob import glob
import os
import re
from zipfile import ZipFile

from agate import csv
import xlrd


STATE_FIPS = '48'
COUNTY_FIPS = '423'
COUNTY_NAME = 'Smith County'

FORMAT1_DIR = 'data/format1'
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
    'agi_thousands'
]

COUNTY_NORMALIZATION = {
    f'{COUNTY_NAME} Tot Mig-Diff S': f'{COUNTY_NAME} Total Migration-Different State',
    f'{COUNTY_NAME} Tot Mig-Diff St': f'{COUNTY_NAME} Total Migration-Different State',
    f'{COUNTY_NAME} Tot Mig-Foreig': f'{COUNTY_NAME} Total Migration-Foreign',
    f'{COUNTY_NAME} Tot Mig-Foreign': f'{COUNTY_NAME} Total Migration-Foreign',
    f'{COUNTY_NAME} Tot Mig-Same S': f'{COUNTY_NAME} Total Migration-Same State',
    f'{COUNTY_NAME} Tot Mig-Same St': f'{COUNTY_NAME} Total Migration-Same State',
    f'{COUNTY_NAME} Tot Mig-US': f'{COUNTY_NAME} Total Migration-US',
    f'{COUNTY_NAME} Tot Mig-US & F': f'{COUNTY_NAME} Total Migration-US and Foreign',
    f'{COUNTY_NAME} Tot Mig-US & For': f'{COUNTY_NAME} Total Migration-US and Foreign',
    f'{COUNTY_NAME} Non-migrants': f'{COUNTY_NAME} Non-Migrants',
    'Other Flows - Diff State': 'Other flows - Different State',
    'East Baton Rouge Par': 'East Baton Rouge Parish',
    'San Bernardino Count': 'San Bernardino County'
}


def main():
    output = []

    for path in sorted(glob('{}/*.zip'.format(FORMAT1_DIR))):
        output.extend(parse_format1(path))

    for path in sorted(glob('{}/*.zip'.format(FORMAT2_DIR))):
        output.extend(parse_format2(path))

    for path in sorted(glob('{}/*.csv'.format(FORMAT3_DIR))):
        output.extend(parse_format3(path))

    output = sorted(output, key=lambda row: (row['year2'], row['year1_state_fips'], row['year1_county_fips']))

    with open(OUTPUT_PATH, 'w') as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_FIELDNAMES)
        writer.writeheader()
        writer.writerows(output)


def parse_format1(path):
    filename = os.path.basename(path)
    
    year1 = filename[0:4]
    year2 = filename[6:10]

    print('Parsing', year2)

    output = []

    archive = ZipFile(path, 'r')

    filename_formats = [
        '{}to{}CountyMigration/{}to{}CountyMigrationInflow/C{}{}Txi.xls'.format(
            year1, year2, year1, year2, year1[-2:], year2[-2:]),
        '{}to{}CountyMigration/{}to{}CountyMigrationInflow/co{}{}txi.xls'.format(
            year1, year2, year1, year2, year1[-2:], year2[-1:]),
        '{}to{}CountyMigration/{}to{}CountyMigrationInflow/co{}{}txir.xls'.format(
            year1, year2, year1, year2, year1[-2:], year2[-1:]),
        '{}to{}CountyMigration/{}to{}CountyMigrationInflow/co{}{}txi.xls'.format(
            year1, year2, year1, year2, year1[-1:], year2[-2:]),
        '{}to{}CountyMigration/{}to{}CountyMigrationInflow/co{}{}TXi.xls'.format(
            year1, year2, year1, year2, year1[-2:], year2[-2:])
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
            if str(row[0]) == STATE_FIPS and str(row[1]) == COUNTY_FIPS:
                output.append({
                    'year2': year2,
                    'year1_state_fips': str(row[2]),
                    'year1_county_fips': str(row[3]),
                    'year1_state': row[4],
                    'year1_county': COUNTY_NORMALIZATION.get(row[5], row[5]),
                    'returns': str(row[6]),
                    'exemptions': str(row[7]),
                    'agi_thousands': str(row[8])
                })

    return output


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
            try:
                state_fips = str(int(row[0]))
                county_fips = str(int(row[1]))
            except ValueError:
                continue

            if state_fips == STATE_FIPS and county_fips == COUNTY_FIPS:
                output.append({
                    'year2': year,
                    'year1_state_fips': str(row[2]),
                    'year1_county_fips': str(row[3]),
                    'year1_state': row[4],
                    'year1_county': COUNTY_NORMALIZATION.get(row[5], row[5]),
                    'returns': str(row[6]),
                    'exemptions': str(row[7]),
                    'agi_thousands': str(row[8])
                })

    return output


def parse_format3(path):
    year = '20{}'.format(path[-6:-4])

    print('Parsing', year)

    output = []

    with open(path, encoding='latin1') as f:
        reader = csv.DictReader(f)
        
        for row in reader:
            if row['y2_statefips'] == STATE_FIPS and row['y2_countyfips'] == COUNTY_FIPS:
                output.append({
                    'year2': year,
                    'year1_state_fips': int(row['y1_statefips']),
                    'year1_county_fips': int(row['y1_countyfips']),
                    'year1_state': row['y1_state'],
                    'year1_county': COUNTY_NORMALIZATION.get(row['y1_countyname'], row['y1_countyname']),
                    'returns': row['n1'],
                    'exemptions': row['n2'],
                    'agi_thousands': row['agi']
                })

    return output


if __name__ == '__main__':
    main()