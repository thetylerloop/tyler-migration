#!/usr/bin/env python

from glob import glob
import os
import re
from zipfile import ZipFile

from agate import csv
import xlrd


TEXAS_STATE_FIPS = '48'
SMITH_COUNTY_FIPS = '423'

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
    'Smith County Tot Mig-Diff S': 'Smith County Total Migration-Different State',
    'Smith County Tot Mig-Diff St': 'Smith County Total Migration-Different State',
    'Smith County Tot Mig-Foreig': 'Smith County Total Migration-Foreign',
    'Smith County Tot Mig-Foreign': 'Smith County Total Migration-Foreign',
    'Smith County Tot Mig-Same S': 'Smith County Total Migration-Same State',
    'Smith County Tot Mig-Same St': 'Smith County Total Migration-Same State',
    'Smith County Tot Mig-US': 'Smith County Total Migration-US',
    'Smith County Tot Mig-US & F': 'Smith County Total Migration-US and Foreign',
    'Smith County Tot Mig-US & For': 'Smith County Total Migration-US and Foreign',
    'Smith County Non-migrants': 'Smith County Non-Migrants',
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
            if str(row[0]) == TEXAS_STATE_FIPS and str(row[1]) == SMITH_COUNTY_FIPS:
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

            if state_fips == TEXAS_STATE_FIPS and county_fips == SMITH_COUNTY_FIPS:
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
            if row['y2_statefips'] == TEXAS_STATE_FIPS and row['y2_countyfips'] == SMITH_COUNTY_FIPS:
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