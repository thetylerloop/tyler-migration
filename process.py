#!/usr/bin/env python

import os

from agate import csv


TEXAS_STATE_FIPS = '48'
SMITH_COUNTY_FIPS = '423'
FORMAT3_DIR = 'data/format3'
OUTPUT_PATH = 'output/all_years.csv'

OUTPUT_FIELDNAMES = [
    'year',
    'in_state_fips',
    'in_county_fips',
    'in_state',
    'in_county',
    'returns',
    'exemptions',
    'agi'
]


def main():
    output = []

    for filename in os.listdir(FORMAT3_DIR):
        output.extend(parse_format3(filename))

    with open(OUTPUT_PATH, 'w') as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_FIELDNAMES)
        writer.writeheader()
        writer.writerows(output)


def parse_format3(filename):
    path = os.path.join(FORMAT3_DIR, filename)
    year = '20{}'.format(filename[-6:-4])

    print('Parsing', year)

    output = []

    with open(path, encoding='latin1') as f:
        reader = csv.DictReader(f)
        
        for row in reader:
            if row['y2_statefips'] == TEXAS_STATE_FIPS and row['y2_countyfips'] == SMITH_COUNTY_FIPS:
                output.append({
                    'year': year,
                    'in_state_fips': row['y1_statefips'],
                    'in_county_fips': row['y1_countyfips'],
                    'in_state': row['y1_state'],
                    'in_county': row['y1_countyname'],
                    'returns': row['n1'],
                    'exemptions': row['n2'],
                    'agi': row['agi']
                })

    return output





if __name__ == '__main__':
    main()