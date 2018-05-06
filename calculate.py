#!/usr/bin/env python

from collections import defaultdict
import csv

ALL_YEARS_PATH = 'output/all_years.csv'
FIRST_YEAR = 1996
LAST_YEAR = 2016

LABELS = {
    'Smith County Non-Migrants': 'non_migrant',
    'Smith County Total Migration-Same State': 'in_state',
    'Smith County Total Migration-Different State': 'out_of_state',
    'Smith County Total Migration-Foreign': 'foreign'
}

RATES_PATH = 'output/rates.csv'


def main():
    year_label_exemptions = defaultdict(dict) 

    with open(ALL_YEARS_PATH) as f:
        reader = csv.DictReader(f)

        for row in reader:
            label = LABELS.get(row['year1_county'])

            if not label:
                continue

            year = row['year2']

            year_label_exemptions[year][label] = float(row['exemptions'])

    output = []

    for year in range(FIRST_YEAR, LAST_YEAR + 1):
        year = str(year)

        output.append({
            'year': year,
            'in_state': round(year_label_exemptions[year]['in_state'] / year_label_exemptions[year]['non_migrant'] * 100, 2),
            'out_of_state': round(year_label_exemptions[year]['out_of_state'] / year_label_exemptions[year]['non_migrant'] * 100, 2),
            'foreign': round(year_label_exemptions[year]['foreign'] / year_label_exemptions[year]['non_migrant'] * 100, 2),
        })

    with open(RATES_PATH, 'w') as f:
        writer = csv.DictWriter(f, fieldnames=['year', 'in_state', 'out_of_state', 'foreign'])
        writer.writeheader()

        for row in output:
            writer.writerow(row)


if __name__ == '__main__':
    main()