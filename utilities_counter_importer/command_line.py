#!/usr/bin/env python3

import argparse
from datetime import datetime, date
from utilities_counter_importer import UtilitiesCounterImporter

def main():
    parser = argparse.ArgumentParser(
        prog='utilities_counter_importer',
        description='Tool that imports utilities counters to a Google Sheets document.')

    parser.add_argument('spreadsheet_id', help='ID of a Google Sheets document where utilities counters are to be imported to.')
    parser.add_argument('service_name', choices=UtilitiesCounterImporter.ROW_FORMATTERS.keys(), help='Name of the utilitiy service.')
    parser.add_argument('--counter_value', help='Counter value as of report date. Use previous value by default.', type=int)
    parser.add_argument('--credentials', metavar='FILE_PATH', help='Path to a JSON file containing Google Cloud app credentials. Default is ./credentials.json', default='./credentials.json')
    parser.add_argument('--date', metavar='DD.MM.YYYY', help=f'Report date in the format {UtilitiesCounterImporter.DATE_FORMAT.replace("%", "%%")}, today by default', type=lambda v: datetime.strptime(v, UtilitiesCounterImporter.DATE_FORMAT).date())
    parser.add_argument('--tariff', metavar='TARIFF', help='Main tariff', type=float)
    parser.add_argument('--tariff_preferential', metavar='TARIFF', help='Service-specific preferential tariff', type=float)

    args = parser.parse_args()

    importer = UtilitiesCounterImporter(args.credentials, args.spreadsheet_id)
    result = importer.add_record(args.service_name, args.date or date.today(), args.counter_value, args.tariff, args.tariff_preferential)
    print(result)
