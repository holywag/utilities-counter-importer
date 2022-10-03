#!/usr/bin/env python3

import re
from datetime import datetime, date
import google_cloud.oauth as oauth
import google_cloud.sheets as gsheets

class RowFormatter:
    def __init__(self, service_name):
        self.service_name = service_name

    def make_range_ref(self, row_idx):
        return f'{self.service_name}!{self.RANGE_TEMPLATE.format(row_idx=row_idx)}'

    def make_range(self, range_values, **format_options):
        if self.get_row_date(range_values[-1]) == format_options['date']:
            # Overwrite existing row
            target_row_idx = len(range_values)
            last_row = range_values[-1]
        else:
            # Append a new row
            target_row_idx = len(range_values) + 1
            last_row = list(map(RowFormatter.__shift_row_refs, range_values[-1]))
            
        return (self.make_range_ref(target_row_idx), [self.make_row(last_row, **format_options)])

    def get_row_date(self, row):
        return datetime.strptime(row[0], UtilitiesCounterImporter.DATE_FORMAT).date()

    def __shift_row_refs(formula):
        if not type(formula) is str:
            return formula
        r = re.compile('(?<=[A-Z])\d+')
        start_idx = 0
        match = r.search(formula[start_idx:])
        while match:
            i,j = match.span()
            new_ref = str(int(match.group()) + 1)
            formula = formula[:start_idx+i] + new_ref + formula[start_idx+i+len(new_ref):]
            start_idx = start_idx + i + len(new_ref)
            match = r.search(formula[start_idx:])
        return formula

class RowDefaultFormatter(RowFormatter):
    RANGE_TEMPLATE = 'C{row_idx}:H{row_idx}'

    def make_row(self, last_row, **format_args):
        return [
            format_args['date'].strftime(UtilitiesCounterImporter.DATE_FORMAT),
            format_args['counter_value'] or last_row[1],
            format_args['tariff'] or last_row[2],
            last_row[3],
            0,
            last_row[5]
        ]

class RowRentFormatter(RowFormatter):
    RANGE_TEMPLATE = 'B{row_idx}:H{row_idx}'
    EX_RATE_LINK = '=HYPERLINK("https://minfin.com.ua/ua/currency/auction/archive/usd/ivano-frankovsk/#fromDate={date_arg}&toDate={date_arg})", {ex_rate})'
    EX_RATE_LINK_RE = re.compile(r'(?:.*\D)(\d+\.\d+)(?:\))$')
    RENT_START = 18
    RENT_END = 17

    def get_row_date(self, row):
        return datetime.strptime(row[2], UtilitiesCounterImporter.DATE_FORMAT).date()

    def make_row(self, last_row, **format_args):
        date_arg = format_args['date']
        period_start = date_arg.replace(day=self.RENT_START)
        if period_start.month < 12:
            period_end = period_start.replace(month=period_start.month+1, day=self.RENT_END)
        else:
            period_end = period_start.replace(year=period_start.year+1, month=1, day=self.RENT_END)
        ex_rate_link = self.EX_RATE_LINK.format(
            date_arg=date_arg.strftime('%d-%m-%Y'),
            ex_rate=format_args['tariff'] or self.EX_RATE_LINK_RE.match(last_row[3]).groups()[0]
        )
        return [
            period_start.strftime(UtilitiesCounterImporter.DATE_FORMAT),
            period_end.strftime(UtilitiesCounterImporter.DATE_FORMAT),
            date_arg.strftime(UtilitiesCounterImporter.DATE_FORMAT),
            ex_rate_link,
            last_row[4],
            0,
            last_row[6]
        ]

class RowWithPreferentialTariffFormatter(RowFormatter):
    RANGE_TEMPLATE = 'B{row_idx}:H{row_idx}'

    def make_row(self, last_row, **format_args):
        return [
            format_args['date'].strftime(UtilitiesCounterImporter.DATE_FORMAT),
            format_args['counter_value'] or last_row[1],
            format_args['tariff_preferential'] or last_row[2],
            format_args['tariff'] or last_row[3],
            last_row[4],
            0,
            last_row[6]
        ]

class UtilitiesCounterImporter:
    DATE_FORMAT = '%d.%m.%Y'
    ROW_FORMATTERS = {
        'electric': RowWithPreferentialTariffFormatter,
        'gas':      RowDefaultFormatter,
        'water':    RowDefaultFormatter,
        'rent':     RowRentFormatter
    }

    def __init__(self, google_oauth_creds_file_path, file_id):
        google_creds = oauth.GoogleOAuth(google_oauth_creds_file_path).authenticate(oauth.GoogleOAuthScopes.SHEETS)
        self.sheets_api = gsheets.GoogleSheetsApi(google_creds)
        self.file_id = file_id

    def __get_row_formatter(service_name):
        return UtilitiesCounterImporter.ROW_FORMATTERS[service_name](service_name)

    def add_record(self, service_name, date, counter_value=None, tariff=None, tariff_preferential=None):
        format_args = locals()
        del format_args['self']
        formatter = UtilitiesCounterImporter.__get_row_formatter(service_name)
        range_values = self.sheets_api.get_range(
            self.file_id, formatter.make_range_ref(''), gsheets.ValueRenderOption.FORMULA)
        return self.sheets_api.update_range(
            self.file_id, *formatter.make_range(range_values, **format_args), gsheets.ValueInputOption.USER_ENTERED, include_values_in_response=True)
