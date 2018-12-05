# coding: utf-8
"""Online Scout Manager Interface - generate report from subs.

Usage:
  gocardless.py [-d] [-u] [--year=<year>] [--month=<month>] [--section=<section>] <outdir>
  gocardless.py (-h | --help)
  gocardless.py --version


Options:
  <outdir>              Output directory for spreadsheets.
  -d,--debug            Turn on debug output.
  -u, --upload          Upload to Drive.
  --month=<month>       Which financial month to use [default: current].
  --year=<year>         Which year to use [default: current].
  --section=<section>]  Only output a specific section [default: All].
  -h,--help             Show this screen.
  --version             Show version.
"""

import decimal
import logging
import os.path
import re
from datetime import date
from datetime import datetime

import gocardless_pro as gp

# import requests
# requests.__dummy__ = None
# import requests_cache
# requests_cache.install_cache('.gc_cache')

import numpy as np
import os
import pandas as pd
from dateutil.parser import parse
from dateutil.relativedelta import relativedelta
from dateutil.tz import tzutc
from docopt import docopt
from gc_accounts import SECTION_MAP, DRIVE_FOLDERS
from xlsxwriter.utility import xl_rowcol_to_cell
from gdrive_upload import upload

log = logging.getLogger(__name__)

MONTH_MAP = {1: "10", 2: "11", 3: "12", 4: "01", 5: "02", 6: "03", 7: "04", 8: "05", 9: "06", 10: "07", 11: "08",
             12: "09"}

from decimal import Decimal, getcontext, ROUND_HALF_UP

# GoCardless appear to apply rounding up.
getcontext().rounding = ROUND_HALF_UP

GOCARDLESS_FEE = Decimal('0.01')   # 1% per transaction
GOCARDLESS_LIMIT = Decimal('2.00') # Capped as £2.00
OSM_FEE = Decimal('0.0195')        # 1.95% per transaction
OSM_LIMIT = Decimal('Infinity')    # No cap.

THREE_PLACES = Decimal('0.001')  # For quantizing to 3 places.
TWO_PLACES = Decimal('0.01')     # For quantizing to 2 places.


def fee(gross_amount, percentage_fee, limit=Decimal('Infinity')):
    """Calculate the fee charged for a transaction.

    Return the lower of the calculated fee and *limit*.

    We have to quantize twice to get the same result as GoCardless. First we quantize to
    3 decimal places and then to 2 decimal places. This gets a different result to quantizing
    to 2 decimal places directly.
    """
    return min((gross_amount * percentage_fee).quantize(THREE_PLACES).quantize(TWO_PLACES),
               limit)


def transaction_fee(gross_amount):
    """Calculate the total fee charged by GoCardless and OSM on a single transaction.

    Note: this must be applied at a single GoCardless transaction level. It does not work
    if you apply it to the aggregated payment that you receive into your bank account
    because the rounding must be applied at end transaction."""
    return (fee(gross_amount, GOCARDLESS_FEE, GOCARDLESS_LIMIT) +
            fee(gross_amount, OSM_FEE, OSM_LIMIT))


def total_fee_at_bank(gross_amounts):
    """Return the total fee for a set of transactions as paid by GoCardless to the bank."""
    return sum([transaction_fee(_) for _ in gross_amounts])

def fetch_account(token, frm, to):
    client = gp.Client(access_token=token, environment='live')
    data = []
    for payout in client.payouts.all(params={"created_at[gt]": frm,
                                             "created_at[lte]": to}):
        for parent in client.events.all(params={"payout": payout.id}):
            for event in client.events.all(params={"parent_event": parent.id}):
                payment = client.payments.get(event.links.payment)
                mandate = client.mandates.get(payment.links.mandate)
                acc = client.customer_bank_accounts.get(mandate.links.customer_bank_account)
                customer = client.customers.get(acc.links.customer)
                data.append((parse(parent.created_at).strftime("%Y-%m-%d"),
                             payout.reference,
                             customer.family_name,
                             payment.description,
                             payment.status,
                             Decimal(payment.amount) / 100))
    if len(data) > 0:
        frame = pd.DataFrame(data,
                             columns=["payout_date",
                                      "payout_id",
                                      "customer_family_name",
                                      "payment_description",
                                      "payment_status",
                                      "payment_amount"])
        frame["payout_date"] = pd.to_datetime(frame["payout_date"])
        frame["payment_fee"] = frame["payment_amount"].apply(transaction_fee)
        frame["payment_net"] = frame["payment_amount"] - frame["payment_fee"]
        frame = frame.sort_values(by="payout_date", ascending=True)
        return frame
    return None


def export_section(token, name, directory, frm, to):
    frame = fetch_account(token, frm.isoformat(), to.isoformat())
    if frame is None:
        log.warn("No payments found this month for {}".format(name))
        return None

    cols = frame.columns.tolist()
    # Reorder the columns so that payment_net is first as it is easier to transcribe onto accounts.
    frame = frame[cols[:-3] + [cols[-1], cols[-3], cols[-2]]]

    if frame is not None:

        # Create the page that lists each transaction/description. This allows you to see what income type makes up
        # each payment to the bank account.
        frame2 = frame.drop(["customer_family_name", "payment_status"], axis=1)
        group = frame2.groupby(["payout_id", "payout_date", "payment_description"], as_index=False)
        group2 = group.aggregate(np.sum)
        group2 = group2.sort_values(by="payout_date", ascending=True)

        # Create the page the lists each transaction as it appears on the bank statement.
        frame3 = frame2.drop(["payment_description"], axis=1)
        group3 = frame3.groupby(["payout_date", "payout_id"], as_index=False)
        group4 = group3.aggregate(np.sum)
        group4 = group4.sort_values(by="payout_date", ascending=True)

        # Create pivot that has each event as a separate column.
        frame_cp = frame.copy()

        # Function to rewrite the column names to remove the schedule mame
        def _(s):
            subs = re.match('.*Subscriptions.*', s)
            if subs:
                return "Subs"
            m = re.match('(.*)\((.*)\)', s)
            return m.group(2) if m else s

        frame_cp.payment_description = frame_cp.payment_description.apply(_)

        frame_by_event = frame_cp.drop(["customer_family_name", "payment_status", "payment_amount", "payment_fee"], axis=1)
        frame_by_event.payment_net = frame_by_event.payment_net.astype(decimal.Decimal)
        pivot_by_event = frame_by_event.pivot_table(index=['payout_date', 'payout_id'],
                                                    columns='payment_description',
                                                    values="payment_net", aggfunc=np.sum)


        # Turn the index into real columns.
        pivot_by_event.reset_index(inplace=True)

        filename = os.path.join(directory, "{} {} {} {} {} GoCardless.xls".format(MONTH_MAP[to.month],
                                                                                  to.day - 1,
                                                                                  to.strftime("%b"),
                                                                                  to.year,
                                                                                  name))

        with pd.ExcelWriter(filename,
                            engine='xlsxwriter',
                            datetime_format='dd mmmm yyyy') as writer:
            group4.to_excel(writer, 'By Bank Transaction', index=False)
            group2.to_excel(writer, 'Banks Transaction Breakdown', index=False)
            pivot_by_event.to_excel(writer, 'By Event', index=False)
            frame.to_excel(writer, 'By Payment To GoCardless', index=False)
            workbook = writer.book
            bold = workbook.add_format({'bold': True})

            worksheet = writer.sheets['By Bank Transaction']
            worksheet.set_zoom(90)
            worksheet.set_column('A:A', 14)
            worksheet.set_column('B:B', 22)
            worksheet.set_column('C:E', 16)

            cols = len(group4.columns) - 1
            rows = len(group4)
            worksheet.autofilter(0, 0, rows, cols)
            worksheet.write(rows + 2, cols - 3, "Total", bold)
            for col in (cols - 2, cols - 1, cols):
                worksheet.write_formula(rows + 2, col,
                                        '=SUBTOTAL("109",{}:{})'.format(
                                            xl_rowcol_to_cell(1, col),
                                            xl_rowcol_to_cell(rows, col)),
                                        bold)

            worksheet = writer.sheets['Banks Transaction Breakdown']
            worksheet.set_zoom(90)
            worksheet.set_column('A:A', 22)
            worksheet.set_column('B:B', 12)
            worksheet.set_column('C:C', 50)
            worksheet.set_column('D:F', 16)
            cols = len(group2.columns) - 1
            rows = len(group2)
            worksheet.autofilter(0, 0, rows, cols)
            worksheet.write(rows + 2, cols - 3, "Total", bold)
            for col in (cols - 2, cols - 1, cols):
                worksheet.write_formula(rows + 2, col,
                                        '=SUBTOTAL("109",{}:{})'.format(
                                            xl_rowcol_to_cell(1, col),
                                            xl_rowcol_to_cell(rows, col)),
                                        bold)
            worksheet = writer.sheets['By Payment To GoCardless']
            worksheet.set_zoom(90)
            worksheet.set_column('A:A', 14)
            worksheet.set_column('B:B', 20)
            worksheet.set_column('C:C', 21)
            worksheet.set_column('D:D', 50)
            worksheet.set_column('E:E', 18)
            worksheet.set_column('F:H', 16)
            cols = len(frame.columns) - 1
            rows = len(frame)
            worksheet.autofilter(0, 0, rows, cols)
            worksheet.write(rows + 2, cols - 3, "Total", bold)
            for col in (cols - 2, cols - 1, cols):
                worksheet.write_formula(rows + 2, col,
                                        '=SUBTOTAL("109",{}:{})'.format(
                                            xl_rowcol_to_cell(1, col),
                                            xl_rowcol_to_cell(rows, col)),
                                        bold)

            worksheet = writer.sheets['By Event']
            worksheet.set_zoom(90)
            worksheet.set_column('A:A', 14)
            worksheet.set_column('B:B', 22)
            worksheet.set_column('C:K', 22)
            cols = len(pivot_by_event.columns) - 1
            rows = len(pivot_by_event)
            worksheet.autofilter(0, 0, rows, cols)
            worksheet.write(rows + 2, cols - 3, "Total", bold)
            for col in (cols - 2, cols - 1, cols):
                worksheet.write_formula(rows + 2, col,
                                        '=SUBTOTAL("109",{}:{})'.format(
                                            xl_rowcol_to_cell(1, col),
                                            xl_rowcol_to_cell(rows, col)),
                                        bold)

            return filename

    else:
        log.info("{} has not transactions".format(name))

    return None


if __name__ == '__main__':

    args = docopt(__doc__, version='OSM 2.0')

    if args['--debug']:
        level = logging.DEBUG
    else:
        level = logging.WARN

    if args['--month'] in [None, 'current']:
        args['--month'] = (date.today() - relativedelta(months=+1)).month
    else:
        args['--month'] = int(args['--month'])

    if args['--year'] in [None, 'current']:
        args['--year'] = (date.today() - relativedelta(months=+1)).year
    else:
        args['--year'] = int(args['--year'])

    if args['--section'] in [None, 'All']:
        args['--section'] = SECTION_MAP.keys()
    else:
        args['--section'] = [args['--section']]

    logging.basicConfig(level=level)
    log.debug("Debug On\n")
    import requests
    import http.client as http_client
    # http_client.HTTPConnection.debuglevel = 1
    requests_log = logging.getLogger("requests.packages.urllib3")
    requests_log.setLevel(logging.WARN)
    requests_log.propagate = True

    frm = datetime(args['--year'], args['--month'], 4, 0, 0, 0, tzinfo=tzutc())
    to = frm + relativedelta(months=+1)

    log.debug("from: {}, to: {}".format(frm.isoformat(), to.isoformat()))

    for section in args['--section']:
        token = SECTION_MAP[section]
        filename = export_section(token, section, args['<outdir>'], frm, to)

        if args['--upload']:
            if filename is not None:
                upload(filename, DRIVE_FOLDERS[section],
                       filename=os.path.splitext(os.path.split(filename)[1])[0])
