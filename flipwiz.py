# Python script to organize PixelPOS "Profit by Summary Group" sales reports
# tidied up by saleswiz.py into single tables for chart production.

import os
import csv
import sys

from glob import glob
from collections import OrderedDict
from copy import deepcopy


# *sales_clean.csv columns.
cols = ['Category', 'Subcategory', 'Item', 'Quantity', 'Value',]

def flip(out_base, files):
    """Consolidate monthly or weekly sales_clean.csv files into single table."""
    dl = []
    q = OrderedDict([(c, []) for c in cols[:3]])

    for i, f in enumerate(files):
        dl.append(OrderedDict())
        for c in cols:
            dl[i][c] = []

        try:
            with open(f, 'rb') as in_f:
                in_d = csv.DictReader(in_f, fieldnames=cols, delimiter=',', quotechar='"')
                for row in in_d:
                    [dl[i][c].append(row[c].strip()) for c in cols]
                in_f.close()
        except:
            print 'Problem reading file "%s".' % f.split('/')[-1]
            return False

        period = f.split('/')[-1].split('sales')[0]
        q[period] = []

    for r in xrange(1, len(dl[0][cols[0]])):
        [q[c].append(dl[0][c][r]) for c in cols[:3]]

    v = deepcopy(q)

    for i, d in enumerate(dl):
        for r in xrange(1, len(dl[0][cols[0]])):
            if d['Quantity'][r]:
                q[q.keys()[i+3]].append(d['Quantity'][r])
                v[v.keys()[i+3]].append(d['Value'][r])
            else:
                q[q.keys()[i+3]].append('')
                v[v.keys()[i+3]].append('')

    ok_q = write_file(q, out_base + 'quantity.csv', q.keys())
    ok_v = write_file(v, out_base + 'value.csv', v.keys())

    return ok_q and ok_v


def write_file(d, out_file, columns):
    """Write output csv file."""
    try:
        with open(out_file, 'wb') as out_f:
            out_d = csv.DictWriter(out_f, fieldnames=OrderedDict([(c, None) for c in columns]))
            out_d.writeheader()
            for r in xrange(len(d[columns[0]])):
                out_d.writerow(OrderedDict([(c, d[c][r]) for c in columns]))
            out_f.close()
        return True
    except:
        print 'Problem writing file "%s".' % out_file.split('/')[-1]
        return False


def main():
    try:
        freq = sys.argv[1]
    except:
        freq = ''

    freq = freq.strip()[0].lower() if freq else ''
    if freq not in ['m', 'w', 'b',]:
        print 'Specify monthly, weekly or both.'
        return

    reports_dir_win = 'C:\\Users\\Malky\\Documents\\BD\\Reports'
    reports_dir_py = reports_dir_win.decode('utf-8').replace('\\','/')

    monthly = reports_dir_py + '/monthly_'
    weekly = reports_dir_py + '/weekly_'

    clean_files = glob(reports_dir_py + '/' + '*sales_clean.csv')
    m_files = [f.decode('utf-8').replace('\\','/') for f in clean_files if 'we' not in f]
    w_files = [f.decode('utf-8').replace('\\','/') for f in clean_files if 'we' in f]

    if freq == 'm':
        ok_m = flip(monthly, m_files)
        print 'Monthly File OK.' if ok_m else 'Monthly File No Bueno.'
    elif freq == 'w':
        ok_w = flip(weekly, w_files)
        print 'Weekly File OK.' if ok_w else 'Weekly File No Bueno.'
    else:
        ok_m = flip(monthly, m_files)
        print 'Monthly File OK.' if ok_m else 'Monthly File No Bueno.'
        ok_w = flip(weekly, w_files)
        print 'Weekly File OK.' if ok_w else 'Weekly File No Bueno.'


if __name__ == '__main__':
    main()
