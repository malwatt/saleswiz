# Python script to organize PixelPOS "Profit by Summary Group" sales reports
# into a useful and consistent format.

import os
import glob
import csv

from collections import OrderedDict


beer_order = [('BOCK', 'DADDY BOCK'),]
beer_extra = [('SAMPLE', 'BEER SAMPLER'),]
beer_large = ('CHER', 'PITCHER')

wine_order = [('CHAR', 'PENFOLDS CHARDONNAY'),]
wine_extra = [('SPARK', 'MIONETTO SPARKING SPLIT'),]
wine_large = ('BOTT', 'BOTTLE')

burgers_order = [('CLASSIC', 'CLASSIC'),]
burgers_extra = [('VEG', 'VEGGIE'), ('BELLO', 'PORTABELLO'),]
burgers_large = ('1/2', '1/2 LB')

merchandise_order = [('HAT', 'HAT'), ('SHIRT', 'T-SHIRT'),]

keys = ['Category', 'Subcategory', 'Item', 'Col03', 'Col04', 'Col05', 'Col06',
        'Col07', 'Col08', 'Quantity', 'Col10', 'Col11', 'Col12', 'Col13',
        'Value', 'Col15', 'Cost', 'Col17', 'Col18', '% Cost', 'Col20', 'Col21',
        'Profit', 'Col23',]
cols = ['Category', 'Subcategory', 'Item', 'Quantity', 'Value',]
excludes = ['FIRE', 'Summary', 'Report', 'Description', 'Page', 'PRINTED',]


def parse_sales(sales_raw):

    d = {}
    d2 = {}
    beer = {}
    wine = {}
    burgers = {}
    merchandise = {}
    for c in cols:
        d[c] = []
        d2[c] = []
        beer[c] = []
        wine[c] = []
        burgers[c] = []
        merchandise[c] = []

    with open(sales_raw, 'rb') as in_f:
        in_d = csv.DictReader(in_f, fieldnames=keys, delimiter=',', quotechar='"')
        for row in in_d:
            [d[c].append(row[c].strip()) for c in cols]
        in_f.close()

    category = ''
    subcategory = ''
    for r in xrange(len(d[cols[0]])):
        row = [d[c][r] for c in cols]

        # Ignore empty rows.
        if not any(row):
            continue

        # Ignore rows with exclusion indicators.
        if any(e in ' '.join(row) for e in excludes):
            continue

        # Catch categories and ignore duplicates.
        if d['Category'][r] and not any(row[1:]):
            if category:
                continue
            else:
                category = d['Category'][r]

        # Catch subcategories and ignore duplicates.
        if d['Subcategory'][r] and not any(row[2:]):
            if subcategory:
                continue
            else:
                subcategory = d['Subcategory'][r]

        # Catch end of category.
        if d['Category'][r] and any(row[1:]):
            category = ''

        # Catch end of subcategory and organize if needed.
        if d['Subcategory'][r] and any(row[2:]):
            if subcategory == 'BEER':
                beer2 = organize_complex(beer, beer_order, beer_extra, beer_large)
                [d2[c].append(beer2[c][r2]) for r2 in xrange(len(beer2[cols[0]])) for c in cols]
            elif subcategory == 'WINE':
                wine2 = organize_complex(wine, wine_order, wine_extra, wine_large)
                [d2[c].append(wine2[c][r2]) for r2 in xrange(len(wine2[cols[0]])) for c in cols]
            elif subcategory == 'BURGERS':
                burgers2 = organize_complex(burgers, burgers_order, burgers_extra, burgers_large)
                [d2[c].append(burgers2[c][r2]) for r2 in xrange(len(burgers2[cols[0]])) for c in cols]
            elif subcategory == 'MERCHANDISE':
                merchandise2 = organize_simple(merchandise, merchandise_order)
                [d2[c].append(merchandise2[c][r2]) for r2 in xrange(len(merchandise2[cols[0]])) for c in cols]
            subcategory = ''

        # Catch any items needing organizing.
        if d['Item'][r]:
            if subcategory == 'BEER':
                [beer[c].append(d[c][r]) for c in cols]
                continue
            elif subcategory == 'WINE':
                [wine[c].append(d[c][r]) for c in cols]
                continue
            elif subcategory == 'BURGERS':
                [burgers[c].append(d[c][r]) for c in cols]
                continue
            elif subcategory == 'MERCHANDISE':
                [merchandise[c].append(d[c][r]) for c in cols]
                continue

        # Add row to output dictionary.
        [d2[c].append(d[c][r]) for c in cols]

    return d2


def organize_complex(d, order, extra, large):
    d2 = {}
    for c in cols:
        d2[c] = []

    return d2


def organize_simple(d, order):
    d2 = {}
    for c in cols:
        d2[c] = []

    return d2


def output_new(d, sales_clean):
    try:
        with open(sales_clean, 'wb') as out_f:
            out_d = csv.DictWriter(out_f, fieldnames=OrderedDict([(c, None) for c in cols]))
            out_d.writeheader()
            for r in xrange(len(d[cols[0]])):
                out_d.writerow(OrderedDict([(c, d[c][r]) for c in cols]))
            out_f.close()
        return True
    except:
        print 'Problem writing file "%s".' % sales_clean.split('/')[-1]
        return False


def main():
    sales_file_raw = '201309sales.csv'
    reports_dir = 'C:\\Users\\Malky\\Documents\\BD\\Reports'

    reports_dir = reports_dir.decode('utf-8').replace('\\','/')
    sales_file_clean = sales_file_raw.split('.')[0] + '_clean.csv'
    sales_clean = reports_dir + '/' + sales_file_clean

    try:
        sales_raw = glob.glob(reports_dir + '/' + sales_file_raw)[0]
    except:
        print 'File "%s" not found.' % sales_file_raw
        return

    d = parse_sales(sales_raw)
    done = output_new(d, sales_clean)

    print 'OK' if done else 'No Bueno'


if __name__ == '__main__':
    main()
