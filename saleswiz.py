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

order_merchandise = [('HAT', 'HAT'), ('SHIRT', 'T-SHIRT'),]

keys = ['Category', 'Subcategory', 'Item', 'Col04', 'Col05', 'Col06', 'Col07',
        'Col08', 'Quantity', 'Col10', 'Col11', 'Col12', 'Col13', 'Value',
        'Col15', 'Cost', 'Col17', 'Col18', '% Cost', 'Col20', 'Col21',
        'Profit', 'Col04',]
cols = ['Category', 'Subcategory', 'Item', 'Quantity', 'Value',]
exclude = ['FIRE', 'Summary', 'Report', 'Description', 'Page', 'PRINTED',]


def parse_sales(sales_raw):
    with open(sales_raw, 'rb') as in_f:
        in_f.close()

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

    return d


def output_new(d, sales_clean):
    with open(sales_clean, 'wb') as out_f:
        out_d = csv.DictWriter(out_f, fieldname=OrderedDict([(c, None) for c in cols]))
        out_d.writeheader()
        for r in xrange(len(d[col[0]])):
            out_d.writerow(OrderedDict([(c, d[c][r]) for c in cols]))
        out_f.close()

    return True


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
