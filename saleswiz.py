# Python script to organize PixelPOS "Profit by Summary Group" sales reports
# into a useful and consistent format.

# To do:
# 1. Loop to read, convert and process all reports in folder.
# 2. Find items misplaced in subcategories, like Jr. Crispers in Chicken Wings
#    and place in correct subcategories.
# 3. Add Modifiers and Salads orders.
# 4. Put orders in dictionary, with empty extra and large if needed, send to
#    organize function using key, and with just a temporary dictionary for
#    each subcatergory.

import os
import glob
import csv
import sys
import subprocess

from collections import OrderedDict


# Chosen order for subcategories.
# Position 0 is search term, 1 is search exclude term, 2 is chosen name.
# Same for tuples for finding and naming large size.
beer_order = [('BOCK', '', 'DADDY BOCK'), ('LIGHT', 'BUD', 'DADDY LIGHT'),
              ('STELLA', '', 'STELLA'), ('BLUE', '', 'BLUE MOON'),
              ('STONE IPA', '', 'STONE IPA'), ('SMITH X', '', 'ALESMITH X'),
              ('ROCKET', '', 'RED ROCKET'), ('ROTAT', '', 'ROTATOR'),
              ('CASTLE', '', 'NEWCASTLE'), ('XX', '', 'DOS XX'),
              ('HEIN', '', 'HEINEKEN'), ('TOP BELG', '', 'SHOCK TOP BELGIAN'),
              ('NEGR', '', 'NEGRA MODELO'), ('BUD', '', 'BUD LIGHT'),
              ('ARDEN', '', 'HOEGAARDEN'), ('TOP IPA', '', 'SHOCK TOP IPA'),]
beer_extra = [('SAMPLE', '', 'BEER SAMPLER'),]
beer_large = ('CHER', '', 'PITCHER')

wine_order = [('CHAR', '', 'PENFOLDS CHARDONNAY'), ('LING', '', 'COLUMBIA RIESLING'),
              ('POLA', '', 'COPPOLA PINOT GRIGIO'), ('MOSC', '', 'MOSCATO'),
              ('FOLDS CAB', '', 'PENFOLDS CABERNET'), ('MERL', '', 'WOODBRIDGE MERLOT'),
              ('BERIN', '', 'BERINGER PINOT NOIR'), ('STRONG', '', 'RODNEY STRONG CABERNET'),
              ('MALB', '', 'TAMARI MALBEC'),]
wine_extra = [('SPARK', '', 'MIONETTO SPARKING SPLIT'),]
wine_large = ('BOTT', '', 'BOTTLE')

burgers_order = [('CLASSIC', '', 'CLASSIC'), ('DADDY', '', 'DADDY'),
                 ('TURK', '', 'TURKEY'), ('CHICK', '', 'CHICKEN'),
                 ('HAWAII', '', 'HAWAII FIVE-O'), ('MUSH', '', 'MUSHROOM MADNESS'),
                 ('LAMB', '', 'LAMB'),
                 ('ADD CLA', '', 'ADD CLASSIC PATTY'), ('ADD DAD', '', 'ADD DADDY PATTY'),
                 ('ADD TUR', '', 'ADD TURKEY'), ('ADD CHI', '', 'ADD CHICKEN PATTY'),
                 ('ADD HAW', '', 'ADD HAWAII FIVE-O PATTY'), ('ADD MUS', '', 'ADD MUSHROOM MADNESS PATTY'),
                 ('ADD LAM', '', 'ADD LAMB PATTY'),]
burgers_extra = [('VEG', '', 'VEGGIE'), ('BELLO', '', 'PORTABELLO'),]
burgers_large = ('1/2', '', '1/2 LB')

sandwiches_order = [('CHICK', '', 'CHICKEN'), ('SALM', '', 'SALMON'),
                    ('MIGNON', '', 'FILET MIGNON'), ('CHEE', '', 'GRILLED CHEESE'),
                    ('BLT', '', 'BLT'), ('PORK', '', 'PORK TENDERLOIN'),]

dawgs_sausages_order = [('BIG', '', 'BIG DAWG'), ('BACON', '', 'SMOKED BACON DAWG'),
                          ('ANTON', '', 'SAN ANTONIO CHILI DAWG'), ('CHEE', '', 'NACHO CHEESE DAWG'),
                          ('BRAT', '', 'BRATWURST SAUSAGE'), ('ITAL', '', 'SWEET ITALIAN SAUSAGE'),
                          ('ANDO', '', 'ANDOUILLE SAUSAGE'), ('CHICK', '', 'CHICKEN SOUTHWEST CALIENTE SAUSAGE'),]

chicken_wings_order = [('6 PIEC', '', '6 PIECES'), ('12 PIEC', '', '12 PIECES'),
                       ('18 PIEC', '', '18 PIECES'), ('50 PIE', '', '50 PIECES'),]

fries_rings_order = [('REGULAR FR', '', 'REGULAR FRIES'), ('GARLIC FR', 'CH', 'GARLIC FRIES'),
                       ('SWEET', '', 'SWEET POTATO FRIES'), ('ONION', '', 'ONION RINGS'),
                       ('TRIO', 'GARL', 'DADDY TRIO'), ('TRIO GA', '', 'DADDY TRIO GARLIC'),
                       ('LI CHEESE FR', '', 'CHILI CHEESE FRIES'), ('LI CHEESE GA', '', 'CHILI CHEESE GARLIC FRIES'),
                       ('HO CHEESE FR', '', 'NACHO CHEESE FRIES'), ('HO CHEESE GA', '', 'NACHO CHEESE GARLIC FRIES'),]

combo_order = [('REGULAR FR', '', 'REGULAR FRIES'), ('GARLIC FR', 'CH', 'GARLIC FRIES'),
               ('SWEET', '', 'SWEET POTATO FRIES'), ('ONION', '', 'ONION RINGS'),
               ('LI CHEESE', 'GA', 'CHILI CHEESE FRIES'), ('LI CHEESE GA', '', 'CHILI CHEESE GARLIC FRIES'),
               ('HO CHEESE', 'GA', 'NACHO CHEESE FRIES'), ('HO CHEESE GA', '', 'NACHO CHEESE GARLIC FRIES'),]

drinks_order = [('SMAL', '', 'SMALL DRINK'), ('LARG', '', 'LARGE DRINK'),
                ('COFF', 'DECA', 'COFFEE'), ('DECAF', '', 'DECAF COFFEE'),
                ('HOT TEA', '', 'HOT TEA'), ('ICE', '', 'ICED TEA'),
                ('WATER', '', 'BOTTLED WATER'),]

desserts_order = [('VANIL', '', 'VANILLA SHAKE'), ('CHOC', '', 'CHOCOLATE SHAKE'),
                  ('STRAW', '', 'STRAWBERRY SHAKE'), ('ROOT', '', 'ROOTBEER FLOAT SHAKE'),
                  ('ADD NUT', '', 'ADD NUTELLA'), ('ADD BAN', '', 'ADD BANANA'),
                  ('ADD PEA', '', 'ADD PEANUT BUTTER'), ('CREAM', '', 'SCOOP OF ICE CREAM'),
                  ('ADD SCO', '', 'ADD SCOOP OF ICE CREAM'), ('FUDGE', '', 'HOT FUDGE SUNDAE'),
                  ('SPLIT', '', 'BANANA SPLIT'), ('CAKE', '', 'CAKE POP'),]

kids_menu_order = [('BURG', 'CHEE', 'KIDS BURGER'), ('CHEESE BURG', '', 'KIDS CHEESE BURGER'),
                   ('DAWG', '', 'KIDS DAWG'), ('LED CHEE', '', 'KIDS GRILLED CHEESE'),
                   ('CRISP', '', 'KIDS CHICKEN CRISPERS'), ('MILK', '', 'KIDS MILK'),]

merchandise_order = [('HAT', '', 'HAT'), ('SHIRT', '', 'T-SHIRT'),]

# Items considered misplaced, with exclude term, chosen name and subcategory, large search term and chosen name, and 0 if small, 1 if small&large or 3 if an extra.
out_of_order = [('JR. CHICKEN CRISPERS', '', 'KIDS CHICKEN CRISPERS', 'KIDS MENU', '', '', 0),]

# PixelPOS "Profit by Summary Group" sales report columns.
# Nothing is in Cost, % Cost or Profit columns so ignoring.
keys = ['Category', 'Subcategory', 'Item', 'Col03', 'Col04', 'Col05', 'Col06',
        'Col07', 'Col08', 'Quantity', 'Col10', 'Col11', 'Col12', 'Col13',
        'Value', 'Col15', 'Cost', 'Col17', 'Col18', '% Cost', 'Col20', 'Col21',
        'Profit', 'Col23',]

# Output columns.
cols = ['Category', 'Subcategory', 'Item', 'Quantity', 'Value',]

# Search terms for rows in input file to exclude.
excludes = ['FIRE GRI', 'SUMMARY', 'REPORT', 'DESCRIPTION', 'PAGE', 'PRINTED',]


def parse_sales(sales_raw):
    """Parse PixelPOS "Profit by Summary Group" sales report."""
    d = {}
    d2 = {}
    beer = {}
    wine = {}
    burgers = {}
    sandwiches = {}
    dawgs_sausages = {}
    chicken_wings = {}
    fries_rings = {}
    combo = {}
    drinks = {}
    desserts = {}
    kids_menu = {}
    merchandise = {}

    for c in cols:
        d[c] = []
        d2[c] = []
        beer[c] = []
        wine[c] = []
        burgers[c] = []
        sandwiches[c] = []
        dawgs_sausages[c] = []
        chicken_wings[c] = []
        fries_rings[c] = []
        combo[c] = []
        drinks[c] = []
        desserts[c] = []
        kids_menu[c] = []
        merchandise[c] = []

    try:
        with open(sales_raw, 'rb') as in_f:
            in_d = csv.DictReader(in_f, fieldnames=keys, delimiter=',', quotechar='"')
            for row in in_d:
                [d[c].append(row[c].strip()) for c in cols]
            in_f.close()
    except:
        print 'Problem reading file "%s".' % sales_raw.split('/')[-1]
        return False

    category = ''
    subcategory = ''
    misplaced = []
    for r in xrange(len(d[cols[0]])):
        row = [d[c][r].upper() for c in cols]

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

        # Catch end of subcategory and organize.
        if d['Subcategory'][r] and any(row[2:]):
            if subcategory == 'BEER':
                beer2, misplaced = organize(sales_raw, subcategory, beer, beer_order, beer_extra, beer_large, misplaced)
                [d2[c].append(beer2[c][r2]) for r2 in xrange(len(beer2[cols[0]])) for c in cols]
            elif subcategory == 'WINE':
                wine2, misplaced = organize(sales_raw, subcategory, wine, wine_order, wine_extra, wine_large, misplaced)
                [d2[c].append(wine2[c][r2]) for r2 in xrange(len(wine2[cols[0]])) for c in cols]
            elif subcategory == 'BURGERS':
                burgers2, misplaced = organize(sales_raw, subcategory, burgers, burgers_order, burgers_extra, burgers_large, misplaced)
                [d2[c].append(burgers2[c][r2]) for r2 in xrange(len(burgers2[cols[0]])) for c in cols]
            elif subcategory == 'SANDWICHES':
                sandwiches2, misplaced = organize(sales_raw, subcategory, sandwiches, sandwiches_order, misplaced)
                [d2[c].append(sandwiches2[c][r2]) for r2 in xrange(len(sandwiches2[cols[0]])) for c in cols]
            elif subcategory == 'DAWGS & SAUSAGES':
                dawgs_sausages2, misplaced = organize(sales_raw, subcategory, dawgs_sausages, dawgs_sausages_order, misplaced)
                [d2[c].append(dawgs_sausages2[c][r2]) for r2 in xrange(len(dawgs_sausages2[cols[0]])) for c in cols]
            elif subcategory == 'CHICKEN WINGS':
                chicken_wings2, misplaced = organize(sales_raw, subcategory, chicken_wings, chicken_wings_order, misplaced)
                [d2[c].append(chicken_wings2[c][r2]) for r2 in xrange(len(chicken_wings2[cols[0]])) for c in cols]
            elif subcategory == 'FRIES & RINGS':
                fries_rings2, misplaced = organize(sales_raw, subcategory, fries_rings, fries_rings_order, misplaced)
                [d2[c].append(fries_rings2[c][r2]) for r2 in xrange(len(fries_rings2[cols[0]])) for c in cols]
            elif subcategory == 'COMBO':
                combo2, misplaced = organize(sales_raw, subcategory, combo, combo_order, misplaced)
                [d2[c].append(combo2[c][r2]) for r2 in xrange(len(combo2[cols[0]])) for c in cols]
            elif subcategory == 'DRINKS':
                drinks2, misplaced = organize(sales_raw, subcategory, drinks, drinks_order, misplaced)
                [d2[c].append(drinks2[c][r2]) for r2 in xrange(len(drinks2[cols[0]])) for c in cols]
            elif subcategory == 'DESSERTS':
                desserts2, misplaced = organize(sales_raw, subcategory, desserts, desserts_order, misplaced)
                [d2[c].append(desserts2[c][r2]) for r2 in xrange(len(desserts2[cols[0]])) for c in cols]
            elif subcategory == 'KIDS MENU':
                kids_menu2, misplaced = organize(sales_raw, subcategory, kids_menu, kids_menu_order, misplaced)
                [d2[c].append(kids_menu2[c][r2]) for r2 in xrange(len(kids_menu2[cols[0]])) for c in cols]
            elif subcategory == 'MERCHANDISE':
                merchandise2, misplaced = organize(sales_raw, subcategory, merchandise, merchandise_order, misplaced)
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
            elif subcategory == 'SANDWICHES':
                [sandwiches[c].append(d[c][r]) for c in cols]
                continue
            elif subcategory == 'DAWGS & SAUSAGES':
                [dawgs_sausages[c].append(d[c][r]) for c in cols]
                continue
            elif subcategory == 'CHICKEN WINGS':
                [chicken_wings[c].append(d[c][r]) for c in cols]
                continue
            elif subcategory == 'FRIES & RINGS':
                [fries_rings[c].append(d[c][r]) for c in cols]
                continue
            elif subcategory == 'COMBO':
                [combo[c].append(d[c][r]) for c in cols]
                continue
            elif subcategory == 'DRINKS':
                [drinks[c].append(d[c][r]) for c in cols]
                continue
            elif subcategory == 'DESSERTS':
                [desserts[c].append(d[c][r]) for c in cols]
                continue
            elif subcategory == 'KIDS MENU':
                [kids_menu[c].append(d[c][r]) for c in cols]
                continue
            elif subcategory == 'MERCHANDISE':
                [merchandise[c].append(d[c][r]) for c in cols]
                continue

        # Add row to output dictionary.
        [d2[c].append(d[c][r].upper()) for c in cols]

    # Insert or consolidate misplaced items into correct subcategory positions.
    #for mis in misplaced:
    #    Insert or consolidate in correct d[c][r] at position determined by mis[2] (0=small, 1=small&large, 2=extra).
    #    Delete item from misplaced.

    return d2, misplaced  # MAW


def organize(sales_raw, subcategory, d, order, extra=[], large=(), misplaced=[]):
    """Organize report subcategories into consistent items and order."""
    all = [[d[c][r].upper() for c in cols] for r in xrange(len(d[cols[0]]))]
    item = cols.index('Item')
    quantity = cols.index('Quantity')
    value = cols.index('Value')

    if large:
        if extra:
            smalls = [all[i] for i in xrange(len(all)) if not large[0] in all[i][item] and not any(e[0] in all[i][item] for e in extra)]
            larges = [all[i] for i in xrange(len(all)) if large[0] in all[i][item] and not (large[1] and large[1] in all[i][item]) and not any(e[0] in all[i][item] for e in extra)]
            extras = [all[i] for i in xrange(len(all)) if any(e[0] in all[i][item] for e in extra)]
        else:
            smalls = [all[i] for i in xrange(len(all)) if not large[0] in all[i][item]]
            larges = [all[i] for i in xrange(len(all)) if large[0] in all[i][item] and not (large[1] and large[1] in all[i][item])]
            extras = []
    else:
        if extra:
            smalls = [all[i] for i in xrange(len(all)) if not any(e[0] in all[i][item] for e in extra)]
            larges = []
            extras = [all[i] for i in xrange(len(all)) if any(e[0] in all[i][item] for e in extra)]
        else:
            smalls = all[:][:]
            larges = []
            extras = []

    orders = [order[:], order[:], extra[:],]

    servingtypes = [smalls[:][:], larges[:][:], extras[:][:],]
    out = [[[] for i in xrange(len(ordr)) if servingtypes[j]] for j, ordr in enumerate(orders)]
    for i, servingtype in enumerate(servingtypes):
        ordr = orders[i][:]
        for serving in servingtype:
            # Find items and include in out list, consolidating if necessary.
            found = False
            for k, o in enumerate(ordr):
                if o[0] in serving[item] and not (o[1] and o[1] in serving[item]):
                    found = True
                    if out[i][k]:
                        out[i][k][quantity] = str(int(out[i][k][quantity]) + int(serving[quantity]))
                        out[i][k][value] = str(float(out[i][k][value]) + float(serving[value]))
                    else:
                        out[i][k] = serving[:]
                        out[i][k][item] = o[2] + ' ' + large[2] if large and i == 1 else o[2]

            # If item not found, check if it's a known misplaced item.
            if not found:
                found_out = False
                for k, o in enumerate(out_of_order):
                    if o[0] in serving[item] and not (o[1] and o[1] in serving[item]) and (o[6] == 0 and i == 0 or o[6] == 1 and i <= 1 or o[6] == 2 and i == 2):
                        found_out = True
                        placed = serving[:]
                        placed[item] = o[2] + ' ' + o[5] if o[5] and (i == 1 or o[4] and o[4] in serving[item]) else o[2]
                        misplaced.append((o[3], k, i, placed))
                        print ('File "%s": Moving Subcategory "%s" item "%s" to Subcategory "%s" as item "%s".' %
                               (sales_raw.split('/')[-1], subcategory, serving[item], o[3], placed[item]))

                # Notify if item not known.
                if not found_out:
                    print ('File "%s": Subcategory "%s" item "%s" not known.' %
                           (sales_raw.split('/')[-1], subcategory, serving[item]))

        # Find missing order items and include as empty.
        if servingtype:
            for k, o in enumerate(ordr):
                if not out[i][k]:
                    out[i][k] = ['', '', o[2] + ' ' + large[2] if large and i == 1 else o[2], '0', '0.00',]

    # If there is a large size and there is a small or large misplaced item but not the opposite, include an empty opposite.
    misplaced2 = misplaced[:]
    for m, mis in enumerate(misplaced):
        if out_of_order[mis[1]][6] == 1:
            # Small. Find large.
            if mis[2] == 0:
                found = False
                for m2, mis2 in enumerate(misplaced):
                    if mis2[0] == mis[0] and mis2[2] == 1 and mis2[3][item].split(' ' + out_of_order[mis2[1]][5])[0] in mis[3][item]:
                        found = True
                if not found:
                    misplaced2.append((mis[0], mis[1], 1, ['', '', out_of_order[mis[1]][2] + ' ' + out_of_order[mis[1]][5], '0', '0.00',]))
            # Large. Find small.
            elif mis[2] == 1:
                found = False
                for m2, mis2 in enumerate(misplaced):
                    if mis2[0] == mis[0] and mis2[2] == 0 and mis2[3][item] in mis[3][item]:
                        found = True
                if not found:
                    misplaced2.append((mis[0], mis[1], 0, ['', '', out_of_order[mis[1]][2], '0', '0.00',]))

    # Return organized data as dictionary.
    d2 = {}
    for c in cols:
        d2[c] = []

    for servingtype in out:
        for serving in servingtype:
            for k, c in enumerate(cols):
                [d2[c].append(serving[k])]

    return d2, misplaced2


def output_new(d, sales_clean):
    """Write output csv file."""
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
    try:
        sales_file_in = sys.argv[1]
    except:
        print 'Input file [we]YYYYMM[DD]sales.csv or .xls required.'
        return

    reports_dir = 'C:\\Users\\Malky\\Documents\\BD\\Reports'

    extension = sales_file_in.split('.')[-1]
    if extension == 'xls':
        convert = subprocess.Popen(
            ['soffice.exe', '--headless', '--invisible', '--convert-to', 'csv', '--outdir', reports_dir, reports_dir + '\\' + sales_file_in],
             executable='C:\Program Files (x86)\LibreOffice 4\program\soffice.exe')
        convert.wait()
        if convert.returncode != 0:
            print 'Conversion of xls file to csv failed.'
            return
        sales_file_raw = ''.join(sales_file_in.split('.')[:-1]) + '.csv'
    elif extension !='csv':
        print 'Input file [we]YYYYMM[DD]sales.csv or .xls required.'
        return

    reports_dir = reports_dir.decode('utf-8').replace('\\','/')
    sales_file_clean = sales_file_raw.split('.')[0] + '_clean.csv'
    sales_clean = reports_dir + '/' + sales_file_clean

    try:
        sales_raw = glob.glob(reports_dir + '/' + sales_file_raw)[0]
    except:
        print 'File "%s" not found.' % sales_file_raw
        return

    d, misplaced = parse_sales(sales_raw)  # MAW
    print misplaced
    done = output_new(d, sales_clean) if d else False

    print 'OK' if done else 'No Bueno'


if __name__ == '__main__':
    main()
