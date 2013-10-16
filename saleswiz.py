# Python script to organize PixelPOS "Profit by Summary Group" sales reports
# into a useful and consistent format.

# To do:
# 1. Loop to read, convert and process all reports in folder.
# 2. Place misplaced items in correct subcategories.
# 3. Recalculate subcategory and category tallies after removing misplaced items.
# 4. Order categories and subcategories.
# 5. Put extraneous items like KITCHEN MEMO into OTHER.

import os
import glob
import csv
import sys
import subprocess

from collections import OrderedDict


# Chosen order for subcategories.
# Position 0 is search term, 1 is search exclude term, 2 is chosen name.
# Same for tuples for finding and naming large size.
subcategories = {
    'BEER': {
        'order': [('BOCK', [], 'DADDY BOCK'), ('LIGHT', ['BUD',], 'DADDY LIGHT'),
                  ('STELLA', [], 'STELLA'), ('BLUE', [], 'BLUE MOON'),
                  ('STONE IPA', [], 'STONE IPA'), ('SMITH X', [], 'ALESMITH X'),
                  ('ROCKET', [], 'RED ROCKET'), ('ROTAT', [], 'ROTATOR'),
                  ('CASTLE', [], 'NEWCASTLE'), ('XX', [], 'DOS XX'),
                  ('HEIN', [], 'HEINEKEN'), ('TOP BELG', [], 'SHOCK TOP BELGIAN'),
                  ('NEGR', [], 'NEGRA MODELO'), ('BUD', [], 'BUD LIGHT'),
                  ('ARDEN', [], 'HOEGAARDEN'), ('TOP IPA', [], 'SHOCK TOP IPA'),],
        'extra': [('SAMPLE', [], 'BEER SAMPLER'),],
        'large': ('CHER', '', 'PITCHER'),},

    'WINE': {
        'order': [('CHAR', [], 'PENFOLDS CHARDONNAY'), ('LING', [], 'COLUMBIA RIESLING'),
                  ('POLA', [], 'COPPOLA PINOT GRIGIO'), ('MOSC', [], 'MOSCATO'),
                  ('CABER', ['STRONG',], 'PENFOLDS CABERNET'), ('MERL', [], 'WOODBRIDGE MERLOT'),
                  ('BERIN', [], 'BERINGER PINOT NOIR'), ('STRONG', [], 'RODNEY STRONG CABERNET'),
                  ('MALB', [], 'TAMARI MALBEC'),],
        'extra': [('SPARK', [], 'MIONETTO SPARKING SPLIT'), ('CORKA', [], 'CORKAGE FEE'),],
        'large': ('BOTT', '', 'BOTTLE'),},

    'BURGERS': {
        'order': [('CLASSIC', [], 'CLASSIC'), ('DADDY', [], 'DADDY'),
                  ('TURK', [], 'TURKEY'), ('CHICK', [], 'CHICKEN'),
                  ('HAWAII', [], 'HAWAII FIVE-O'), ('MUSH', [], 'MUSHROOM MADNESS'),
                  ('LAMB', [], 'LAMB'),
                  ('ADD CLA', [], 'ADD CLASSIC PATTY'), ('ADD DAD', [], 'ADD DADDY PATTY'),
                  ('ADD TUR', [], 'ADD TURKEY'), ('ADD CHI', [], 'ADD CHICKEN PATTY'),
                  ('ADD HAW', [], 'ADD HAWAII FIVE-O PATTY'), ('ADD MUS', [], 'ADD MUSHROOM MADNESS PATTY'),
                  ('ADD LAM', [], 'ADD LAMB PATTY'),],
        'extra': [('VEG', [], 'VEGGIE'), ('BELLO', [], 'PORTABELLO'),],
        'large': ('1/2', '', '1/2 LB'),},

    'SANDWICHES': {
        'order': [('CHICK', ['EXTRA',], 'CHICKEN'), ('SALM', ['EXTRA',], 'SALMON'),
                  ('MIGNON', ['EXTRA',], 'FILET MIGNON'),  ('PORK', ['EXTRA',], 'PORK TENDERLOIN'),
                  ('CHEE', [], 'GRILLED CHEESE'), ('BLT', [], 'BLT'),
                  ('EXTRA CHICK', [], 'EXTRA CHICKEN'), ('EXTRA SALM', [], 'EXTRA SALMON'),
                  ('EXTRA FILET M', [], 'EXTRA FILET MIGNON'),  ('EXTRA PORK', [], 'EXTRA PORK TENDERLOIN'),],
        'extra': [],
        'large': (),},

    'DAWGS & SAUSAGES': {
        'order': [('BIG', [], 'BIG DAWG'), ('BACON', [], 'SMOKED BACON DAWG'),
                  ('ANTON', [], 'SAN ANTONIO CHILI DAWG'), ('CHEE', [], 'NACHO CHEESE DAWG'),
                  ('BRAT', [], 'BRATWURST SAUSAGE'), ('ITAL', [], 'SWEET ITALIAN SAUSAGE'),
                  ('ANDO', [], 'ANDOUILLE SAUSAGE'), ('CHICK', [], 'CHICKEN SOUTHWEST CALIENTE SAUSAGE'),],
        'extra': [],
        'large': (),},

    'CHICKEN WINGS': {
        'order': [('6 PIEC', [], '6 PIECES'), ('12 PIEC', [], '12 PIECES'),
                  ('18 PIEC', [], '18 PIECES'), ('50 PIE', [], '50 PIECES'),],
        'extra': [],
        'large': (),},

    'FRIES & RINGS': {
        'order': [('REGULAR FR', [], 'REGULAR FRIES'), ('GARLIC FR', ['CH', 'COMBO',], 'GARLIC FRIES'),
                  ('SWEET', [], 'SWEET POTATO FRIES'), ('ONION', [], 'ONION RINGS'),
                  ('TRIO', ['GARL',], 'DADDY TRIO'), ('TRIO GA', [], 'DADDY TRIO GARLIC'),
                  ('LI CHEESE FR', ['COMBO',], 'CHILI CHEESE FRIES'), ('LI CHEESE GA', ['COMBO',], 'CHILI CHEESE GARLIC FRIES'),
                  ('HO CHEESE FR', ['COMBO',], 'NACHO CHEESE FRIES'), ('HO CHEESE GA', ['COMBO',], 'NACHO CHEESE GARLIC FRIES'),],
        'extra': [],
        'large': (),},

    'COMBO': {
        'order': [('REGULAR FR', [], 'REGULAR FRIES'), ('GARLIC FR', ['CH',], 'GARLIC FRIES'),
                  ('SWEET', [], 'SWEET POTATO FRIES'), ('ONION', [], 'ONION RINGS'),
                  ('LI CHEESE', ['GA',], 'CHILI CHEESE FRIES'), ('LI CHEESE GA', [], 'CHILI CHEESE GARLIC FRIES'),
                  ('HO CHEESE', ['GA',], 'NACHO CHEESE FRIES'), ('HO CHEESE GA', [], 'NACHO CHEESE GARLIC FRIES'),],
        'extra': [],
        'large': (),},

    'SIDES & SALADS': {
        'order': [('BIG', [], 'BIG 50'),
                  ('SALAD', ['SIDE',], 'SALAD'), ('SIDE SALAD', [], 'SIDE SALAD'),
                  ('CHICK', [], 'ADD CHICKEN'), ('SALM', [], 'ADD SALMON'),
                  ('MIGNON', [], 'ADD FILET MIGNON'), ('AVOC', [], 'ADD AVOCADO'),
                  ('SLAW', ['LARGE',], 'SMALL COLE SLAW'), ('SLAW', ['SMALL',], 'LARGE COLE SLAW'),
                  ('CHILI', ['LARGE',], 'SMALL CHILI'), ('CHILI', ['SMALL',], 'LARGE CHILI'),
                  ('CRISP', ['LARGE',], 'SMALL CHICKEN CRISPERS'), ('CRISP', ['SMALL',], 'LARGE CHICKEN CRISPERS'),
                  ('ZUC', [], 'ZUCCHINI DIPPERS'), ('POPP', [], 'MUSHROOM POPPERS'),
                  ('CHIP', [], 'CHIPS'), ('BROWN', [], 'HASH BROWNS'),],
        'extra': [],
        'large': (),},

    'MODIFIERS': {
        'order': [('CHEDD', ['EXTRA',], 'ADD CHEDDAR CHEESE'), ('AMER', ['EXTRA',], 'ADD AMERICAN CHEESE'),
                  ('SWISS', ['EXTRA',], 'ADD SWISS CHEESE'), ('PEPPERJ', ['EXTRA',], 'ADD PEPPERJACK CHEESE'),
                  ('GORG', ['EXTRA',], 'ADD GORGONZOLA CHEESE'), ('NACHO', ['EXTRA',], 'ADD NACHO CHEESE'),
                  ('BACON', ['EXTRA',], 'ADD BACON'), ('AVOC', ['EXTRA',], 'ADD AVOCADO'),
                  ('MUSH', ['EXTRA',], 'ADD SAUTEED MUSHROOMS & ONIONS'), ('PINE', [], 'ADD GRILLED PINEAPPLE'),
                  ('EGG', [], 'ADD FRIED EGG'),  ('CHILI', [], 'ADD CHILI'),
                  ('RING', ['SUB',], 'ADD ONION RINGS'), ('BELL', [], 'ADD GRILLED BELL PEPPER'),
                  ('EXTRA CHEDD', [], 'EXTRA CHEDDAR CHEESE'), ('EXTRA AMER', [], 'EXTRA AMERICAN CHEESE'),
                  ('EXTRA SWISS', [], 'EXTRA SWISS CHEESE'), ('EXTRA PEPPERJ', [], 'EXTRA PEPPERJACK CHEESE'),
                  ('EXTRA GORG', [], 'EXTRA GORGONZOLA CHEESE'), ('EXTRA NACHO', [], 'EXTRA NACHO CHEESE'),
                  ('EXTRA BACON', [], 'EXTRA BACON'), ('EXTRA AVOC', [], 'EXTRA AVOCADO'),
                  ('EXTRA SAUT', [], 'EXTRA SAUTEED MUSHROOMS & ONIONS'),
                  ('SUB GAR', [], 'SUB GARLIC FRIES'), ('SWEET PO', [], 'SUB SWEET POTATO FRIES'),
                  ('RING', ['TOP',], 'SUB ONION RINGS'),
                  ('SAUCE', ['BBQ',], 'ADD SAUCE'), ('BBQ', [], 'ADD BBQ SAUCE'),
                  ('RANCH', [], 'ADD RANCH'), ('PESTO', [], 'ADD PESTO AIOLI'),
                  ('DIJON', [], 'ADD DIJON-MAYO AIOLI'), ('MAYO', ['DIJ',], 'ADD MAYO AIOLI'),
                  ('KETCH', [], 'ADD SMOKED KETCHEP'),
                  ('ONIONS-D', [], 'ADD CARAMELIZED ONIONS - DAWGS'), ('KRAUT', [], 'ADD SAUERKRAUT - DAWGS'),
                  ('SWEET PE', [], 'ADD SWEET PEPPERS - DAWGS'), ('SPICY PE', [], 'ADD SPICY PEPPERS - DAWGS'),],
        'extra': [],
        'large': (),},

    'DRINKS': {
        'order': [('SMAL', [], 'SMALL SODA'), ('LARG', [], 'LARGE SODA'),
                  ('PITCH', [], 'PITCHER SODA'),
                  ('COFF', ['DECA',], 'COFFEE'), ('DECAF', [], 'DECAF COFFEE'),
                  ('HOT TEA', [], 'HOT TEA'), ('ICE', [], 'ICED TEA'),
                  ('WATER', [], 'BOTTLED WATER'), ('MONST', [], 'MONSTER'),],
        'extra': [],
        'large': (),},

    'DESSERTS': {
        'order': [('VANIL', [], 'VANILLA SHAKE'), ('CHOC', [], 'CHOCOLATE SHAKE'),
                  ('STRAW', [], 'STRAWBERRY SHAKE'), ('ROOT', [], 'ROOTBEER FLOAT SHAKE'),
                  ('POLIT', [], 'NEAPOLITAN SHAKE'),
                  ('ADD NUT', [], 'ADD NUTELLA'), ('ADD BAN', [], 'ADD BANANA'),
                  ('ADD PEA', [], 'ADD PEANUT BUTTER'), ('CREAM', [], 'SCOOP OF ICE CREAM'),
                  ('ADD SCO', [], 'ADD SCOOP OF ICE CREAM'), ('FUDGE', [], 'HOT FUDGE SUNDAE'),
                  ('SPLIT', [], 'BANANA SPLIT'), ('CAKE', [], 'CAKE POP'),],
        'extra': [],
        'large': (),},

    'KIDS MENU': {
        'order': [('BURG', ['CHEE', 'TUR',], 'KIDS BURGER'), ('CHEESE BURG', [], 'KIDS CHEESE BURGER'),
                  ('KEY BURG', [], 'KIDS TURKEY BURGER'), ('DAWG', [], 'KIDS DAWG'),
                  ('LED CHEE', [], 'KIDS GRILLED CHEESE'), ('CRISP', [], 'KIDS CHICKEN CRISPERS'),
                  ('MILK', [], 'KIDS MILK'),],
        'extra': [],
        'large': (),},

    'BREAKFAST SANDWICHES': {
        'order': [('PLAIN', [], 'PLAIN OMELET'), ('WHITE', [], 'EGG WHITE OMELET'),
                  ('CHOR', [], 'CHORIZO OMELET'), ('THREE', [], '3 CHEESE OMELET'),
                  ('BAC', [], 'BACON ONION 3 CHEESE OMELET'),],
        'extra': [],
        'large': (),},

    'SUB PATTIES': {
        'order': [],
        'extra': [],
        'large': (),},

    'MERCHANDISE': {
        'order': [('HAT', [], 'HAT'), ('SHIRT', [], 'T-SHIRT'),],
        'extra': [],
        'large': (),},

    'GIFT CARD': {
        'order': [('OPEN', [], 'OPEN ITEM'),],
        'extra': [],
        'large': (),},
}

# Items considered misplaced, with exclude term, chosen name and subcategory, large search term and chosen name,
# and 0 if small, 1 if small&large or 3 if an extra.
out_of_order = [('ADD 1/2 CLASSIC', [], 'ADD CLASSIC PATTY', 'BURGERS', '1/2', '1/2 LB', 1),
                ('ADD 1/2 TURKEY', [], 'ADD TURKEY PATTY', 'BURGERS', '1/2', '1/2 LB', 1),
                ('SUB DADDY PATTY', [], 'ADD DADDY PATTY', 'BURGERS', '1/2', '1/2 LB', 1),
                ('SUB 1/2LB DADDY PATTY', [], 'ADD DADDY PATTY', 'BURGERS', '1/2', '1/2 LB', 1),
                ('EXTRA CHICKEN', [], 'EXTRA CHICKEN', 'SANDWICHES', '', '', 0),
                ('EXTRA FILET MIGNON', [], 'EXTRA FILET MIGNON', 'SANDWICHES', '', '', 0),
                ('SUPERBOWL SPECIAL', [], '50 PIECES', 'CHICKEN WINGS', '', '', 0),
                ('COMBO REGULAR FRIES', [], 'REGULAR FRIES', 'COMBO', '', '', 0),
                ('COMBO GARLIC FRIES', [], 'GARLIC FRIES', 'COMBO', '', '', 0),
                ('COMBO SWEET POTATO', [], 'SWEET POTATO FRIES', 'COMBO', '', '', 0),
                ('COMBO ONION RINGS', [], 'ONION RINGS', 'COMBO', '', '', 0),
                ('COMBO NACHO CHEESE FRIES', [], 'NACHO CHEESE FRIES', 'COMBO', '', '', 0),
                ('JR. CHICKEN CRISPERS', [], 'KIDS CHICKEN CRISPERS', 'KIDS MENU', '', '', 0),]

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
    for c in cols:
        d[c] = []
        d2[c] = []

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
                dtemp = {}
                for c in cols:
                    dtemp[c] = []

        # Catch end of category.
        if d['Category'][r] and any(row[1:]):
            category = ''

        # Catch end of subcategory and organize.
        if d['Subcategory'][r] and any(row[2:]):
            dtemp2, misplaced = organize(sales_raw, subcategory, dtemp, misplaced)
            [d2[c].append(dtemp2[c][r2]) for r2 in xrange(len(dtemp2[cols[0]])) for c in cols]
            subcategory = ''

        # Having caught any out of order data in a subcategory, ignore subcategory if it should have no data.
        if d['Subcategory'][r] and not subcategories[d['Subcategory'][r]]['order']:
            continue

        # Catch any items for organizing.
        if d['Item'][r]:
            [dtemp[c].append(d[c][r]) for c in cols]
            continue

        # Add row to output dictionary.
        [d2[c].append(d[c][r].upper()) for c in cols]

    # Insert or consolidate misplaced items into correct subcategory positions.
    #for mis in misplaced:
    #    Insert or consolidate in correct d[c][r] at position determined by mis[2] (0=small, 1=small&large, 2=extra).
    #    Delete item from misplaced.
    #for m, mis in misplaced:
    #    serving =

    return d2, misplaced  # MAW


def organize(sales_raw, subcategory, d, misplaced=[]):
    """Organize report subcategories into consistent items and order."""
    all = [[d[c][r].upper() for c in cols] for r in xrange(len(d[cols[0]]))]
    order = subcategories[subcategory]['order']
    extra = subcategories[subcategory]['extra']
    large = subcategories[subcategory]['large'][:]
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
    include = [True if smalls or large else False, True if larges or large else False, True if extra else False]
    servingtypes = [smalls[:][:], larges[:][:], extras[:][:],]
    out = [[[] for i in xrange(len(ordr)) if include[j]] for j, ordr in enumerate(orders)]
    for i, servingtype in enumerate(servingtypes):
        ordr = orders[i][:]
        for serving in servingtype:
            # Find items and include in out list, consolidating if necessary.
            found = False
            for k, o in enumerate(ordr):
                if o[0] in serving[item] and not any(o[1] and o[1][j] in serving[item] for j in xrange(len(o[1]))):
                    # and not any(ooo[0] in serving[item] and not (ooo[1] and ooo[1] in serving[item]) and ooo[3] not in subcategory for ooo in out_of_order):
                    found = True
                    if out[i][k]:
                        out[i][k][quantity] = str(float(out[i][k][quantity]) + float(fixnum(serving[quantity])))
                        out[i][k][value] = str(float(out[i][k][value]) + float(fixnum(serving[value])))
                    else:
                        out[i][k] = serving[:]
                        out[i][k][item] = o[2] + ' ' + large[2] if large and i == 1 else o[2]
                        out[i][k][quantity] = fixnum(serving[quantity])
                        out[i][k][value] = fixnum(serving[value])

            # If item not found, check if it's a known misplaced item.
            if not found:
                found_out = False
                for k, o in enumerate(out_of_order):
                    if o[0] in serving[item] and not any(o[1] and o[1][j] in serving[item] for j in xrange(len(o[1]))):
                        if o[6] == 0 and i == 0 or o[6] == 1 or o[6] == 2 and i == 2:
                            found_out = True
                            placed = serving[:]
                            placed[item] = o[2]
                            placed[quantity] = fixnum(serving[quantity])
                            placed[value] = fixnum(serving[value])
                            if o[6] == 1:
                                if o[4] in serving[item]:
                                    placed[item] += ' ' + o[5]
                                    mistype = 1
                                else:
                                    mistype = 0
                            else:
                                mistype = i
                            misplaced.append((o[3], k, mistype, placed))
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
        else:
            if large and i <= 1 or extra and i == 2:
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


def fixnum(num):
    """Fix LibreOffice conversion glitch which sometimes puts '.' instead of
       ',' in numbers."""
    snum = str(num)
    bits = snum.split('.')
    if len(bits) > 1:
        if len(bits[-1]) > 2:
            fixed = ','.join(bits)
        else:
            fixed = ','.join(bits[:-1]) + '.' + bits[-1]
    else:
        fixed = snum

    return fixed


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
    elif extension == 'csv':
        sales_file_raw = sales_file_in
    else:
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
