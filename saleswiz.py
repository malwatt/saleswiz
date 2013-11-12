# Python script to organize PixelPOS "Profit by Summary Group" sales reports
# into a useful and consistent format.

import os
import csv
import sys
import subprocess

from glob import glob
from collections import OrderedDict


# Chosen order for subcategories. Grouped into categories.
# Position 0 is search term, 1 is search exclude terms, 2 is chosen name.
# Same for tuples for finding and naming extra items and similarly for large size.
subcategories = OrderedDict([
    ('BEER', {
        'category': 'BEER',
        'order': [('BOCK', [], 'DADDY BOCK'), ('LIGHT', ['BUD',], 'DADDY LIGHT'),
                  ('STELLA', [], 'STELLA'), ('BLUE', [], 'BLUE MOON'),
                  ('STONE IPA', [], 'STONE IPA'), ('EXTRA P', [], 'ALESMITH X'),
                  ('ROCKET', [], 'RED ROCKET'), ('ROTAT', [], 'ROTATOR'),
                  ('CASTLE', [], 'NEWCASTLE'), ('XX', [], 'DOS XX'),
                  ('HEIN', [], 'HEINEKEN'), ('TOP BELG', [], 'SHOCK TOP BELGIAN'),
                  ('NEGR', [], 'NEGRA MODELO'), ('BUD', [], 'BUD LIGHT'),
                  ('ARDEN', [], 'HOEGAARDEN'), ('TOP IPA', [], 'SHOCK TOP IPA'),],
        'extra': [('SAMPLE', [], 'BEER SAMPLER'),],
        'large': ('CHE', '', 'PITCHER'),},),

    ('WINE', {
        'category': 'WINE',
        'order': [('CHAR', [], 'PENFOLDS CHARDONNAY'), ('LING', [], 'COLUMBIA RIESLING'),
                  ('POLA', [], 'COPPOLA PINOT GRIGIO'), ('MOSC', [], 'MOSCATO'),
                  ('CABER', ['STRONG',], 'PENFOLDS CABERNET'), ('MERL', [], 'WOODBRIDGE MERLOT'),
                  ('BERIN', [], 'BERINGER PINOT NOIR'), ('STRONG', [], 'RODNEY STRONG CABERNET'),
                  ('MALB', [], 'TAMARI MALBEC'),],
        'extra': [('SPARK', [], 'MIONETTO SPARKING SPLIT'), ('CORKA', [], 'CORKAGE FEE'),],
        'large': ('BOTT', '', 'BOTTLE'),},),

    ('DRINKS', {
        'category': 'BEVERAGES',
        'order': [('SMAL', [], 'SMALL SODA'), ('LARG', [], 'LARGE SODA'),
                  ('PITCH', [], 'PITCHER SODA'),
                  ('COFF', ['DECA',], 'COFFEE'), ('DECAF', [], 'DECAF COFFEE'),
                  ('HOT TEA', [], 'HOT TEA'), ('ICE', [], 'ICED TEA'),
                  ('WATER', [], 'BOTTLED WATER'), ('MONST', [], 'MONSTER'),],
        'extra': [],
        'large': (),},),

    ('BURGERS', {
        'category': 'FOOD',
        'order': [('CLASSIC', ['PAT',], 'CLASSIC'), ('DADDY', ['PAT',], 'DADDY'),
                  ('TURK', ['PAT',], 'TURKEY'), ('CHICK', ['PAT',], 'CHICKEN'),
                  ('HAWAII', ['PAT'], 'HAWAII FIVE-O'), ('MUSH', ['PAT',], 'MUSHROOM MADNESS'),
                  ('LAMB', ['PAT',], 'LAMB'),
                  ('ADD CLA', [], 'ADD CLASSIC PATTY'), ('ADD DAD', [], 'ADD DADDY PATTY'),
                  ('ADD TUR', [], 'ADD TURKEY PATTY'), ('ADD CHI', [], 'ADD CHICKEN PATTY'),
                  ('ADD HAW', [], 'ADD HAWAII FIVE-O PATTY'), ('ADD MUS', [], 'ADD MUSHROOM MADNESS PATTY'),
                  ('ADD LAM', [], 'ADD LAMB PATTY'),],
        'extra': [('VEG', ['PAT',], 'VEGGIE'), ('BELLO', [], 'PORTABELLO'),
                  ('ADD VEG', [], 'ADD VEGGIE PATTY'),],
        'large': ('1/2', '', '1/2 LB'),},),

    ('SANDWICHES', {
        'category': 'FOOD',
        'order': [('CHICK', ['EXTRA',], 'CHICKEN'), ('SALM', ['EXTRA',], 'SALMON'),
                  ('MIGNON', ['EXTRA',], 'FILET MIGNON'),  ('PORK', ['EXTRA',], 'PORK TENDERLOIN'),
                  ('CHEE', [], 'GRILLED CHEESE'), ('BLT', [], 'BLT'),
                  ('EXTRA CHICK', [], 'EXTRA CHICKEN'), ('EXTRA SALM', [], 'EXTRA SALMON'),
                  ('EXTRA FILET M', [], 'EXTRA FILET MIGNON'),  ('EXTRA PORK', [], 'EXTRA PORK TENDERLOIN'),],
        'extra': [],
        'large': (),},),

    ('DAWGS & SAUSAGES', {
        'category': 'FOOD',
        'order': [('BIG', [], 'BIG DAWG'), ('BACON', [], 'SMOKED BACON DAWG'),
                  ('ANTON', [], 'SAN ANTONIO CHILI DAWG'), ('CHEE', [], 'NACHO CHEESE DAWG'),
                  ('BRAT', [], 'BRATWURST SAUSAGE'), ('ITAL', [], 'SWEET ITALIAN SAUSAGE'),
                  ('ANDO', [], 'ANDOUILLE SAUSAGE'), ('CHICK', [], 'CHICKEN SOUTHWEST CALIENTE SAUSAGE'),],
        'extra': [],
        'large': (),},),

    ('CHICKEN WINGS', {
        'category': 'FOOD',
        'order': [('6 PIEC', [], '6 PIECES'), ('12 PIEC', [], '12 PIECES'),
                  ('18 PIEC', [], '18 PIECES'), ('50 PIE', [], '50 PIECES'),],
        'extra': [],
        'large': (),},),

    ('FRIES & RINGS', {
        'category': 'FOOD',
        'order': [('REGULAR FR', [], 'REGULAR FRIES'), ('GARLIC FR', ['CH', 'COMBO',], 'GARLIC FRIES'),
                  ('SWEET', [], 'SWEET POTATO FRIES'), ('ONION', [], 'ONION RINGS'),
                  ('TRIO', ['GARL',], 'DADDY TRIO'), ('TRIO GA', [], 'DADDY TRIO GARLIC'),
                  ('LI CHEESE FR', ['COMBO',], 'CHILI CHEESE FRIES'), ('LI CHEESE GA', ['COMBO',], 'CHILI CHEESE GARLIC FRIES'),
                  ('HO CHEESE FR', ['COMBO',], 'NACHO CHEESE FRIES'), ('HO CHEESE GA', ['COMBO',], 'NACHO CHEESE GARLIC FRIES'),],
        'extra': [],
        'large': (),},),

    ('COMBO', {
        'category': 'FOOD',
        'order': [('REGULAR FR', [], 'REGULAR FRIES'), ('GARLIC FR', ['CH',], 'GARLIC FRIES'),
                  ('SWEET', [], 'SWEET POTATO FRIES'), ('ONION', [], 'ONION RINGS'),
                  ('LI CHEESE', ['GA',], 'CHILI CHEESE FRIES'), ('LI CHEESE GA', [], 'CHILI CHEESE GARLIC FRIES'),
                  ('HO CHEESE', ['GA',], 'NACHO CHEESE FRIES'), ('HO CHEESE GA', [], 'NACHO CHEESE GARLIC FRIES'),],
        'extra': [],
        'large': (),},),

    ('SIDES & SALADS', {
        'category': 'FOOD',
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
        'large': (),},),

    ('MODIFIERS', {
        'category': 'FOOD',
        'order': [('CHEDD', ['EXTRA',], 'ADD CHEDDAR CHEESE'), ('AMER', ['EXTRA',], 'ADD AMERICAN CHEESE'),
                  ('SWISS', ['EXTRA',], 'ADD SWISS CHEESE'), ('PEPPERJ', ['EXTRA',], 'ADD PEPPERJACK CHEESE'),
                  ('GORG', ['EXTRA',], 'ADD GORGONZOLA CHEESE'), ('NACHO', ['EXTRA',], 'ADD NACHO CHEESE'),
                  ('BACON', ['EXTRA',], 'ADD BACON'), ('AVOC', ['EXTRA',], 'ADD AVOCADO'),
                  ('MUSH', ['EXTRA',], 'ADD SAUTEED MUSHROOMS & ONIONS'), ('PINE', [], 'ADD GRILLED PINEAPPLE'),
                  ('EGG', [], 'ADD FRIED EGG'),  ('CHILI', [], 'ADD CHILI'),
                  ('RING', ['SUB',], 'ADD ONION RINGS'), ('BELL', [], 'ADD GRILLED BELL PEPPER'),
                  ('SAUCE', ['BBQ',], 'ADD SAUCE'), ('BBQ', [], 'ADD BBQ SAUCE'),
                  ('RANCH', [], 'ADD RANCH'), ('PESTO', [], 'ADD PESTO AIOLI'),
                  ('DIJON', [], 'ADD DIJON-MAYO AIOLI'), ('MAYO', ['DIJ',], 'ADD MAYO AIOLI'),
                  ('KETCH', [], 'ADD SMOKED KETCHEP'),
                  ('EXTRA CHEDD', [], 'EXTRA CHEDDAR CHEESE'), ('EXTRA AMER', [], 'EXTRA AMERICAN CHEESE'),
                  ('EXTRA SWISS', [], 'EXTRA SWISS CHEESE'), ('EXTRA PEPPERJ', [], 'EXTRA PEPPERJACK CHEESE'),
                  ('EXTRA GORG', [], 'EXTRA GORGONZOLA CHEESE'), ('EXTRA NACHO', [], 'EXTRA NACHO CHEESE'),
                  ('EXTRA BACON', [], 'EXTRA BACON'), ('EXTRA AVOC', [], 'EXTRA AVOCADO'),
                  ('EXTRA SAUT', [], 'EXTRA SAUTEED MUSHROOMS & ONIONS'),
                  ('SUB GAR', [], 'SUB GARLIC FRIES'), ('SWEET PO', [], 'SUB SWEET POTATO FRIES'),
                  ('RING', ['TOP',], 'SUB ONION RINGS'),
                  ('ONIONS-D', [], 'ADD CARAMELIZED ONIONS - DAWGS'), ('KRAUT', [], 'ADD SAUERKRAUT - DAWGS'),
                  ('SWEET PE', [], 'ADD SWEET PEPPERS - DAWGS'), ('SPICY PE', [], 'ADD SPICY PEPPERS - DAWGS'),],
        'extra': [],
        'large': (),},),

    ('DESSERTS', {
        'category': 'FOOD',
        'order': [('VANIL', [], 'VANILLA SHAKE'), ('CHOC', [], 'CHOCOLATE SHAKE'),
                  ('STRAW', [], 'STRAWBERRY SHAKE'), ('ROOT', [], 'ROOTBEER FLOAT SHAKE'),
                  ('POLIT', [], 'NEAPOLITAN SHAKE'),
                  ('ADD NUT', [], 'ADD NUTELLA'), ('ADD BAN', [], 'ADD BANANA'),
                  ('ADD PEA', [], 'ADD PEANUT BUTTER'), ('CREAM', [], 'SCOOP OF ICE CREAM'),
                  ('ADD SCO', [], 'ADD SCOOP OF ICE CREAM'), ('FUDGE', [], 'HOT FUDGE SUNDAE'),
                  ('SPLIT', [], 'BANANA SPLIT'), ('CAKE', [], 'CAKE POP'),],
        'extra': [],
        'large': (),},),

    ('KIDS MENU', {
        'category': 'FOOD',
        'order': [('BURG', ['CHEE', 'TUR',], 'KIDS BURGER'), ('CHEESE BURG', [], 'KIDS CHEESE BURGER'),
                  ('KEY BURG', [], 'KIDS TURKEY BURGER'), ('DAWG', [], 'KIDS DAWG'),
                  ('LED CHEE', [], 'KIDS GRILLED CHEESE'), ('CRISP', [], 'KIDS CHICKEN CRISPERS'),
                  ('MILK', [], 'KIDS MILK'),],
        'extra': [],
        'large': (),},),

    ('BREAKFAST SANDWICHES', {
        'category': 'FOOD',
        'order': [('PLAIN', [], 'PLAIN OMELET'), ('WHITE', [], 'EGG WHITE OMELET'),
                  ('CHOR', [], 'CHORIZO OMELET'), ('THREE', [], '3 CHEESE OMELET'),
                  ('BAC', [], 'BACON ONION 3 CHEESE OMELET'),],
        'extra': [],
        'large': (),},),

    ('MERCHANDISE', {
        'category': 'MERCHANDISE',
        'order': [('HAT', [], 'HAT'), ('SHIRT', [], 'T-SHIRT'),],
        'extra': [],
        'large': (),},),

    ('GIFT CARD', {
        'category': 'OTHER',
        'order': [('OPEN', [], 'OPEN ITEM'),],
        'extra': [],
        'large': (),},),

    ('MISCELLANEOUS', {
        'category': 'OTHER',
        'order': [('KITCHEN MEMO', [], 'KITCHEN MEMO'),],
        'extra': [],
        'large': (),},),

    ('SUB PATTIES', {
        'category': '',
        'order': [],
        'extra': [],
        'large': (),},),
])

# Items considered misplaced, with exclude term, chosen name and subcategory, large search term and chosen name,
# and 0 if small, 1 if small&large or 3 if an extra.
out_of_order = [('ADD 1/2 CLASSIC', [], 'ADD CLASSIC PATTY', 'BURGERS', '1/2', '1/2 LB', 1),
                ('ADD 1/2 DADDY', [], 'ADD DADDY PATTY', 'BURGERS', '1/2', '1/2 LB', 1),
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
                ('JR. CHICKEN CRISPERS', [], 'KIDS CHICKEN CRISPERS', 'KIDS MENU', '', '', 0),
                ('KITCHEN MEMO', [], 'KITCHEN MEMO', 'MISCELLANEOUS', '', '', 0),]

# PixelPOS "Profit by Summary Group" sales report columns.
# Nothing is in Cost, % Cost or Profit columns so ignoring.
keys = ['Category', 'Subcategory', 'Item', 'Col03', 'Col04', 'Col05', 'Col06',
        'Col07', 'Col08', 'Quantity', 'Col10', 'Col11', 'Col12', 'Col13',
        'Value', 'Col15', 'Cost', 'Col17', 'Col18', '% Cost', 'Col20', 'Col21',
        'Profit', 'Col23',]

# Output columns.
cols = ['Category', 'Subcategory', 'Item', 'Quantity', 'Value',]
item = cols.index('Item')
quantity = cols.index('Quantity')
value = cols.index('Value')

# Search terms for rows in input file to exclude.
excludes = ['FIRE GRI', 'SUMMARY', 'REPORT', 'DESCRIPTION', 'PAGE', 'PRINTED',]


def parse_sales(sales_raw):
    """Parse PixelPOS "Profit by Summary Group" sales report."""
    ok = True
    d = {}
    d2 = {}
    d3 = {}
    for c in cols:
        d[c] = []
        d2[c] = []
        d3[c] = []

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
    categories_found = []
    subcategories_found = []
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
                categories_found.append(category)

        # Catch subcategories and ignore duplicates.
        if d['Subcategory'][r] and not any(row[2:]):
            if subcategory:
                continue
            else:
                subcategory = d['Subcategory'][r]
                subcategories_found.append(subcategory)
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

        # Having caught any out of order data in a subcategory, ignore subcategory if it is not for final inclusion.
        if d['Subcategory'][r] and not subcategories[d['Subcategory'][r]]['category']:
            continue

        # Catch items for organizing.
        if d['Item'][r]:
            [dtemp[c].append(d[c][r]) for c in cols]
            continue

        # Add row to output dictionary.
        [d2[c].append(d[c][r].upper()) for c in cols]

    # Add categories and subcategories that are missing, with empty items.
    for s in subcategories:
        if subcategories[s]['category'] and s not in subcategories_found:
            if subcategories[s]['category'] in categories_found:
                r = len(d2[cols[0]]) - d2[cols[0]][::-1].index(subcategories[s]['category']) - 1
                row = ['', s, '', '', '',]
                [d2[c].insert(r, row[k]) for k, c in enumerate(cols)]
                for o in subcategories[s]['order']:
                    r += 1
                    row = ['', '', o[2], '0', '0.00',]
                    [d2[c].insert(r, row[k]) for k, c in enumerate(cols)]
                if subcategories[s]['large']:
                    for o in subcategories[s]['order']:
                        r += 1
                        row = ['', '', o[2] + ' ' + subcategories[s]['large'][2], '0', '0.00',]
                        [d2[c].insert(r, row[k]) for k, c in enumerate(cols)]
                for o in subcategories[s]['extra']:
                    r += 1
                    row = ['', '', o[2], '0', '0.00',]
                    [d2[c].insert(r, row[k]) for k, c in enumerate(cols)]
                r += 1
                row = ['', s, '', '0', '0.00',]
                [d2[c].insert(r, row[k]) for k, c in enumerate(cols)]
            else:
                row = [subcategories[s]['category'], '', '', '', '',]
                [d2[c].insert(-1, row[k]) for k, c in enumerate(cols)]
                row = ['', s, '', '', '',]
                [d2[c].insert(-1, row[k]) for k, c in enumerate(cols)]
                for o in subcategories[s]['order']:
                    row = ['', '', o[2], '0', '0.00',]
                    [d2[c].insert(-1, row[k]) for k, c in enumerate(cols)]
                if subcategories[s]['large']:
                    for o in subcategories[s]['order']:
                        row = ['', '', o[2] + ' ' + subcategories[s]['large'][2], '0', '0.00',]
                        [d2[c].insert(-1, row[k]) for k, c in enumerate(cols)]
                for o in subcategories[s]['extra']:
                    row = ['', '', o[2], '0', '0.00',]
                    [d2[c].insert(-1, row[k]) for k, c in enumerate(cols)]
                row = ['', s, '', '0', '0.00',]
                [d2[c].insert(-1, row[k]) for k, c in enumerate(cols)]
                row = [subcategories[s]['category'], '', '', '0', '0.00',]
                [d2[c].insert(-1, row[k]) for k, c in enumerate(cols)]

    # Consolidate misplaced items into correct item positions.
    total_quantity = 0
    total_value = 0
    for r in xrange(len(d2[cols[0]])):
        row = [d2[c][r] for c in cols]

        # Catch categories.
        if d2['Category'][r] and not any(row[1:]):
            category = d2['Category'][r]
            cat_quantity = 0
            cat_value = 0
            continue

        # Catch subcategories.
        if d2['Subcategory'][r] and not any(row[2:]):
            subcategory = d2['Subcategory'][r]
            subcat_quantity = 0
            subcat_value = 0
            continue

        # Catch end of category and update tallies.
        if d2['Category'][r] and d2['Category'][r] != 'GRAND TOTALS' and any(row[1:]):
            d2['Quantity'][r] = '%.2f' % cat_quantity
            d2['Value'][r] = '%.2f' % cat_value
            total_quantity += cat_quantity
            total_value += cat_value
            cat_quantity = 0
            cat_value = 0
            category = ''
            continue

        # Catch end of subcategory and update tallies.
        if d2['Subcategory'][r] and any(row[2:]):
            d2['Quantity'][r] = '%.2f' % subcat_quantity
            d2['Value'][r] = '%.2f' % subcat_value
            cat_quantity += subcat_quantity
            cat_value += subcat_value
            subcat_quantity = 0
            subcat_value = 0
            subcategory = ''
            continue

        # Catch items. Consolidate with any matching misplaced item, and tally.
        if d2['Item'][r]:
            found = False
            for m, mis in enumerate(misplaced):
                if subcategory == mis[0] and d2['Item'][r] == mis[3][item]:
                    found = True
                    d2['Quantity'][r] = str(float(d2['Quantity'][r]) + float(mis[3][quantity]))
                    d2['Value'][r] = str(float(d2['Value'][r]) + float(mis[3][value]))
                    break
            if found:
                del misplaced[m]
            subcat_quantity += float(d2['Quantity'][r])
            subcat_value += float(d2['Value'][r])
            d2['Quantity'][r] = '%.2f' % float(d2['Quantity'][r])
            d2['Value'][r] = '%.2f' % float(d2['Value'][r])

    # Grand total tallies should match new tallies.
    total_quantity = '%.2f' % total_quantity
    total_value = '%.2f' % total_value
    d2['Quantity'][-1] = '%.2f' % float(d2['Quantity'][-1])
    d2['Value'][-1] = '%.2f' % float(d2['Value'][-1])

    if total_quantity != d2['Quantity'][-1] or total_value != d2['Value'][-1]:
        ok = False
        print ('File "%s": Grand Totals do not match.' % (sales_raw.split('/')[-1]))

    d2['Quantity'][-1] = total_quantity
    d2['Value'][-1] = total_value

    # Sort the dictionary in order of subcategories OrderedDict.
    # First, get count of subcategories in each category.
    category_counts = {}
    for s in subcategories:
        if subcategories[s]['category']:
            if subcategories[s]['category'] in category_counts:
                category_counts[subcategories[s]['category']] += 1
            else:
                category_counts[subcategories[s]['category']] = 1

    category = ''
    subcategory = ''
    subcategory_count = 0
    for s in subcategories:
        if subcategories[s]['category']:
            for r in xrange(len(d2[cols[0]])):
                row = [d2[c][r] for c in cols]

                # Catch categories.
                if d2['Category'][r] == subcategories[s]['category'] and not any(row[1:]):
                    if not category:
                        category = d2['Category'][r]
                        [d3[c].append(d2[c][r]) for c in cols]
                    continue

                # Catch subcategories.
                if d2['Subcategory'][r] == s and not any(row[2:]):
                    if not subcategory:
                        subcategory = d2['Subcategory'][r]
                        [d3[c].append(d2[c][r]) for c in cols]
                    continue

                # Catch end of subcategory.
                if d2['Subcategory'][r] == s and any(row[2:]):
                    [d3[c].append(d2[c][r]) for c in cols]
                    subcategory_count += 1
                    subcategory = ''
                    continue

                # Catch end of category.
                if d2['Category'][r] == subcategories[s]['category'] and any(row[1:]):
                    if subcategory_count >= category_counts[subcategories[s]['category']]:
                        [d3[c].append(d2[c][r]) for c in cols]
                        subcategory_count = 0
                        category = ''
                    continue

                # Catch items.
                if d2['Item'][r] and category == subcategories[s]['category'] and subcategory == s:
                    [d3[c].append(d2[c][r]) for c in cols]

    [d3[c].append(d2[c][-1]) for c in cols]

    return d3, ok


def organize(sales_raw, subcategory, d, misplaced=[]):
    """Organize report subcategories into consistent items and order."""
    all = [[d[c][r].upper() for c in cols] for r in xrange(len(d[cols[0]]))]
    order = subcategories[subcategory]['order']
    extra = subcategories[subcategory]['extra']
    large = subcategories[subcategory]['large'][:]

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
                            #print ('File "%s": Moving Subcategory "%s" item "%s" to Subcategory "%s" as item "%s".' %
                            #       (sales_raw.split('/')[-1], subcategory, serving[item], o[3], placed[item]))

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
        sales_file_in = ''

    reports_dir_win = 'C:\\Users\\Malky\\Documents\\BD\\Reports\\sales'
    reports_dir_py = reports_dir_win.decode('utf-8').replace('\\','/')

    csv_files = glob(reports_dir_py + '/' + '*sales.csv')
    xls_files = glob(reports_dir_py + '/' + '*sales.xls')
    csv_files = [f.decode('utf-8').replace('\\','/') for f in csv_files]
    xls_files = [f.decode('utf-8').replace('\\','/') for f in xls_files]

    sales_files = [reports_dir_py + '/' + sales_file_in,] if sales_file_in else xls_files[:]

    count_tried = 0
    count_clean = 0
    for f in sales_files:
        count_tried += 1
        extension = f.split('.')[-1]
        if extension == 'xls':
            f_short = f.split('/')[-1]
            f_csv = ''.join(f.split('.')[:-1]) + '.csv'
            if f_csv not in csv_files:
                convert = subprocess.Popen(
                    ['soffice.exe', '--headless', '--invisible', '--convert-to', 'csv', '--outdir', reports_dir_win, reports_dir_win + '\\' + f_short],
                     executable='C:\Program Files (x86)\LibreOffice 4\program\soffice.exe')
                convert.wait()
                if convert.returncode != 0:
                    print 'Conversion of file "%s" to csv failed.' % f_short
                    return
            sales_raw = f_csv
        elif extension == 'csv':
            sales_raw = f
        else:
            print 'Input file [we]YYYYMM[DD]sales.csv or .xls required.'
            print 'Or to process all sales files, do not specify an input file.'
            return

        if sales_file_in:
            try:
                sales_raw = glob(sales_raw)[0]
            except:
                print 'Input file "%s" not found.' % sales_file_in
                return

        sales_clean = ''.join(sales_raw.split('.')[:-1]) + '_clean.csv'

        d, ok1 = parse_sales(sales_raw)
        ok2 = output_new(d, sales_clean) if d else False

        if ok1 and ok2:
            count_clean += 1
            print 'File "%s" OK.' % sales_clean.split('/')[-1]
        else:
            print 'File "%s" No Bueno.' % sales_clean.split('/')[-1]

    print '%d files clean out of %d tried.' % (count_clean, count_tried)


if __name__ == '__main__':
    main()
