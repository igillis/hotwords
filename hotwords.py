# You'll need to install these libraries locally
from xlrd import open_workbook
from fuzzywuzzy import fuzz

# Change this to point to the correct location in your file system
wb = open_workbook('../Downloads/hotwords_simple.xlsx')

resale = wb.sheet_by_name('Resale')
national_sample = wb.sheet_by_name('National Sample')

# This assumes you have the contract paragraphs at column 1 
# and the hotwords at column 2 (zero indexed)
#
# This is _slow_. Might have to run overnight for large data sets
match_count = 0
for ns_row in range(national_sample.nrows):
    match = 0
    # You can make this more strict if you want. 70
    # seems to be teh sweet spot though
    max_ratio = 70
    for r_row in range(resale.nrows):
        ns_val = national_sample.cell(ns_row, 1)
        r_val = resale.cell(r_row, 1)

        ratio = fuzz.ratio(ns_val, r_val)
        if ratio > max_ratio:
          match = r_row
          max_ratio = ratio
    if match:
        match_count += 1
        print '-----------------------------------'
        print 'Match at row ', ns_row, ' with ratio ', max_ratio
        print national_sample.cell(ns_row, 1).value
        print
        print resale.cell(match, 1).value
        print
        print resale.cell(match, 2).value
print match_count

# Should be easy to do the other kind of matching (looking for
# hotwords) with the xlrd library
