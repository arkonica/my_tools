#!/usr/bin/env python

import argparse
import xlrd
import simplejson as json
from collections import OrderedDict

parser = argparse.ArgumentParser()
parser.add_argument("source", help="path to source file")
parser.add_argument("-target", help="path to target file")
parser.add_argument("-nh", "--no-header", action="store_true", help="source file does not contain header row")
parser.add_argument("-c", "--compact-format", action="store_true", help="dump json file using compact format")
parser.add_argument("-sh", "--sheet-index", type=int, default=0, help="specify target sheet index")
args = parser.parse_args()

source_path = args.source
target_path = args.target if args.target else source_path + ".json"
has_header = False if args.no_header else True
compact_format = args.compact_format
sheet_index = args.sheet_index

print "Source: " + source_path
print "Target: " + target_path
print "Has header row: " + str(has_header)
print "Use Compact format: " + str(compact_format)
print "Target sheet index: " + str(sheet_index)
print " ... converting ..."

wb = xlrd.open_workbook(source_path)
sh = wb.sheet_by_index(sheet_index)
item_list = []
headers = []

first_row = 1 if has_header else 0

if has_header:
    header_row = sh.row_values(0)
    for header in header_row:
        headers.append(header)
else:
    for col_index in range(len(sh.row_values(first_row))):
        headers.append("column_" + str(col_index))

for row_index in range(first_row, sh.nrows):
    item = OrderedDict()
    row_values = sh.row_values(row_index)
    for col_index in range(len(row_values)):
        item[headers[col_index]] = row_values[col_index]
    
    item_list.append(item)

json_content = json.dumps(item_list) if compact_format else json.dumps(item_list, indent=4)
with open(target_path, "w") as json_file:
    json_file.write(json_content)

print "Finished: " + target_path