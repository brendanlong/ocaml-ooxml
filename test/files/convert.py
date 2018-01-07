#!/usr/bin/env python3
import csv
import os
import subprocess
import sys

import xlrd


files = sys.argv[1:]
for file in files:
    name, _ = file.rsplit(".", maxsplit=1)
    workbook = xlrd.open_workbook(file)
    for sheet in workbook.sheets():
        with open("%s_%s.csv" % (name, sheet.name), "w") as f:
            writer = csv.writer(f)
            for row in sheet.get_rows():
                values = [cell.value for cell in row]
                writer.writerow(values)
    print("; `Skip, \"%s\", [ %s ]" % (name, "; ".join(["\"%s\"" % sheet.name for sheet in workbook.sheets()])))
