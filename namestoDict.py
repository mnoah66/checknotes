
import openpyxl
from collections import defaultdict

trngfile = openpyxl.load_workbook('names.xlsx', read_only=True)
ws = trngfile['Sheet1']

example_dictionary = defaultdict(list)
for row in ws:
    a = row[0]
    b = row[1]
    if a.value:
        example_dictionary[a.value].append(b.value)

namesClean = {}

for names, values in example_dictionary.items():
    namesClean[names] = set(values)

for names, values in namesClean.items():
    if len(values) < 2:
        print(names + ' has less than two dates')

print(namesClean)

