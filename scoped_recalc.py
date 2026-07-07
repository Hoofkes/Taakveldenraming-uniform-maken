import sys, openpyxl, json, warnings
warnings.filterwarnings("ignore")
srcfile, outprefix = sys.argv[1], sys.argv[2]
# maak scoped kopie met alleen T01 + Indexering
wb = openpyxl.load_workbook(srcfile)
keep = {'T01','Indexering'}
for n in list(wb.sheetnames):
    if n not in keep:
        del wb[n]
scoped = f'/tmp/{outprefix}_scoped.xlsx'
wb.save(scoped)
print("scoped opgeslagen:", scoped, "sheets:", wb.sheetnames)
