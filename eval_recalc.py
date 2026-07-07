import sys, json, warnings, re
warnings.filterwarnings("ignore")
import formulas, openpyxl
scoped = sys.argv[1]; prefix=sys.argv[2]
sheet = sys.argv[3] if len(sys.argv)>3 else 'T01'
xl = formulas.ExcelModel().loads(scoped).finish(circular=True)
sol = xl.calculate()
out={}
pat = re.compile(r"'?\[.*?\]"+re.escape(sheet)+r"'?!([A-Z]+\d+)$", re.I)
for k,v in sol.items():
    m = pat.search(k)
    if m:
        try: val=v.value[0,0]
        except Exception:
            try: val=v.value
            except Exception: val=str(v)
        out[m.group(1)]=val
# vul aan met constante cellen (niet-formule) rechtstreeks uit het bestand
wb = openpyxl.load_workbook(scoped, data_only=False)
ws = wb[sheet]
for row in ws.iter_rows():
    for cell in row:
        v=cell.value
        if v is None: continue
        if isinstance(v,str) and v.startswith('='): continue   # formule: uit sol
        if cell.coordinate not in out:
            out[cell.coordinate]=v
json.dump(out, open('/tmp/'+prefix+'_t01.json','w'), default=str)
print(f"{sheet} cellen (formule+constant):", len(out))
