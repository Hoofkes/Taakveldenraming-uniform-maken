import sys, json, warnings, re
warnings.filterwarnings("ignore")
import formulas
scoped = sys.argv[1]
xl = formulas.ExcelModel().loads(scoped).finish(circular=True)
sol = xl.calculate()
# verzamel T01 waarden
out={}
pat = re.compile(r"'?\[.*?\]T01'?!([A-Z]+\d+)$", re.I)
for k,v in sol.items():
    m = pat.search(k)
    if m:
        cell=m.group(1)
        try:
            val=v.value[0,0]
        except Exception:
            try: val=v.value
            except Exception: val=str(v)
        out[cell]=val
json.dump(out, open('/tmp/'+sys.argv[2]+'_t01.json','w'), default=str)
print("T01 cellen berekend:", len(out))
