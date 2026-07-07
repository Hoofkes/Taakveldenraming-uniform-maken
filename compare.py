import json, sys, re
snap = json.load(open('snapshot_fase0_waarden.json'))['T01']
calc = json.load(open('/tmp/'+sys.argv[1]+'_t01.json'))
def num(x):
    try: return float(x)
    except: return None
mism=[]; checked=0
for cell,sv in snap.items():
    n_s=num(sv)
    if n_s is None: continue          # alleen numerieke cellen vergelijken
    cv=calc.get(cell)
    n_c=num(cv)
    checked+=1
    if n_c is None:
        mism.append((cell,sv,cv,'niet berekend')); continue
    if abs(n_s-n_c)>1.0:
        mism.append((cell,round(n_s,2),round(n_c,2),round(n_c-n_s,2)))
print(f"Numerieke cellen gecontroleerd: {checked}, afwijkingen >EUR1: {len(mism)}")
for m in mism[:40]:
    print("  ",m)
