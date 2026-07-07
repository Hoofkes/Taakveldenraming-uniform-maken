import zipfile, re, json, shutil
SRC='04_Taakvelden_2027_backup_00_origineel.xlsx'
DST='/tmp/T01_geuniformeerd_fidelity.xlsx'
SHEET='xl/worksheets/sheet12.xml'
changes=json.load(open('/tmp/t01_changes.json'))

def esc(s): return str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
def col_idx(a):
    m=re.match(r'([A-Z]+)',a); n=0
    for ch in m.group(1): n=n*26+(ord(ch)-64)
    return n
def tofloat(x):
    try: return float(x)
    except: return None
def numstr(v):
    f=tofloat(v)
    if f is None: return None
    if f==int(f) and abs(f)<1e15: return str(int(f))
    return repr(f)

def build(addr, s_attr, ch):
    k=ch['kind']
    if k=='text':
        return f'<c r="{addr}"{s_attr} t="inlineStr"><is><t xml:space="preserve">{esc(ch["v"])}</t></is></c>'
    if k=='number':
        return f'<c r="{addr}"{s_attr}><v>{numstr(ch["v"])}</v></c>'
    if k=='formula':
        f=esc(ch['f']); ns=numstr(ch.get('v'))
        if ns is None:
            # geen numerieke waarde -> als string-resultaat indien tekst, anders laat Excel rekenen
            v=ch.get('v')
            if v not in (None,'None',''):
                return f'<c r="{addr}"{s_attr} t="str"><f>{f}</f><v>{esc(v)}</v></c>'
            return f'<c r="{addr}"{s_attr}><f>{f}</f></c>'
        return f'<c r="{addr}"{s_attr}><f>{f}</f><v>{ns}</v></c>'
    raise ValueError(k)

with zipfile.ZipFile(SRC) as z:
    names=z.namelist()
    data={n:z.read(n) for n in names}

xml=data[SHEET].decode('utf-8')
applied=0; inserted=0
for ch in changes:
    addr=ch['addr']
    pat=re.compile(r'<c r="%s"((?: [^>]*?)?)(?:/>|>.*?</c>)' % re.escape(addr), re.S)
    m=pat.search(xml)
    s_attr=''
    if m:
        sm=re.search(r'\bs="(\d+)"', m.group(1))
        if sm: s_attr=f' s="{sm.group(1)}"'
        xml=xml[:m.start()]+build(addr,s_attr,ch)+xml[m.end():]
        applied+=1
    else:
        # cel bestaat niet -> invoegen in juiste rij op kolomvolgorde
        rn=re.match(r'[A-Z]+(\d+)',addr).group(1)
        ci=col_idx(addr)
        rowpat=re.compile(r'(<row r="%s"[^>]*>)(.*?)(</row>)' % rn, re.S)
        rm=rowpat.search(xml)
        newc=build(addr,'',ch)
        if rm:
            body=rm.group(2)
            cells=re.findall(r'<c r="[A-Z]+\d+".*?(?:/>|</c>)', body, re.S)
            pos=len(cells)
            for i,cxml in enumerate(cells):
                a=re.search(r'r="([A-Z]+\d+)"',cxml).group(1)
                if col_idx(a)>ci: pos=i; break
            cells.insert(pos,newc)
            xml=xml[:rm.start(2)]+''.join(cells)+xml[rm.end(2):]
        else:
            # rij bestaat niet -> invoegen in sheetData op rijvolgorde
            newrow=f'<row r="{rn}">{newc}</row>'
            allrows=list(re.finditer(r'<row r="(\d+)"',xml))
            insert_at=None
            for rmm in allrows:
                if int(rmm.group(1))>int(rn): insert_at=rmm.start(); break
            if insert_at is None:
                insert_at=xml.find('</sheetData>')
            xml=xml[:insert_at]+newrow+xml[insert_at:]
        inserted+=1
data[SHEET]=xml.encode('utf-8')

# fullCalcOnLoad in workbook.xml
wb=data['xl/workbook.xml'].decode('utf-8')
if '<calcPr' in wb:
    wb=re.sub(r'<calcPr([^>]*?)/>', lambda m: '<calcPr'+re.sub(r'\s*fullCalcOnLoad="[^"]*"','',m.group(1))+' fullCalcOnLoad="1"/>', wb, count=1)
else:
    wb=wb.replace('</workbook>','<calcPr fullCalcOnLoad="1"/></workbook>')
data['xl/workbook.xml']=wb.encode('utf-8')

with zipfile.ZipFile(DST,'w',zipfile.ZIP_DEFLATED) as zout:
    for n in names:
        zout.writestr(n, data[n])
print(f'toegepast(vervangen)={applied}, ingevoegd={inserted}, doel={DST}')
