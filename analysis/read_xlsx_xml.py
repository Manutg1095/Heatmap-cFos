import zipfile, xml.etree.ElementTree as ET, re
import pandas as pd
from pathlib import PurePosixPath
NS={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
RNS='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
RELNS={'p':'http://schemas.openxmlformats.org/package/2006/relationships'}
def col_idx(cellref):
    m=re.match(r'([A-Z]+)',cellref)
    n=0
    for ch in m.group(1):n=n*26+ord(ch)-64
    return n-1
class XLSX:
    def __init__(self,path):
        self.z=zipfile.ZipFile(path)
        try:
            root=ET.fromstring(self.z.read('xl/sharedStrings.xml'))
            self.ss=[''.join(t.text or '' for t in si.iter('{%s}t'%NS['m'])) for si in root]
        except KeyError:self.ss=[]
        wb=ET.fromstring(self.z.read('xl/workbook.xml'))
        rel=ET.fromstring(self.z.read('xl/_rels/workbook.xml.rels'))
        mp={e.attrib['Id']:e.attrib['Target'] for e in rel.findall('p:Relationship',RELNS)}
        self.sheets={}
        for s in wb.findall('m:sheets/m:sheet',NS):
            rid=s.attrib[RNS];t=mp[rid].lstrip('/')
            if not t.startswith('xl/'):t='xl/'+t
            self.sheets[s.attrib['name']]=t
    def read(self,name):
        root=ET.fromstring(self.z.read(self.sheets[name])); rows=[];maxc=0
        for r in root.findall('.//m:sheetData/m:row',NS):
            vals={}
            for c in r.findall('m:c',NS):
                j=col_idx(c.attrib['r']);typ=c.attrib.get('t');v=c.find('m:v',NS)
                if typ=='inlineStr':
                    x=''.join(t.text or '' for t in c.findall('.//m:t',NS))
                elif v is None:x=None
                else:
                    raw=v.text
                    if typ=='s':x=self.ss[int(raw)]
                    elif typ=='b':x=(raw=='1')
                    elif typ in ('str','e'):x=raw
                    else:
                        try:
                            f=float(raw);x=int(f) if f.is_integer() else f
                        except:x=raw
                vals[j]=x;maxc=max(maxc,j)
            rows.append(vals)
        mat=[[r.get(i) for i in range(maxc+1)] for r in rows]
        if not mat:return pd.DataFrame()
        headers=[]
        seen={}
        for i,h in enumerate(mat[0]):
            h=str(h) if h is not None else f'Unnamed_{i}'
            if h in seen:seen[h]+=1;h=f'{h}_{seen[h]}'
            else:seen[h]=0
            headers.append(h)
        return pd.DataFrame(mat[1:],columns=headers)
