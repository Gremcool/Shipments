"""Merge the Excel tracker into the current app export and write
seed/Review_and_changes.xlsx (review sheet only - the app applies the data itself).
Run:  python scripts/prepare_update_file.py
"""
import pandas as pd, re, warnings, sys, json
from datetime import datetime, timedelta
warnings.filterwarnings('ignore')
import os; ROOT=os.path.dirname(os.path.dirname(os.path.abspath(__file__))); sys.path.insert(0,ROOT); os.chdir(ROOT); os.makedirs('seed',exist_ok=True)
import db as D   # project db.py (for pick_existing)

TRK='original_data/Shipment_tracker_last_update.xlsx'
EXP='original_data/RMS_Shipment_Tracker_2026-2027.xlsx'
APP_COLS=list(pd.read_excel(EXP,nrows=0).columns)
PO,SUP,CARGO,INV,VAL,CUR,QTY,MODE,INC,CONT,CAR,BL,POD,ETAP,APORT,ETAR,ARMS,DOC,STAT,DTC,DEAD,DEL,WARN,MAH=APP_COLS
DATES=[ETAP,APORT,ETAR,ARMS,DOC,DEAD]
EXTRA_UI=['Import license stage','Missing cargo document','Cargo dispatched','Clearing agent','Arrival notice ready','Actual arrival in Kigali','Remarks']
def infer_mode(pod,carrier,bl):
    t=(nz(pod)+' '+nz(carrier)).lower()
    if re.search(r'\b(kia|kai|kigali (international )?airport|airport)\b',t) or re.search(r'rwandair|ethiopian|kenya airways|emirates|qatar|turkish|lufthansa|klm|dhl',t): return 'AIR'
    if re.search(r'gatuna|rusumo|kagitumba|nemba|cyanika|truck|road',t): return 'ROAD'
    if re.search(r'dar ?es|mombasa|djibouti|durban|beira|maersk|msc|cma|hapag|one line|evergreen|cosco',t): return 'SEA'
    if re.fullmatch(r'\d{3}[- ]?\d{7,8}',nz(bl).replace(' ','')): return 'AIR'
    return ''
COMPUTED=[DTC,DEL,WARN]
PH={'','tba','pending','na','n/a','none','nan','null','tbc','-'}
def nz(x):
    if x is None or (isinstance(x,float) and pd.isna(x)) or x is pd.NaT: return ''
    return re.sub(r'[ \t]+',' ',str(x)).strip()
def isph(x): return nz(x).lower() in PH
def po_key(p):
    q=re.sub(r'\s+',' ',nz(p)).upper()
    return 'NO PO' if q in ('','NO PO','NOPO','NO P.O','NA','N/A','TBA','NONE','PENDING','-') else q
def pdate(x):
    if isinstance(x,(datetime,pd.Timestamp)): return x.strftime('%Y-%m-%d')
    s=nz(x)
    if s.lower() in PH: return ''
    if re.fullmatch(r'\d{5}',s): return (datetime(1899,12,30)+timedelta(days=int(s))).strftime('%Y-%m-%d')
    if re.match(r'^\d{4}-\d{1,2}-\d{1,2}',s):   # ISO text: never day/month-swap it
        t=pd.to_datetime(s[:10],format='%Y-%m-%d',errors='coerce')
    else:
        t=pd.to_datetime(s,errors='coerce',dayfirst=True)
    return '' if pd.isna(t) else t.strftime('%Y-%m-%d')
def pnum(x):
    s=nz(x)
    if s.lower() in PH: return None
    s=re.sub(r'[^\d,.\-]','',s)
    if re.search(r',\d{1,2}$',s): s=s.replace('.','').replace(',','.')
    else: s=s.replace(',','')
    try: return float(s)
    except: return None
def pmode(x):
    s=nz(x).lower()
    if 'sea' in s: return 'SEA'
    if 'air' in s: return 'AIR'
    if 'road' in s or 'truck' in s: return 'ROAD'
    return ''
def pinc(x):
    s=nz(x)
    if isph(s): return ''
    u=s.upper()
    if re.match(r'^EX[\s\-]?W',u): return 'EX-W'
    return u
def pcur(x):
    u=nz(x).upper()
    return 'EUR' if u in ('EURO','EUR','€') else u

# ---------- load ----------
E=pd.read_excel(EXP,dtype=object,keep_default_na=False,na_values=[''])
for c in DATES: E[c]=E[c].map(lambda v: pdate(v))
rows=[]   # working pool: dict app-col -> value, plus meta
for i,r in E.iterrows():
    d={c:('' if pd.isna(r[c]) else r[c]) for c in APP_COLS}
    d['__orig']=dict(d); d['__src']='DB'; d['__new']=False; d['__changes']=[]; d['__extra']={}
    for c in APP_COLS:
        if c==VAL: d[c]=float(d[c]) if d[c]!='' else ''
        elif c not in DATES: d[c]=nz(d[c]) if c!=CARGO else ('' if d[c]=='' else str(d[c]).strip())
    for e_ in EXTRA_UI: d[e_]=''
    d['__orig']={k:d[k] for k in APP_COLS+EXTRA_UI}
    d['__raw']={c:(None if pd.isna(r[c]) else r[c]) for c in APP_COLS}
    rows.append(d)

def load(sheet,hdr):
    t=pd.read_excel(TRK,sheet_name=sheet,header=hdr,dtype=object,keep_default_na=False,na_values=[''])
    t.columns=[re.sub(r'\s+',' ',str(c)).strip() for c in t.columns]
    t['__xlrow']=range(hdr+2,hdr+2+len(t))
    t=t[t.drop(columns=['__xlrow']).apply(lambda r: any(nz(v)!='' for v in r),axis=1)]
    return t[~t['Purchase Order'].map(nz).str.upper().str.contains('LEGEND')]
sheets=[('Delivered-new,_current_sheet',0),('upcoming_&_at_the_port',1)]
def col(t,prefix): return next((c for c in t.columns if c.lower().startswith(prefix.lower())),None)

def derive_status(tr_text,dispatched,arms,aport):
    tx=tr_text.lower()
    if arms or re.search(r'\breceived\b.*(warehouse|rms)|delivered to (the )?warehouse',tx): return 'Received in Warehouse'
    if 'clearance' in tx or 'cleared' in tx: return 'Under Clearance'
    if aport: return 'Arrived at Port'
    d=nz(dispatched).lower()
    if d.startswith('y'): return 'In Transit'
    if d.startswith('n'): return 'Awaiting Departure'
    return ''

review=[]; changes=[]; extras=[]
for sheet,hdr in sheets:
    t=load(sheet,hdr)
    docc=col(t,'DATE OF SHARING'); remc=col(t,'Additional remarks'); licc=col(t,'Import license'); agent=col(t,'Clearing agent') or col(t,'Forwarder')
    for _,r in t.iterrows():
        po=nz(r['Purchase Order']); po='NO PO' if po_key(po)=='NO PO' else po; xl=f"{sheet} row {int(r['__xlrow'])}"
        new={PO:po,SUP:nz(r.get('Vendor')),CARGO:str(r.get('Cargo description',r.get('Cargo Description',''))).strip() if not pd.isna(r.get('Cargo description',r.get('Cargo Description',''))) else '',
             INV:nz(r.get('Invoice number')),VAL:pnum(r.get('Inv value')),CUR:pcur(r.get('Currency')),QTY:nz(r.get('Quantity')),
             MODE:pmode(r.get('Shipping mode')),INC:pinc(r.get('Incoterm')),CONT:nz(r.get('No of container(s)')),CAR:nz(r.get('Carrier')),
             BL:nz(r.get('BILL OF LADING OR AIRWAY BILL')),POD:nz(r.get('POD')),
             ETAP:pdate(r.get('ETA POD')),APORT:pdate(r.get('Actual arrival at port')),ETAR:pdate(r.get('Cargo ETA RMS')),
             ARMS:pdate(r.get('Actual arrival at RMS')),DOC:pdate(r.get(docc)),DEAD:pdate(r.get('Deadline to clear the shipment'))}
        for k in (CARGO,SUP): 
            if isph(new[k]): new[k]=''
        if not new[MODE]: new[MODE]=infer_mode(new[POD],new[CAR],new[BL])
        raw_mode=nz(r.get('Shipping mode')); raw_inc=nz(r.get('Incoterm')); raw_bl=nz(r.get('BILL OF LADING OR AIRWAY BILL'))
        shifted=len(raw_mode)>25 or len(raw_inc)>25 or len(raw_bl)>45
        if shifted:
            for k_ in list(new):
                if k_ not in (PO,SUP,CARGO,INV,VAL,CUR,QTY): new[k_]=''
            review.append((po,raw_bl[:40],'Row cells are shifted in the tracker - partly imported','Only PO, vendor, cargo, invoice, value, currency and quantity were imported (those cells are in the right place). Mode, dates, BL, carrier etc. sit in the wrong columns; fix the row in the tracker.',xl))
        stt=nz(r.get('Current status')); rem=nz(r.get(remc)) if remc else ''
        text=(rem if stt.startswith('Missing document') else (stt+' '+rem)).strip(); disp=r.get('Cargo dispatched?')
        st='' if shifted else derive_status(text,disp,new[ARMS],new[APORT]); new[STAT]=st
        extra={'Import license stage':nz(r.get(licc)),'Missing cargo document?':nz(r.get('Missing cargo document?')),'Cargo dispatched?':nz(disp),
               'Clearing agent / Forwarder':nz(r.get(agent)) if agent else '','Arrival notice ready?':nz(r.get('Arrrival notice ready?')),
               'Actual arrival in Kigali':pdate(r.get('Actual arrival in Kigali')),'Tracker status text':text,'Additional remarks':nz(r.get(remc)) if remc else ''}
        gen=text.lower().startswith('missing document')
        xui={'Import license stage':extra['Import license stage'],'Missing cargo document':extra['Missing cargo document?'],'Cargo dispatched':extra['Cargo dispatched?'],
             'Clearing agent':extra['Clearing agent / Forwarder'],'Arrival notice ready':extra['Arrival notice ready?'],'Actual arrival in Kigali':extra['Actual arrival in Kigali'],
             'Remarks':'' if gen else text}
        if shifted: xui={k_:'' for k_ in xui}
        for k_,v_ in list(xui.items()):
            if k_=='Actual arrival in Kigali': continue
            if isph(v_) or v_ in ('-',): xui[k_]=''
        # --- match ---
        cands=[(i,x[BL],x[INV]) for i,x in enumerate(rows) if po_key(x[PO])==po_key(po) and (po_key(po)!='NO PO' or (new[INV] and not isph(new[INV]) and nz(x[INV])==new[INV]))]
        mi=D.pick_existing(cands,new[BL],new[INV]) if cands else None
        if mi is None:
            d={c:'' for c in APP_COLS}
            for c in APP_COLS:
                if c in new and new[c] is not None: d[c]=new[c]
            for c in DATES: d[c]=new.get(c,'')
            if d[VAL] in ('',None): d[VAL]=''
            if not d[STAT]: d[STAT]='In Transit'
            for e_ in EXTRA_UI: d[e_]=xui[e_]
            d.update({'__orig':None,'__src':xl,'__new':True,'__changes':[],'__extra':extra})
            rows.append(d)
            if PO and re.search(r'&|,|/| and ',po,re.I): review.append((po,new[BL],'Compound PO in one cell','Imported as typed. Matches only if typed identically next time; consider one row per PO.',xl))
            if re.search(r'\b(RFDA|THE following|delivered|inspected)\b',new[BL],re.I) or len(new[BL])>40: review.append((po,new[BL][:60],'BL/AWB cell contains free text','Looks like a note, not a BL/AWB number. Check the tracker.',xl))
            continue
        x=rows[mi]; x['__extra']=extra; x['__src']=xl
        for c in APP_COLS:
            if c in COMPUTED or c==MAH or c==PO: continue
            nv=new.get(c)
            if c==VAL:
                if nv is None: continue
            elif nv in ('',None) or (c not in DATES and isph(nv)): continue
            if c==BL and re.search(r'\b(RFDA|THE following)\b',str(nv),re.I): 
                review.append((po,str(nv)[:60],'BL/AWB cell contains free text','Not applied to existing row.',xl)); continue
            old=x[c]
            same=(abs(float(old)-float(nv))<0.005) if (c==VAL and old!='') else (nz(old)==nz(nv) if c!=CARGO else re.sub(r'\s+',' ',str(old))==re.sub(r'\s+',' ',str(nv)))
            if not same:
                changes.append((po,x[BL] or new[BL],c,old,nv,xl)); x[c]=nv
                if c==VAL and old!='' and float(old)>0 and nv>0 and max(float(old),nv)/min(float(old),nv)>1.5:
                    review.append((po,x[BL] or new[BL],'Invoice value changed a lot',('App had %s, tracker has %s - looks like a decimal-point slip in the app (tracker value applied).'%(old,nv)) if any(abs(max(float(old),nv)/min(float(old),nv)-k)<k*0.02 for k in (10,100,1000,10000)) else ('App had %s, tracker has %s - tracker value applied; confirm which is correct.'%(old,nv)),xl))
        for e_ in EXTRA_UI:
            v_=xui[e_]
            if v_ and v_!=x[e_]:
                changes.append((po,x[BL],e_,x[e_],v_,xl)); x[e_]=v_
        # deadline: tracker value wins; else follow doc date + 2d if doc date known
        if not new[DEAD] and x[DOC]:
            auto=(datetime.strptime(x[DOC],'%Y-%m-%d')+timedelta(days=2)).strftime('%Y-%m-%d')
            if x[DEAD]!=auto:
                changes.append((po,x[BL],DEAD,x[DEAD],auto,xl+' (auto: doc date + 2 days)')); x[DEAD]=auto
        if st and st!=x[STAT]:
            changes.append((po,x[BL],STAT,x[STAT],st,xl+' (derived from tracker dates/status text)')); x[STAT]=st

# ---------- classify output ----------
out=[]; 
for x in rows:
    if x['__new']: out.append(x)
    elif any(x[c]!=x['__orig'][c] for c in APP_COLS+EXTRA_UI if c not in COMPUTED): out.append(x)
n_new=sum(1 for x in out if x['__new']); n_upd=len(out)-n_new
# review: flags on new rows duplicating a PO already present (possible duplicates)
dup=pd.Series([x[PO] for x in rows]).value_counts()
for x in out:
    if x['__new'] and dup[x[PO]]>1 and any((not y['__new']) and y[PO]==x[PO] for y in rows):
        review.append((x[PO],x[BL],'New row for a PO that already exists','Added because BL/invoice did not match any existing row for this PO. Confirm it is a separate shipment, not a duplicate.',x['__src']))
    if x[MODE]=='' and x['__new']: review.append((x[PO],x[BL],'Shipping mode not recognised','Left blank.',x['__src']))
    if x['__new'] and x[VAL]=='': review.append((x[PO],x[BL],'Invoice value missing/TBA','Left blank (not 0).',x['__src']))
    if x['__new'] and re.search(r'^(TBA|)$',x[BL],re.I) and dup[x[PO]]>1: review.append((x[PO],x[BL],'Several rows share this PO with no BL yet','Distinct invoice numbers are used to tell them apart on later uploads.',x['__src']))
pd.to_pickle({'EXTRA_UI':EXTRA_UI,'out':out,'rows':rows,'review':review,'changes':changes,'APP_COLS':APP_COLS},'seed/_merged.pkl')
print('DB rows',len(E),'| output upload rows',len(out),'new',n_new,'updated',n_upd,'| changes',len(changes),'| review',len(review))

# ---- step 2: write the workbook ----
import pandas as pd, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
m=pd.read_pickle('seed/_merged.pkl'); EX=m['EXTRA_UI']; out=m['out']; review=list(m['review']); changes=m['changes']; A=m['APP_COLS']
PO,INV,BL,VAL,STAT=A[0],A[3],A[11],A[4],A[18]
COMPUTED=[A[19],A[21],A[22]]
# extra review items: rows with no real PO
for x in out:
    if re.fullmatch(r'(TBA|NO PO|NA|N/A|NONE|PENDING)?',x[PO].strip(),re.I):
        review.append((x[PO],x[BL],'No real PO number','Imported with the PO text as typed. Rows like this are matched on invoice number only; give it a PO when you have one.',x['__src']))
# dedupe review
seen=set(); rv=[]
for r in review:
    k=(r[0],r[1],r[2])
    if k in seen: continue
    seen.add(k); rv.append(r)
review=sorted(rv,key=lambda r:(r[2],r[0]))
F=Font(name='Arial',size=10); H=Font(name='Arial',size=10,bold=True,color='FFFFFF'); fill=PatternFill('solid',fgColor='117A8B')
wb=Workbook(); ws=wb.active; ws.title='Shipments'
def sheet(ws,header,data,widths=None):
    ws.append(header)
    for r in data: ws.append(r)
    for c in ws[1]: c.font=H; c.fill=fill; c.alignment=Alignment(wrap_text=True,vertical='center')
    for row in ws.iter_rows(min_row=2):
        for c in row: c.font=F; c.alignment=Alignment(vertical='top',wrap_text=False)
    ws.freeze_panes='B2'
    for i,h in enumerate(header,1):
        ws.column_dimensions[get_column_letter(i)].width=(widths or {}).get(i, min(34,max(14,len(str(h))+3)))
    ws.auto_filter.ref=ws.dimensions
data=[]
for x in sorted(out,key=lambda x:(x['__new'],x[PO])):
    row=[]
    for c in A+EX:
        v=x[c]
        if (not x['__new']) and x.get('__raw') is not None and c in x['__raw'] and x[c]==x['__orig'][c] and x['__raw'][c] is not None: v=x['__raw'][c]
        if c in COMPUTED: v=''
        if c==VAL and v=='': v=None
        row.append(v if v!='' else None)
    data.append(row)
sheet(ws,A+EX,data,{3:40,2:26,19:20})
ws2=wb.create_sheet('Changes (existing rows)')
sheet(ws2,['PO','BL/AWB','Field','Old value (app)','New value (tracker)','Where it came from'],[[str(a) if not isinstance(a,(int,float)) else a for a in (c[0],c[1])]+[c[2],c[3] if c[3]!='' else None,c[4] if c[4]!='' else None,c[5]] for c in changes],{1:18,2:22,3:30,4:30,5:30,6:42})
ws3=wb.create_sheet('Review before upload')
sheet(ws3,['PO','BL/AWB','Issue','What was done / what to check','Tracker location'],[list(r) for r in review],{1:20,2:26,3:38,4:80,5:42})
for w in (ws2,ws3): w.sheet_view.zoomScale=100
wb.save('seed/Review_and_changes.xlsx')
print('saved; Shipments rows',len(data),'changes',len(changes),'review',len(review))
import collections; print(collections.Counter(r[2] for r in review))


