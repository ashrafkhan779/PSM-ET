#!/usr/bin/env python3
"""
convert.py  —  turn the PSM working Excel file into data.json for the dashboard.
Usage:   python convert.py PSM_WORKING_FILE.xlsx      (writes data.json next to it)
Handles the current "Database" schema and the older schema automatically.
Requires: pandas, openpyxl   ->   pip install pandas openpyxl
"""
import sys, re, json, datetime as dt
import pandas as pd

def col(df,*names):
    for n in names:
        for c in df.columns:
            if str(c).strip().lower()==n.strip().lower(): return c
    return None
def to_date(v):
    if v is None: return None
    try:
        if pd.isna(v): return None
    except (TypeError,ValueError): pass
    if isinstance(v,(pd.Timestamp,dt.datetime,dt.date)):
        ts=pd.Timestamp(v); return None if pd.isna(ts) else ts.normalize()
    try:
        d=pd.to_datetime(v,errors='coerce'); return None if pd.isna(d) else d.normalize()
    except Exception: return None
def parse_tf(v):
    if v is None or (isinstance(v,float) and pd.isna(v)) or str(v).strip()=='' : return 'Unknown',None,None
    s=str(v); m=re.search(r'(\d+)',s); n=int(m.group(1)) if m else None
    if re.search(r'delay|overdue|late',s,re.I): return 'Delayed',None,n
    if re.search(r'remaining',s,re.I): return 'On Track',n,None
    return 'On Time',0,0
def old_stage(v):
    if v is None or pd.isna(v): return 'Scheduled'
    s=str(v).strip().lower()
    if 'ship' in s: return 'Shipped'
    if 'complete' in s: return 'Completed'
    if 'qc submitted' in s or 'parts ready' in s: return 'QC Submitted'
    if 'qc approved' in s or 'pickup' in s: return 'Ready for Pickup'
    if 'approved' in s: return 'Approved'
    return str(v).strip()
def state(cur,stage,tf):
    c=((str(cur) if cur is not None and not (isinstance(cur,float) and pd.isna(cur)) else '')+' '+(stage or '')).lower()
    if 'shipped' in c or 'delivered' in c or 'complete' in c: return 'Delivered - On Time'
    if tf=='Delayed': return 'Overdue'
    return 'Pending - On Track'

def main(path):
    xl=pd.ExcelFile(path); sheet=xl.sheet_names[0]
    df=pd.read_excel(path,sheet_name=sheet).dropna(how='all').reset_index(drop=True)
    g=lambda *n: df[col(df,*n)] if col(df,*n) is not None else pd.Series([None]*len(df))
    POc=g('PSM PO#'); recs=[]
    def ds(d): return None if d is None or pd.isna(d) else pd.Timestamp(d).strftime('%Y-%m-%d')
    def dm(d): return None if d is None or pd.isna(d) else pd.Timestamp(d).strftime('%b %Y')
    hasCur=col(df,'Current Status') is not None; hasRTS=col(df,'ET RTS Date') is not None
    for i in range(len(df)):
        if pd.isna(POc.iloc[i]): continue
        po=str(int(POc.iloc[i])) if str(POc.iloc[i]).replace('.0','').isdigit() else str(POc.iloc[i]).strip()
        tfcat,rem,delay=parse_tf(g('Time Frame').iloc[i])
        cur=g('Current Status').iloc[i] if hasCur else None
        rts=to_date(g('ET RTS Date').iloc[i]) if hasRTS else None
        if cur is not None and not (isinstance(cur,float) and pd.isna(cur)) and str(cur).strip()!='':
            stage=str(cur).strip(); rev=rts
        else:
            oldst=g('STATUS').iloc[i]; d=to_date(oldst)
            if d is not None: stage='Scheduled'; rev=d
            else: stage=old_stage(oldst); rev=rts
        delst=g('Delivery Status').iloc[i]
        stt=state(cur if hasCur else delst, stage, tfcat)
        oq=pd.to_numeric(g('Order Quantity').iloc[i],errors='coerce'); oq=0 if pd.isna(oq) else int(oq)
        val=pd.to_numeric(g('Total Value','Total Price').iloc[i],errors='coerce'); val=0.0 if pd.isna(val) else float(val)
        unit=val/oq if oq>0 else 0.0
        pend_raw=g('Still to be delivered (qty)').iloc[i]
        if stt=='Delivered - On Time': deliv,pendq=oq,0
        elif pend_raw is not None and not (isinstance(pend_raw,float) and pd.isna(pend_raw)) and not pd.isna(pd.to_numeric(pend_raw,errors='coerce')):
            pendq=int(round(float(pend_raw))); deliv=max(0,oq-pendq)
        else: pendq=oq; deliv=0
        pendv=round(pendq*unit,2)
        prod=g('Products').iloc[i]; prod='Unspecified' if prod is None or pd.isna(prod) or str(prod).strip()=='' else str(prod).strip()
        pod=to_date(g('PSM PO DATE').iloc[i]); ln=g('Line Item#').iloc[i]
        recs.append(dict(po=po,line=None if pd.isna(ln) else int(ln),
            material=str(g('Material#').iloc[i]).strip(),desc=str(g('Material Description').iloc[i]),
            product=prod,plant=str(g('Supplying Plant').iloc[i]),supplier=str(g('Supplier').iloc[i]),
            buyer=str(g('Buyer').iloc[i]).strip() if col(df,'Buyer') else '',
            category=str(g('PO Category').iloc[i]).strip() if col(df,'PO Category') else '',
            rmStatus=str(g('RM Status').iloc[i]).strip() if col(df,'RM Status') else '',
            remarks=str(g('Remarks').iloc[i]).strip() if col(df,'Remarks') and g('Remarks').iloc[i] is not None and not (isinstance(g('Remarks').iloc[i],float) and pd.isna(g('Remarks').iloc[i])) else '',
            state=stt,stage=stage,poDate=ds(pod),
            reqDate=ds(to_date(g('PSM Required Date','PSM Delivery Date').iloc[i])),
            expShip=ds(to_date(g('ET Promised Date','ET Expected Ship Date').iloc[i])),
            actShip=ds(to_date(g('Actual Ship date','ET Actual Ship date').iloc[i])),revShip=ds(rev),
            orderQty=oq,delivQty=deliv,pendQty=pendq,value=round(val,2),pendValue=pendv,
            delay=None if delay is None else int(delay),remaining=None if rem is None else int(rem),poMonth=dm(pod)))
    dts=sorted([r['poDate'] for r in recs if r['poDate']])
    meta=dict(supplier=next((r['supplier'] for r in recs if r['supplier']),'—'),
              plant=next((r['plant'] for r in recs if r['plant']),'—'),
              generated=dt.date.today().strftime('%d %b %Y'),poMin=dts[0] if dts else None,poMax=dts[-1] if dts else None)
    json.dump(dict(meta=meta,records=recs),open('data.json','w'))
    print(f'data.json written · {len(recs)} lines · as of {meta["generated"]}')

if __name__=='__main__':
    if len(sys.argv)<2: print('Usage: python convert.py <PSM_WORKING_FILE.xlsx>'); sys.exit(1)
    main(sys.argv[1])
