"""Build privacy-safe decision data from a LenDenClub manual-lending report."""
import json, re, sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path
import openpyxl

BANDS = [(700,712),(713,724),(725,736),(737,748),(749,760),(761,772),(773,784),(785,999)]
FIELDS = ['order_id','loan_id','date','amount','repayment','start','illustrative','received','principal','interest','reported_fee','pnl','npa','status','close_date','dpd','rate','tenure','score']

def num(v):
    return float(re.sub(r'[^0-9.\-]', '', str(v or 0)) or 0)
def date(v):
    if hasattr(v, 'date'): return v.date()
    for fmt in ('%d/%m/%Y','%Y-%m-%d'):
        try: return datetime.strptime(str(v), fmt).date()
        except ValueError: pass
    return None
def fee_rate(tenure, disbursed):
    if tenure in (2,3): return .01
    if tenure == 4: return .03 if disbursed and disbursed >= datetime(2026,4,1).date() else .023
    return {5:.025, 6:.03, 12:.06}.get(tenure, 0)
def band(score):
    for low, high in BANDS:
        if low <= score <= high: return f'{low}+' if high == 999 else f'{low}–{high}'
    return 'Unknown'
def safe(n, d): return round(n / d, 4) if d else 0
def money(n): return round(n, 2)
def xirr(flows):
    """Annual IRR from dated aggregate cash flows; returns a percentage or None."""
    if not flows or not any(v < 0 for _,v in flows) or not any(v > 0 for _,v in flows): return None
    start = min(d for d,_ in flows)
    def npv(rate): return sum(v / (1 + rate) ** ((d-start).days / 365) for d,v in flows)
    low, high = -.999, 1000.0
    if npv(low) * npv(high) > 0: return None
    for _ in range(120):
        mid = (low + high) / 2
        if npv(low) * npv(mid) <= 0: high = mid
        else: low = mid
    return round(mid * 100, 2)

def read(path):
    sheet = openpyxl.load_workbook(path, read_only=True, data_only=True).active
    rows = []
    for values in list(sheet.values)[20:]:
        if not values[1]: continue
        x = dict(zip(FIELDS, values))
        for key in ('amount','received','principal','interest','reported_fee','npa','dpd','rate','tenure','score'): x[key] = num(x[key])
        x['date'] = date(x['date']); x['status'] = str(x['status'] or '').upper(); x['band'] = band(x['score'])
        x['close_date'] = date(x['close_date'])
        x['fee'] = x['principal'] * fee_rate(x['tenure'], x['date'])
        x['default_loss'] = max(0, x['amount'] - x['principal']) if x['status'] in ('CLOSED','NPA') or x['npa'] else 0
        x['net_income'] = x['interest'] - x['fee'] - x['default_loss']
        x['complete'] = x['status'] in ('CLOSED','NPA')
        rows.append(x)
    return rows

def aggregate(rows, label):
    out = {'label':label, 'loans':len(rows), 'lent':money(sum(x['amount'] for x in rows)), 'interest':money(sum(x['interest'] for x in rows)), 'fees':money(sum(x['fee'] for x in rows)), 'default_loss':money(sum(x['default_loss'] for x in rows)), 'net':money(sum(x['net_income'] for x in rows))}
    out['net_per_1000'] = round(safe(out['net'],out['lent'])*1000,1); out['net_roi'] = round(safe(out['net'],out['lent'])*100,2)
    out['default_rate'] = round(safe(sum(1 for x in rows if x['default_loss'] > 0),len(rows))*100,2)
    out['mean_rate'] = round(safe(sum(x['rate'] for x in rows),len(rows)),2)
    flows = []
    for x in rows:
        end = x['close_date'] or x['date']
        cash = x['principal'] + x['interest'] - x['fee']
        if x['date'] and end and end > x['date'] and cash > 0:
            flows.extend([(x['date'], -x['amount']), (end, cash)])
    out['annual_xirr'] = xirr(flows)
    return out

def build(rows):
    completed = [x for x in rows if x['complete']]
    groups = {}
    for tenure in sorted(set(int(x['tenure']) for x in completed)):
        for lo, hi in BANDS:
            name = f'{lo}+' if hi == 999 else f'{lo}–{hi}'; chunk = [x for x in completed if x['tenure']==tenure and x['band']==name]
            if chunk:
                g = aggregate(chunk, f'{tenure}m · {name}'); g.update({'tenure':tenure,'band':name})
                g['verdict'] = 'Low evidence' if g['loans'] < 10 else ('Avoid' if g['net'] <= 0 or g['default_rate'] >= 10 else ('Fund' if g['net_roi'] > 3 else 'Caution'))
                groups[f'{tenure}|{name}'] = g
    portfolio = aggregate(completed, 'Completed loans')
    by_tenure = [aggregate([x for x in completed if x['tenure']==t], f'{t} months') | {'tenure':t} for t in sorted(set(x['tenure'] for x in completed))]
    by_band = [aggregate([x for x in completed if x['band']==b], b) | {'band':b} for b in [f'{a}+' if z==999 else f'{a}–{z}' for a,z in BANDS]]
    chart_specs = []
    metrics = [('annual_xirr','Annualised net XIRR %'),('net_per_1000','Net ₹ / ₹1,000'),('net_roi','Net ROI %'),('default_rate','Default rate %'),('loans','Completed loans'),('mean_rate','Quoted rate %')]
    dimensions = [('Tenure',by_tenure),('Score band',by_band),('Verdict cell',list(groups.values()))]
    for dimension, data in dimensions:
        for metric, title in metrics:
            chart_specs.append({'title':f'{dimension}: {title}','metric':metric,'labels':[x['label'] for x in data],'values':[x[metric] for x in data]})
    # 35 cell-specific, non-duplicated comparisons complete the evidence wall.
    for g in sorted(groups.values(), key=lambda x:(x['tenure'],x['band'])):
        chart_specs.append({'title':f"{g['label']}: income bridge",'metric':'bridge','labels':['Interest','Fee','Default loss','Net'],'values':[g['interest'],-g['fees'],-g['default_loss'],g['net']]})
        chart_specs.append({'title':f"{g['label']}: evidence and default risk",'metric':'risk','labels':['Completed loans','Default rate × 10'],'values':[g['loans'],g['default_rate']*10]})
    best = sorted([x for x in groups.values() if x['verdict']=='Fund'], key=lambda x:x['net_per_1000'], reverse=True)
    avoid = sorted([x for x in groups.values() if x['verdict']=='Avoid'], key=lambda x:x['net_per_1000'])
    return {'generated':datetime.now().strftime('%Y-%m-%d'), 'portfolio':portfolio, 'completed_loans':len(completed), 'all_loans':len(rows), 'fee_rule':'Schedule rate × principal returned; 2m/3m 1%, 4m 2.3% before 1 Apr 2026 then 3%, 5m 2.5%, 6m 3%, 12m 6%.', 'groups':list(groups.values()), 'best':best[:8], 'avoid':avoid[:8], 'by_tenure':by_tenure, 'by_band':by_band, 'charts':chart_specs}

if __name__ == '__main__':
    if len(sys.argv) != 3: raise SystemExit('Usage: python scripts/analyze.py REPORT.xlsx public/data.json')
    output = Path(sys.argv[2]); output.parent.mkdir(parents=True, exist_ok=True); output.write_text(json.dumps(build(read(sys.argv[1])), ensure_ascii=False), encoding='utf-8')
