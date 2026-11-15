#!/usr/bin/env python3
"""Vectorized ERP inventory analysis. Column names can be mapped with environment variables."""
from pathlib import Path
import json, os, sys
import pandas as pd

INPUT_DIR = Path(os.getenv('INPUT_DIR', 'data'))
OUT = Path(os.getenv('OUTPUT_DIR', 'reports'))
OUT.mkdir(parents=True, exist_ok=True)
ALIASES = {
 'department':['department','dept','القسم','الادارة','الإدارة','اسم القسم'],
 'quantity':['quantity','qty','balance','stock','الكمية','الرصيد','الرصيد الكمي'],
 'cost':['unit_cost','cost','cost_price','التكلفة','تكلفة الوحدة','سعر التكلفة'],
 'age':['age_days','aging_days','days_old','days','عمر المخزون','أيام','ايام'],
 'value':['inventory_value','stock_value','total_cost','قيمة المخزون','القيمة'],
}
def norm(s): return str(s).strip().lower().replace(' ', '_')
def load_file(p):
    if p.suffix.lower() in ('.xlsx','.xls'):
        return pd.read_excel(p)
    for enc in ('utf-8-sig','utf-8','cp1256'):
        try: return pd.read_csv(p, encoding=enc, sep=None, engine='python')
        except UnicodeDecodeError: continue
    raise ValueError(f'Unsupported encoding: {p}')
def main():
    files = sorted([*INPUT_DIR.glob('*.csv'),*INPUT_DIR.glob('*.xlsx'),*INPUT_DIR.glob('*.xls')])
    if not files: raise SystemExit(f'No CSV/Excel inventory files found in {INPUT_DIR}')
    df = pd.concat([load_file(f).assign(source_file=f.name) for f in files], ignore_index=True)
    original = list(df.columns); by_norm = {norm(c): c for c in original}
    cols = {}
    for key, aliases in ALIASES.items():
        env = os.getenv('COL_'+key.upper())
        hit = env if env in df.columns else next((by_norm[norm(a)] for a in ([env] if env else [])+aliases if norm(a) in by_norm), None)
        if hit is not None: cols[key] = hit
    missing = [k for k in ('department','quantity') if k not in cols]
    if missing: raise SystemExit('Could not identify required columns: '+', '.join(missing)+'; set COL_DEPARTMENT and COL_QUANTITY')
    for key in ('quantity','cost','age','value'):
        if key in cols: df[cols[key]] = pd.to_numeric(df[cols[key]].astype(str).str.replace(',','',regex=False), errors='coerce')
    dept = df[cols['department']].fillna('').astype(str)
    rented_names = ['الملحمة','الملابس','الاجبان','المحمصة','العبايات']
    rented = dept.str.contains('مؤجر', case=False, regex=False) | dept.str.strip().isin(rented_names)
    df['is_rented_department'] = rented
    active = df.loc[~rented].copy()
    q = active[cols['quantity']]
    active['negative_quantity_flag'] = q.lt(0).fillna(False)
    if 'cost' in cols: active['negative_cost_flag'] = active[cols['cost']].lt(0).fillna(False)
    else: active['negative_cost_flag'] = False
    if 'value' in cols: active['frozen_value'] = active[cols['value']]
    elif 'cost' in cols: active['frozen_value'] = q * active[cols['cost']]
    else: active['frozen_value'] = pd.NA
    if 'age' in cols:
        days = active[cols['age']]
        active['aging_bucket'] = pd.cut(days, [-float('inf'),90,180,float('inf')], labels=['0-90','91-180','180+'], right=True).astype('string')
        # day zero belongs to first bucket; exclude negative/unknown ages
        active.loc[days.eq(0), 'aging_bucket'] = '0-90'
        active.loc[days.lt(0) | days.isna(), 'aging_bucket'] = 'Unknown/invalid'
    else: active['aging_bucket'] = 'Age unavailable'
    summary = {
      'input_files': [p.name for p in files], 'input_rows': int(len(df)),
      'excluded_rented_rows': int(rented.sum()), 'analyzed_rows': int(len(active)),
      'negative_quantity_rows': int(active['negative_quantity_flag'].sum()),
      'negative_cost_rows': int(active['negative_cost_flag'].sum()),
      'frozen_value_total': None if active['frozen_value'].isna().all() else float(active['frozen_value'].sum()),
      'frozen_value_by_aging_bucket': {str(k): None if pd.isna(v) else float(v) for k,v in active.groupby('aging_bucket', dropna=False, observed=False)['frozen_value'].sum(min_count=1).items()},
      'columns_detected': cols,
    }
    (OUT/'summary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
    with pd.ExcelWriter(OUT/'inventory_analysis.xlsx', engine='openpyxl') as writer:
        pd.DataFrame([summary]).to_excel(writer, sheet_name='Executive Summary', index=False)
        active.to_excel(writer, sheet_name='Analyzed Inventory', index=False)
        df.loc[rented].to_excel(writer, sheet_name='Excluded Rented', index=False)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
if __name__ == '__main__': main()
