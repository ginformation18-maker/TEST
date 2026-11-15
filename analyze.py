#!/usr/bin/env python3
"""Vectorized inventory cleaning, aging, anomaly checks, and Excel export.

Dependencies: Python 3.10+, pandas, numpy, xlsxwriter
Install: python -m pip install pandas numpy xlsxwriter
Run: python inventory_analysis.py agestats_detail_20261114.csv

All row-level operations are vectorized. Loops only iterate over a fixed
number of columns, aging bands, worksheets, or file encodings; no apply(axis=1),
iterrows(), itertuples(), or Python loops over inventory records are used.

Amounts retain their source currency/unit, which is not specified in the CSV.
"""

import argparse
import csv
import re
from pathlib import Path

import numpy as np
import pandas as pd


EXCLUSION_TERMS = ['مؤجر', 'الملحمة', 'الملابس', 'الأجبان', 'المحمصة', 'العبايات']
BANDS = [
    ('Active', 'نشط 0-90', 0, 6),
    ('Slow-moving', 'بطيء 91-180', 6, 12),
    ('Frozen', 'مجمد أكثر من 180', 12, 13),
]
REQUIRED = ['رقم الصنف', 'اسم الصنف', 'الفئة', 'رقم المورد',
            'اسم المورد', 'الرصيد', 'القيمة', 'آخر توريد', 'آخر بيع']


def normalize_arabic(series):
    """Normalize for matching only; retain original category labels in output."""
    return (series.astype('string').fillna('')
            .str.replace('[إأآٱ]', 'ا', regex=True)
            .str.replace('ـ', '', regex=False)
            .str.replace(r'[\u064b-\u065f\u0670]', '', regex=True)
            .str.replace(r'\s+', ' ', regex=True).str.strip())


def load_csv(path, encoding=None, separator=None):
    """Try common encodings strictly and require the expected source schema."""
    errors = []
    encodings = [encoding] if encoding else ['utf-8-sig', 'utf-16', 'cp1256']
    for enc in encodings:
        try:
            with path.open('r', encoding=enc, errors='strict', newline='') as f:
                sample = f.read(65536)
            if separator is not None:
                sep = separator
            else:
                try:
                    sep = csv.Sniffer().sniff(sample, delimiters=',;\t|').delimiter
                except csv.Error:
                    sep = ','
            frame = pd.read_csv(path, encoding=enc, sep=sep, dtype='string',
                                keep_default_na=False, on_bad_lines='error')
            frame.columns = frame.columns.str.replace('\ufeff', '', regex=False).str.strip()
            if frame.columns.duplicated().any():
                raise ValueError('Duplicate column names after trimming.')
            missing = set(REQUIRED) - set(frame.columns)
            if missing:
                raise ValueError(f'Missing required columns: {sorted(missing)}')
            return frame, enc, sep
        except (UnicodeError, LookupError, ValueError, pd.errors.ParserError) as exc:
            errors.append(f'{enc}: {exc}')
    raise ValueError('Unable to load CSV with the expected schema:\n' + '\n'.join(errors))


def get_age_columns(columns):
    """Validate exact intervals; avoid accidentally including total القيمة."""
    quantity, value = {}, {}
    for col in columns:
        text = str(col).replace('–', '-').replace('—', '-')
        match = re.fullmatch(r'\s*(\d+)\s*-\s*(\d+)\s+يوم\s+(كمية|قيمة)\s*', text)
        if match:
            start, end, kind = match.groups()
            key = (int(start), int(end))
        else:
            match = re.fullmatch(r'\s*181\s+يوم\s+فأكثر\s+(كمية|قيمة)\s*', text)
            if not match:
                continue
            kind = match.group(1)
            key = (181, None)
        target = quantity if kind == 'كمية' else value
        if key in target:
            raise ValueError(f'Duplicate aging interval: {key}, {kind}')
        target[key] = col
    expected = [(0, 15)] + [(start, start + 14) for start in range(16, 181, 15)] + [(181, None)]
    if set(quantity) != set(expected) or set(value) != set(expected):
        raise ValueError('Expected 13 quantity and 13 value bands: 0-15 through 166-180, then 181+.')
    return [quantity[k] for k in expected], [value[k] for k in expected]


def write_excel(path, sheets):
    with pd.ExcelWriter(path, engine='xlsxwriter', datetime_format='yyyy-mm-dd',
                        engine_kwargs={'options': {'strings_to_formulas': False,
                                                  'strings_to_urls': False}}) as writer:
        workbook = writer.book
        header = workbook.add_format({'bold': True, 'bg_color': '#183B56',
                                      'font_color': 'white', 'text_wrap': True})
        number = workbook.add_format({'num_format': '#,##0.00;[Red](#,##0.00)'})
        percent = workbook.add_format({'num_format': '0.00%'})
        red = workbook.add_format({'bg_color': '#FCE4D6', 'font_color': '#9C0006'})
        for name, frame in sheets.items():
            if len(frame) > 1048575:
                raise ValueError(f'{name} exceeds Excel row capacity.')
            frame.to_excel(writer, sheet_name=name, index=False)
            sheet = writer.sheets[name]
            sheet.right_to_left()
            sheet.freeze_panes(1, 1)
            sheet.autofilter(0, 0, len(frame), len(frame.columns) - 1)
            sheet.set_row(0, 42)
            sheet.hide_gridlines(2)
            for j, col in enumerate(frame.columns):
                sheet.write(0, j, col, header)
                fmt = None
                if pd.api.types.is_numeric_dtype(frame[col]) and not pd.api.types.is_bool_dtype(frame[col]):
                    fmt = percent if ('share' in col or 'النسبة' in col) else number
                width = 24
                if col in ['اسم الصنف', 'اسم المورد', 'الفئة']:
                    width = 40
                if col in ['سبب محتمل يحتاج تحقق', 'التوضيح']:
                    width = 95
                sheet.set_column(j, j, width, fmt)
                if len(frame) and col in ['الرصيد', 'القيمة']:
                    sheet.conditional_format(1, j, len(frame), j,
                                             {'type': 'cell', 'criteria': '<', 'value': 0, 'format': red})


def analyze(input_path, output_path, encoding=None, separator=None, as_of=None):
    raw, used_encoding, used_separator = load_csv(input_path, encoding, separator)
    data = raw.copy()
    # Column-wise operations only. Empty strings become missing, not numeric zero.
    for col in data.columns:
        data[col] = data[col].str.strip().replace('', pd.NA)
    for col in ['رقم الصنف', 'رقم المورد']:
        data[col] = data[col].str.replace(r'^=\"(.*)\"$', r'\1', regex=True)

    qty_cols, val_cols = get_age_columns(data.columns)
    numeric_cols = ['الرصيد', 'القيمة'] + qty_cols + val_cols
    for col in numeric_cols:
        converted = pd.to_numeric(data[col], errors='coerce')
        bad = converted.isna() | ~np.isfinite(converted.astype(float))
        if bad.any():
            raise ValueError(f'{col}: {int(bad.sum())} missing/invalid/nonfinite numeric values; '
                             'analysis stopped rather than silently replacing them with zero.')
        data[col] = converted.astype(float)

    date_issues = {}
    for col in ['آخر توريد', 'آخر بيع']:
        parsed = pd.to_datetime(data[col], format='%Y-%m-%d', errors='coerce')
        invalid = data[col].notna() & parsed.isna()
        date_issues[col] = int(invalid.sum())
        data[col + ' نص المصدر'] = data[col]
        data[col + ' غير صالح'] = invalid
        data[col] = parsed

    if as_of:
        snapshot = pd.Timestamp(as_of).normalize()
        snapshot_source = 'Explicit --as-of argument'
    else:
        match = re.search(r'(?<!\d)(20\d{6})(?!\d)', input_path.stem)
        snapshot = pd.to_datetime(match.group(1), format='%Y%m%d') if match else pd.NaT
        snapshot_source = 'Inferred from filename' if match else 'Not inferred; no reference date supplied'
    data['أيام منذ آخر بيع'] = (snapshot - data['آخر بيع']).dt.days
    data['تاريخ حركة بعد اللقطة'] = ((data['آخر بيع'] > snapshot) | (data['آخر توريد'] > snapshot))

    category = normalize_arabic(data['الفئة'])
    data['سبب الاستبعاد'] = ''
    for term in EXCLUSION_TERMS:
        normalized_term = normalize_arabic(pd.Series([term])).iloc[0]
        hit = category.str.contains(normalized_term, regex=False, na=False)
        prefix = np.where(data['سبب الاستبعاد'].ne(''), '، ', '')
        data['سبب الاستبعاد'] += np.where(hit, prefix + term, '')
    data['مستبعد'] = data['سبب الاستبعاد'].ne('')
    data['نطاق الملكية'] = np.where(data['مستبعد'], 'مستبعد حسب القاعدة', 'ضمن صافي الشركة')

    for _, label, start, stop in BANDS:
        data[label + ' كمية'] = data[qty_cols[start:stop]].sum(axis=1)
        data[label + ' قيمة'] = data[val_cols[start:stop]].sum(axis=1)

    data['فرق مطابقة الكمية'] = data[qty_cols].sum(axis=1) - data['الرصيد']
    data['فرق مطابقة القيمة'] = data[val_cols].sum(axis=1) - data['القيمة']
    data['شذوذ سالب'] = data['الرصيد'].lt(0) | data['القيمة'].lt(0)
    data['حالة الرصيد'] = np.select([data['الرصيد'].gt(0), data['الرصيد'].lt(0)],
                                            ['متوفر', 'سالب'], default='صفري')
    data['سبب محتمل يحتاج تحقق'] = np.select(
        [data['الرصيد'].lt(0) & data['القيمة'].lt(0), data['الرصيد'].lt(0), data['القيمة'].lt(0)],
        ['صرف قبل تسجيل التوريد أو تسويات/مرتجعات غير مطابقة؛ يلزم فحص الحركات',
         'كمية سالبة مع قيمة غير سالبة؛ يلزم فحص الكمية والتكلفة',
         'قيمة سالبة دون كمية سالبة؛ يلزم فحص تسويات التكلفة والمرتجعات'], default='')
    data['كمية موجبة وقيمة صفرية'] = data['الرصيد'].gt(0) & data['القيمة'].eq(0)
    data['كمية صفرية وقيمة غير صفرية'] = data['الرصيد'].eq(0) & data['القيمة'].abs().gt(.01)
    data['شذوذ داخل الشرائح'] = data[qty_cols].lt(0).any(axis=1) | data[val_cols].lt(0).any(axis=1)
    data['تصنيف عمر الصنف'] = ''
    for _, label, _, _ in BANDS:
        hit = data[label + ' كمية'].gt(1e-8) & data['الرصيد'].gt(0)
        prefix = np.where(data['تصنيف عمر الصنف'].ne(''), ' + ', '')
        data['تصنيف عمر الصنف'] += np.where(hit, prefix + label, '')
    data.loc[data['الرصيد'].eq(0), 'تصنيف عمر الصنف'] = 'رصيد صفري'
    data.loc[data['الرصيد'].lt(0), 'تصنيف عمر الصنف'] = 'شذوذ رصيد سالب'
    data.loc[data['تصنيف عمر الصنف'].eq(''), 'تصنيف عمر الصنف'] = 'يحتاج مراجعة'

    net = data.loc[~data['مستبعد']].copy()
    excluded = data.loc[data['مستبعد']].copy()
    anomalies = data.loc[data['شذوذ سالب']].sort_values('القيمة')
    gross_value, net_value = data['القيمة'].sum(), net['القيمة'].sum()
    rows = []
    for scope, frame in [('Gross', data), ('Excluded', excluded), ('Net', net)]:
        rows.append({'scope': scope, 'items': len(frame),
                     'available_items': int(frame['الرصيد'].gt(0).sum()),
                     'zero_items': int(frame['الرصيد'].eq(0).sum()),
                     'negative_quantity_items': int(frame['الرصيد'].lt(0).sum()),
                     'negative_value_items': int(frame['القيمة'].lt(0).sum()),
                     'negative_either_items': int(frame['شذوذ سالب'].sum()),
                     'inventory_value': frame['القيمة'].sum(),
                     'positive_value': frame.loc[frame['القيمة'].gt(0), 'القيمة'].sum(),
                     'negative_value': frame.loc[frame['القيمة'].lt(0), 'القيمة'].sum()})
    summary = pd.DataFrame(rows)

    rows = []
    for band, label, _, _ in BANDS:
        values = net[label + ' قيمة']
        rows.append({'band': band, 'age_range': label, 'net_value': values.sum(),
                     'share_of_net': values.sum() / net_value if net_value != 0 else np.nan,
                     'positive_value': values.clip(lower=0).sum(),
                     'negative_value': values.clip(upper=0).sum(),
                     'items_with_positive_quantity': int(net[label + ' كمية'].gt(0).sum())})
    aging = pd.DataFrame(rows)

    frozen_col, frozen_qty = 'مجمد أكثر من 180 قيمة', 'مجمد أكثر من 180 كمية'
    frozen = net.loc[net['الرصيد'].gt(0) & net[frozen_qty].gt(0) & net[frozen_col].gt(0)].copy()
    frozen = frozen.sort_values([frozen_col, 'رقم الصنف'], ascending=[False, True], kind='stable')
    frozen_total = frozen[frozen_col].sum()
    frozen['share_of_positive_frozen'] = frozen[frozen_col] / frozen_total if frozen_total else np.nan
    top100 = frozen.head(100).copy()
    top100.insert(0, 'rank', np.arange(1, len(top100) + 1))
    top100['cumulative_share_of_positive_frozen'] = top100['share_of_positive_frozen'].cumsum()
    top_columns = ['rank', 'رقم الصنف', 'اسم الصنف', 'الفئة', 'الرصيد', 'القيمة', frozen_qty,
                   frozen_col, 'share_of_positive_frozen', 'cumulative_share_of_positive_frozen',
                   'آخر بيع', 'آخر توريد']
    top100 = top100[top_columns]

    # No currency aggregation across quantities: units may differ by product.
    category_summary = net.groupby('الفئة', dropna=False).agg(
        items=('رقم الصنف', 'size'), inventory_value=('القيمة', 'sum'),
        active_value=('نشط 0-90 قيمة', 'sum'), slow_value=('بطيء 91-180 قيمة', 'sum'),
        frozen_value=(frozen_col, 'sum')).reset_index()
    exclusion_summary = excluded.groupby(['الفئة', 'سبب الاستبعاد'], dropna=False).agg(
        items=('رقم الصنف', 'size'), excluded_value=('القيمة', 'sum')).reset_index()

    duplicate_id = data['رقم الصنف'].notna() & data['رقم الصنف'].duplicated(keep=False)
    bad_qty_reconciliation = data['فرق مطابقة الكمية'].abs().gt(1e-6)
    bad_value_reconciliation = data['فرق مطابقة القيمة'].abs().gt(.01)
    quality = pd.DataFrame([
        ('Source rows', len(raw)), ('Source columns', len(raw.columns)),
        ('Exact duplicate source rows', int(raw.duplicated().sum())),
        ('Rows with duplicated nonmissing item IDs', int(duplicate_id.sum())),
        ('Missing item IDs', int(data['رقم الصنف'].isna().sum())),
        ('Missing categories', int(data['الفئة'].isna().sum())),
        ('Missing supplier names', int(data['اسم المورد'].isna().sum())),
        ('Missing/invalid last receipt dates', int(data['آخر توريد'].isna().sum())),
        ('Missing/invalid last sale dates', int(data['آخر بيع'].isna().sum())),
        ('Invalid nonempty receipt dates', date_issues['آخر توريد']),
        ('Invalid nonempty sale dates', date_issues['آخر بيع']),
        ('Quantity reconciliation failures > 0.000001', int(bad_qty_reconciliation.sum())),
        ('Value reconciliation failures > 0.01', int(bad_value_reconciliation.sum())),
        ('Positive quantity with zero value', int(data['كمية موجبة وقيمة صفرية'].sum())),
        ('Zero quantity with nonzero value', int(data['كمية صفرية وقيمة غير صفرية'].sum())),
        ('Movement dates after reference date', int(data['تاريخ حركة بعد اللقطة'].sum())),
    ], columns=['check', 'result'])
    review_mask = (data['شذوذ سالب'] | data['كمية موجبة وقيمة صفرية'] |
                   data['كمية صفرية وقيمة غير صفرية'] | bad_qty_reconciliation |
                   bad_value_reconciliation | duplicate_id | data['رقم الصنف'].isna() |
                   data['الفئة'].isna() | data['آخر بيع غير صالح'] |
                   data['آخر توريد غير صالح'] | data['تاريخ حركة بعد اللقطة'] |
                   data['شذوذ داخل الشرائح'])
    dictionary = pd.DataFrame({'column': raw.columns,
                               'source_dtype': raw.dtypes.astype(str).to_numpy(),
                               'clean_dtype': data[raw.columns].dtypes.astype(str).to_numpy(),
                               'missing_after_cleaning': data[raw.columns].isna().sum().to_numpy()})
    notes = pd.DataFrame([
        ('Source', str(input_path)), ('Encoding', used_encoding), ('Separator', repr(used_separator)),
        ('Reference date', str(snapshot.date()) if pd.notna(snapshot) else 'Not available'),
        ('Reference date basis', snapshot_source),
        ('Valuation', 'Gross = all source values. Net = Gross minus excluded departments. Signed balances retained.'),
        ('Currency', 'Unspecified in source; no assumed currency or conversion.'),
        ('Exclusion', 'Category substring matching after Arabic normalization: ' + ', '.join(EXCLUSION_TERMS)),
        ('Aging boundaries', 'Active 0-90 inclusive; Slow-moving 91-180 inclusive; Frozen >180 (181+) days. Day 180 is NOT frozen.'),
        ('Aging method', 'Sum source aging buckets; do not infer stock age from last sale or last receipt.'),
        ('Interpretation', 'Source stock age is not proof of no sales, expiry, FIFO methodology, or saleability.'),
        ('Classification', 'An item can span multiple bands; band item counts overlap. No whole-item forced assignment.'),
        ('Top 100', 'Net company items only; rank positive frozen portion, not total item value; require positive total and frozen quantity.'),
        ('Negative values', 'Retained for accounting reconciliation; positive frozen value reported separately for disposal review.'),
        ('Cleaning', 'Trim strings, preserve IDs as text and leading zeros, unwrap Excel ID formulas; keep missing text/dates missing.'),
        ('Duplicates', 'Flagged, not automatically deleted; duplicate item IDs may require business review.'),
        ('Numeric failures', 'Missing, invalid, or nonfinite numeric input stops processing rather than inventing zeros.'),
        ('Potential causes', 'Hypotheses only: unposted receipts, timing differences, returns or cost/quantity adjustments; verify transaction ledger.'),
        ('Precision', 'Compute with unrounded values; display two decimals. Quantity tolerance 1e-6; value tolerance 0.01.'),
    ], columns=['البند', 'التوضيح'])

    if not np.isclose(gross_value - excluded['القيمة'].sum(), net_value, rtol=0, atol=.01):
        raise AssertionError('Gross/excluded/net reconciliation failed.')
    if len(net) + len(excluded) != len(data):
        raise AssertionError('Ownership partition reconciliation failed.')
    # Source band discrepancies remain visible; never silently rebalance source data.
    sheets = {'Summary': summary, 'Aging': aging, 'Top 100 Frozen': top100,
              'Net Inventory': net, 'Excluded Departments': excluded,
              'Exclusion Summary': exclusion_summary, 'Negative Balances': anomalies,
              'Positive Frozen': frozen, 'Category Summary': category_summary,
              'Anomaly Review': data.loc[review_mask], 'Data Quality': quality,
              'Source Columns': dictionary, 'Methodology': notes}
    output_path.parent.mkdir(parents=True, exist_ok=True)
    write_excel(output_path, sheets)
    print(f'Loaded {len(raw):,} rows x {len(raw.columns)} columns; encoding={used_encoding}')
    print('\nSource columns:\n' + '\n'.join(raw.columns))
    print('\nInventory summary:\n' + summary.to_string(index=False))
    print('\nAging:\n' + aging.to_string(index=False))
    print('\nData quality:\n' + quality.to_string(index=False))
    print(f'\nPositive frozen value: {frozen_total:,.2f}')
    print(f'Top {len(top100)} frozen portion value: {top100[frozen_col].sum():,.2f}')
    print(f'Workbook saved: {output_path.resolve()}')
    return sheets


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('input', type=Path, nargs='?', default=Path('inventory.csv'), help='Inventory CSV path (defaults to inventory.csv)')
    parser.add_argument('-o', '--output', type=Path, default=Path('reports/inventory_analysis_top100.xlsx'))
    parser.add_argument('--encoding', help='Override automatic encoding detection, e.g. utf-8-sig')
    parser.add_argument('--separator', help='Override separator detection, e.g. comma or semicolon')
    parser.add_argument('--as-of', help='Optional YYYY-MM-DD reference date; affects sale recency, not source aging')
    args = parser.parse_args()
    if args.input.resolve() == args.output.resolve():
        parser.error('Input and output paths must differ.')
    analyze(args.input, args.output, args.encoding, args.separator, args.as_of)


if __name__ == '__main__':
    main()