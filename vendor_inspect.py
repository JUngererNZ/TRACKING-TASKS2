import sys
from pathlib import Path
try:
    import pandas as pd
except Exception as e:
    print('PANDAS IMPORT ERROR:', e)
    sys.exit(2)

p = Path('vendor-report')
files = sorted([f for f in p.glob('*.xlsx')])
candidates = [f for f in files if 'ORIENTO' in f.name.upper() or 'CATERPILLAR' in f.name.upper()]
if not candidates:
    print('No ORIENTO/CATERPILLAR candidates found in vendor-report')
    sys.exit(0)

for f in candidates:
    print('\n=== FILE:', f.name, '===')
    try:
        xl = pd.ExcelFile(f)
        print('Sheets:', xl.sheet_names)
        for s in xl.sheet_names:
            print('\n-- Sheet:', repr(s))
            try:
                df = pd.read_excel(f, sheet_name=s, header=None, nrows=60, dtype=str)
            except Exception as e:
                print('  Could not read sheet', s, '->', e)
                continue
            # print first 12 rows and up to 40 columns
            for i, row in df.fillna('').astype(str).iloc[:12, :40].iterrows():
                joined = '|'.join([c.strip() for c in row.tolist()])
                print('   ', i, '|', joined[:100])
    except Exception as e:
        print('  ERROR reading file:', e)

print('\nDone')
