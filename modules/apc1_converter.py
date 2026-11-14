#!/usr/bin/env python
"""APC1 -> OIT converter v8.
Put this .py and APC1 Excel in the same folder, then run:
    python APC1_to_OIT_v8.py
The clean OIT2 row-2 header template with Feature Group in column A is embedded in this Python file.
"""
from pathlib import Path
from io import BytesIO
from copy import copy
from datetime import datetime
import base64, os, re, sys
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, PatternFill
from openpyxl.utils import get_column_letter

HERE=Path(__file__).resolve().parent
EMBEDDED_OIT_TEMPLATE="""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"""
MAPPED_APC_COLUMNS={8,9,10,11,12,13,14,15,17,22,23} # H,I,J,K,L,M,N,O,Q,V,W

def choose_apc1():
    files=sorted(p for p in HERE.iterdir() if p.suffix.lower() in {'.xlsx','.xlsm'} and not p.name.startswith('~$') and '_OIT_result_' not in p.name)
    if not files: raise FileNotFoundError('No APC1 Excel file found in the script folder.')
    print('\nWhich file is APC1?')
    for i,p in enumerate(files,1): print(f'  {i}. {p.name}')
    while True:
        s=input('Enter file number: ').strip()
        if s.isdigit() and 1<=int(s)<=len(files): return files[int(s)-1]
        print('Invalid selection. Try again.')

def find_data_sheet(wb):
    req={'Device Category / Feature Group','Required?','Selection Type','Device Model Name/ Part Number','Device Name/Feature Description','Associated Mask Layers','Is Experimental Device','RESTRICTED','Group Exclusion','Comments'}
    for ws in wb.worksheets:
        headers={str(c.value).strip() for c in ws[1] if c.value is not None}
        if req.issubset(headers): return ws
    raise ValueError('Expected APC1 headers were not found.')

def copy_style(src,dst):
    dst._style=copy(src._style); dst.font=copy(src.font); dst.fill=copy(src.fill)
    dst.border=copy(src.border); dst.alignment=copy(src.alignment)
    dst.number_format=src.number_format; dst.protection=copy(src.protection)

def normalize_yes_no(v):
    t=str(v or '').strip().casefold()
    return 'Yes' if t=='yes' else 'No' if t=='no' else str(v or '').strip()

def qualified_from_experimental(v):
    t=str(v or '').strip().casefold()
    return 'No' if t=='yes' else 'Yes' if t=='no' else ''

def category_range(required_value, selection_value):
    r=str(required_value or '').strip().casefold()
    s=str(selection_value or '').strip().casefold()
    mapping={
        ('yes','multi select'):'Include one or any (1-99)',
        ('yes','single select'):'Must include only one (1-1)',
        ('no','multi select'):'Include none or any (0-99)',
    }
    # For an unexpected combination, preserve Selection Type rather than inventing a rule.
    return mapping.get((r,s), str(selection_value or '').strip())

def n_a_if_blank(v): return 'n/a' if v in (None,'') else v

def main():
    print('=== APC1 to OIT Converter v9 ===')
    ap=choose_apc1()
    awb=load_workbook(ap,data_only=True); aws=find_data_sheet(awb)
    owb=load_workbook(BytesIO(base64.b64decode(EMBEDDED_OIT_TEMPLATE))); ows=owb[owb.sheetnames[0]]
    # Keep only OIT2 title/column-name structure; discard row 3 onward.
    if ows.max_row>=3: ows.delete_rows(3,ows.max_row-2)
    for sh in list(owb.worksheets):
        if sh is not ows: owb.remove(sh)
    ows.title='OIT_result'

    # Remaining APC1 columns go to L and after.
    remaining=[c for c in range(1,aws.max_column+1) if c not in MAPPED_APC_COLUMNS]
    for oc,sc in enumerate(remaining,13):
        h=aws.cell(1,sc).value; ows.cell(2,oc).value=h
        ows.column_dimensions[get_column_letter(oc)].width=max(14,min(35,len(str(h or ''))+2))

    outrow=3; converted=0
    for sr in range(2,aws.max_row+1):
        if all(aws.cell(sr,c).value in (None,'') for c in range(1,aws.max_column+1)): continue
        vals={
            1:aws.cell(sr,14).value, # N -> A Feature Group (FEOL/BEOL)
            2:aws.cell(sr,8).value, # H -> B Feature Category
            3:category_range(aws.cell(sr,9).value,aws.cell(sr,10).value), # I+J -> C Category Range
            4:'n/a', # D Notes
            5:aws.cell(sr,11).value, # K -> E Feature ID
            6:aws.cell(sr,12).value, # L -> F Feature Description
            7:n_a_if_blank(aws.cell(sr,13).value), # M -> G Additional Mask List
            8:normalize_yes_no(aws.cell(sr,17).value), # Q -> H Restricted
            9:qualified_from_experimental(aws.cell(sr,15).value), # O -> I Qualified
            10:'n/a', # J Required FPN
            11:n_a_if_blank(aws.cell(sr,22).value), # V -> K Disallowed FPN
            12:n_a_if_blank(aws.cell(sr,23).value), # W -> L Pcell name
        }
        for c,v in vals.items(): ows.cell(outrow,c).value=v
        for oc,sc in enumerate(remaining,13): ows.cell(outrow,oc).value=aws.cell(sr,sc).value
        outrow+=1; converted+=1

    # A:L thin borders; M+ no fill and no borders.
    thin=Side(style='thin',color='000000'); bordered=Border(left=thin,right=thin,top=thin,bottom=thin)
    last=max(2,outrow-1)
    for row in ows.iter_rows(min_row=2,max_row=last,min_col=1,max_col=12):
        for cell in row: cell.border=bordered
    clear_fill=PatternFill(fill_type=None); clear_border=Border()
    if ows.max_column>=13:
        for row in ows.iter_rows(min_row=1,max_row=last,min_col=13,max_col=ows.max_column):
            for cell in row: cell.fill=clear_fill; cell.border=clear_border

    ows.freeze_panes='A3'; ows.auto_filter.ref=f'A2:{get_column_letter(ows.max_column)}{last}'
    mmdd=datetime.now().strftime('%m%d')
    m=re.search(r'PT\d+',ap.name,re.I); pt=m.group(0).upper() if m else 'PT_UNKNOWN'
    out=HERE/f'{mmdd}_OIT_result_{pt}.xlsx'; owb.save(out)
    print(f'\nCompleted. Rows converted: {converted}\nResult: {out}')
    try:
        if sys.platform.startswith('win'): os.startfile(out)
    except Exception: pass

# Streamlit wrapper calls convert_uploaded_file() below.


def convert_uploaded_file(uploaded_bytes: bytes, uploaded_name: str, output_path: Path, progress_callback=None) -> int:
    """Run v9 conversion logic from uploaded workbook bytes and save to output_path."""
    from io import BytesIO
    from openpyxl import load_workbook
    import base64, re
    from copy import copy
    from openpyxl.styles import Border, Side, PatternFill
    from openpyxl.utils import get_column_letter

    awb = load_workbook(BytesIO(uploaded_bytes), data_only=True)
    aws = find_data_sheet(awb)
    owb = load_workbook(BytesIO(base64.b64decode(EMBEDDED_OIT_TEMPLATE)))
    ows = owb[owb.sheetnames[0]]
    if ows.max_row >= 3:
        ows.delete_rows(3, ows.max_row - 2)
    for sh in list(owb.worksheets):
        if sh is not ows:
            owb.remove(sh)
    ows.title = "OIT_result"

    remaining = [c for c in range(1, aws.max_column + 1) if c not in MAPPED_APC_COLUMNS]
    for oc, sc in enumerate(remaining, 13):
        h = aws.cell(1, sc).value
        ows.cell(2, oc).value = h
        ows.column_dimensions[get_column_letter(oc)].width = max(14, min(35, len(str(h or "")) + 2))

    outrow = 3
    converted = 0
    total_rows = max(1, aws.max_row - 1)
    for sr in range(2, aws.max_row + 1):
        if progress_callback and ((sr - 1) % max(1, total_rows // 100) == 0 or sr == aws.max_row):
            progress_callback(min(85, 15 + int((sr - 1) / total_rows * 70)), f"Converting APC1 rows: {sr - 1:,} / {total_rows:,}")
        if all(aws.cell(sr, c).value in (None, "") for c in range(1, aws.max_column + 1)):
            continue
        vals = {
            1: aws.cell(sr, 14).value,
            2: aws.cell(sr, 8).value,
            3: category_range(aws.cell(sr, 9).value, aws.cell(sr, 10).value),
            4: "n/a",
            5: aws.cell(sr, 11).value,
            6: aws.cell(sr, 12).value,
            7: n_a_if_blank(aws.cell(sr, 13).value),
            8: normalize_yes_no(aws.cell(sr, 17).value),
            9: qualified_from_experimental(aws.cell(sr, 15).value),
            10: "n/a",
            11: n_a_if_blank(aws.cell(sr, 22).value),
            12: n_a_if_blank(aws.cell(sr, 23).value),
        }
        for c, v in vals.items():
            ows.cell(outrow, c).value = v
        for oc, sc in enumerate(remaining, 13):
            ows.cell(outrow, oc).value = aws.cell(sr, sc).value
        outrow += 1
        converted += 1

    thin = Side(style="thin", color="000000")
    bordered = Border(left=thin, right=thin, top=thin, bottom=thin)
    last = max(2, outrow - 1)
    for row in ows.iter_rows(min_row=2, max_row=last, min_col=1, max_col=12):
        for cell in row:
            cell.border = bordered
    clear_fill = PatternFill(fill_type=None)
    clear_border = Border()
    if ows.max_column >= 13:
        for row in ows.iter_rows(min_row=1, max_row=last, min_col=13, max_col=ows.max_column):
            for cell in row:
                cell.fill = clear_fill
                cell.border = clear_border

    ows.freeze_panes = "A3"
    ows.auto_filter.ref = f"A2:{get_column_letter(ows.max_column)}{last}"
    if progress_callback:
        progress_callback(92, "Saving OIT workbook...")
    owb.save(output_path)
    awb.close()
    owb.close()
    return converted
