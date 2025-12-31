from pathlib import Path
from xltpl.writerx import BookWriter
import tempfile
import win32com.client

p = Path('examples/sample_invoice.xlsx')
w = BookWriter(str(p))
payload = [{
    'customer': {'name': 'test', 'company': 'test'},
    'issue_date': '2025-12-29',
    'note': 'test',
    'items': [],
    'remarks': 'test'
}]
w.render_book(payload)
print('Defined names after render:', list(w.workbook.defined_names.keys())[:10])

out = Path(tempfile.mkdtemp()) / 'out.xlsx'
print(f'Saving to: {out}')
w.workbook.save(out)
print('Saved with openpyxl')

excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
try:
    print('Opening with COM...')
    wb2 = excel.Workbooks.Open(
        str(out.absolute()),
        UpdateLinks=False,
        ReadOnly=False,
        AddToMru=False
    )
    print('Opened successfully')
    wb2.Close(SaveChanges=False)
except Exception as e:
    print(f'Error opening: {e}')
finally:
    excel.Quit()
