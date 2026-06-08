import openpyxl

wb = openpyxl.load_workbook('3400021285580.2.xlsx', data_only=True)
ws = wb.active
terms = ['Ultra 7', 'NVIDIA', 'Windows 11', 'SSD', '64GB', 'FCM2250', 'PS NBD', 'Processor']
out = []
for idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
    text = ' | '.join('' if x is None else str(x) for x in row)
    if any(t.lower() in text.lower() for t in terms):
        out.append((idx, text))
print('FOUND', len(out))
for idx, text in out[:200]:
    print(idx, text)
