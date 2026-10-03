import win32com.client as wc
import subprocess

DB = r"D:\Users\Dkeiz_Dao\Desktop\access\QL_SINH_VIEN_21.accdb"
OUT = r"D:\Users\Dkeiz_Dao\Desktop\access\_db_dump.txt"

try:
    subprocess.run(['taskkill', '/F', '/IM', 'ACCESS.EXE'], capture_output=True, timeout=15)
except Exception:
    pass

con = wc.Dispatch('ADODB.Connection')
rs = wc.Dispatch('ADODB.Recordset')
con.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + DB + ';')

def dump(table):
    rs.Open(table, con)
    n = rs.Fields.Count
    cols = [rs.Fields(i).Name for i in range(n)]
    types = [rs.Fields(i).Type for i in range(n)]
    rows = []
    while not rs.EOF:
        row = []
        for i in range(n):
            try:
                v = rs.Fields(i).Value
                row.append(v)
            except Exception:
                row.append('<err>')
        rows.append(row)
        rs.MoveNext()
    rs.Close()
    return cols, types, rows

typemap = {2:'tinyint',3:'smallint',4:'long',5:'single',6:'double',7:'currency',8:'datetime',9:'text',10:'yesno',11:'boolean',12:'bstr',13:'date',15:'char',16:'byte',17:'decimal',20:'guid',64:'longtext',202:'varwchar/text',203:'varbinary',128:'sybase',200:'decimal'}

lines = []
for t in ['SINHVIEN','MONHOC','KETQUA']:
    try:
        cols, types, rows = dump(t)
        lines.append('=== TABLE %s === rows=%d' % (t, len(rows)))
        lines.append('COLUMNS (name:type):')
        for c, ty in zip(cols, types):
            lines.append('    %s = %s' % (c, typemap.get(ty, str(ty))))
        lines.append('DATA:')
        for r in rows:
            lines.append('    ' + repr(r))
    except Exception as e:
        lines.append('=== TABLE %s === ERROR: %r' % (t, e))
lines.append('DONE')

con.Close()

with open(OUT, 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))
print('WROTE', OUT, 'lines=', len(lines))
