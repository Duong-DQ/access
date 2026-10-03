import win32com.client as wc
import subprocess, sys

DB = r"D:\Users\Dkeiz_Dao\Desktop\access\QL_SINH_VIEN_21.accdb"

# kill orphan access
try:
    subprocess.run(['taskkill', '/F', '/IM', 'ACCESS.EXE'], capture_output=True, timeout=15)
except Exception:
    pass

def val(v):
    try:
        import pythoncom
        if v is None:
            return None
        return v
    except Exception:
        return v

con = wc.Dispatch('ADODB.Connection')
rs = wc.Dispatch('ADODB.Recordset')
con.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + DB + ';')

def dump(table):
    rs.Open(table, con)
    n = rs.Fields.Count
    cols = []
    types = []
    for i in range(n):
        cols.append(rs.Fields(i).Name)
        types.append(rs.Fields(i).Type)
    rows = []
    while not rs.EOF:
        row = []
        for i in range(n):
            try:
                v = rs.Fields(i).Value
                row.append(v)
            except Exception as e:
                row.append('<err>')
        rows.append(row)
        rs.MoveNext()
    rs.Close()
    return cols, types, rows

# ADO type map
typemap = {2:'tinyint',3:'smallint',4:'long',5:'single',6:'double',7:'currency',8:'datetime',9:'text',10:'oyes',11:'error',12:'iid',13:'boolean',14:'variant',15:'char',16:'byte',17:'decimal',20:'guid',64:'longtext'}

for t in ['SINHVIEN','MONHOC','KETQUA']:
    try:
        cols, types, rows = dump(t)
        print('=== TABLE', t, '===  rows=', len(rows))
        print('COLUMNS (name:type):')
        for c, ty in zip(cols, types):
            print('   ', c, '=', typemap.get(ty, str(ty)))
        print('DATA:')
        for r in rows:
            print('   ', r)
    except Exception as e:
        print('=== TABLE', t, '=== ERROR:', repr(e))

con.Close()
print('DONE')
