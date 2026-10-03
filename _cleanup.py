# -*- coding: utf-8 -*-
import win32com.client as wc

DB = r"D:\Users\Dkeiz_Dao\Desktop\access\QL_SINH_VIEN_21.accdb"
OUT = r"D:\Users\Dkeiz_Dao\Desktop\access\_final_check.txt"
PREFIXES = ('mRel', 'mWiz', 'mFW', 'mAcCmd', 'mDiag')

app = wc.Dispatch('Access.Application')
app.Visible = False
app.OpenCurrentDatabase(DB, False, False)
removed = []
try:
    proj = app.VBE.VBProjects(1)
    for m in list(proj.VBComponents):
        n = m.Name
        if n.startswith(PREFIXES):
            proj.VBComponents.Remove(m)
            removed.append(n)
except Exception as e:
    removed = removed + ['ERR:' + str(e)]
finally:
    try:
        app.CloseCurrentDatabase()
    except Exception:
        pass

lines = ['REMOVED_MODULES: ' + (repr(removed) if removed else 'none')]

c = wc.Dispatch('ADODB.Connection')
c.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + DB + ';')
r = wc.Dispatch('ADODB.Recordset')
for t in ['KHOA', 'SINHVIEN', 'MONHOC', 'KETQUA']:
    try:
        r.Open(t, c)
        n = r.Fields.Count
        cnt = r.RecordCount
        r.Close()
        lines.append('%s cols=%d rows=%d' % (t, n, cnt))
    except Exception as e:
        lines.append('%s ERR %s' % (t, e))
c.Close()

open(OUT, 'w', encoding='utf-8').write('\n'.join(lines))
print('done')
