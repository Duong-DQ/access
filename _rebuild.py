# -*- coding: utf-8 -*-
import datetime
import win32com.client as wc

DB  = r"D:\Users\Dkeiz_Dao\Desktop\access\QL_SINH_VIEN_21.accdb"
LOG = r"D:\Users\Dkeiz_Dao\Desktop\access\_rebuild_log.txt"
lines = []
def log(s): lines.append(str(s))

con = wc.Dispatch('ADODB.Connection')
rs  = wc.Dispatch('ADODB.Recordset')
con.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + DB + ';')

def read_table(name):
    rs.Open(name, con, 3)
    n = rs.Fields.Count
    cols = [rs.Fields(i).Name for i in range(n)]
    rows = []
    while not rs.EOF:
        row = {cols[i]: rs.Fields(i).Value for i in range(n)}
        rows.append(row)
        rs.MoveNext()
    rs.Close()
    return cols, rows

def q(s):
    return "'" + str(s).replace("'", "''") + "'"

def date_str(v):
    if v is None: return None
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime('%Y-%m-%d')
    return str(v)

def truthy(v):
    return v not in (0, False, None, '0')

# 1. READ original data
sv_cols, sv_rows = read_table('SINHVIEN')
mh_cols, mh_rows = read_table('MONHOC')
kq_cols, kq_rows = read_table('KETQUA')
log('READ  SINHVIEN=%d  MONHOC=%d  KETQUA=%d' % (len(sv_rows), len(mh_rows), len(kq_rows)))
log('SINHVIEN cols: ' + ','.join(sv_cols))
log('MONHOC   cols: ' + ','.join(mh_cols))
log('KETQUA   cols: ' + ','.join(kq_cols))

# 2. DROP old
for t in ['KETQUA', 'SINHVIEN', 'MONHOC', 'KHOA']:
    try:
        con.Execute('DROP TABLE ' + t); log('DROPPED ' + t)
    except Exception as e:
        log('DROP ' + t + ' (skip): ' + str(e))

# 3. CREATE page-2 schema
con.Execute("CREATE TABLE KHOA (MAKH TEXT(10) NOT NULL PRIMARY KEY, TENKH TEXT(50))")
con.Execute("CREATE TABLE SINHVIEN (MASV TEXT(10) NOT NULL PRIMARY KEY, HOSV TEXT(30), TENSV TEXT(30), PHAI YESNO, NGAYSINH DATETIME, NOISINH TEXT(50), MAKH TEXT(10), HOCBONG YESNO)")
con.Execute("CREATE TABLE MONHOC (MAMH TEXT(10) NOT NULL PRIMARY KEY, TENMH TEXT(50), SOTIET COUNTER)")
con.Execute("CREATE TABLE KETQUA (MASV TEXT(10), MAMH TEXT(10), DIEM SINGLE)")
log('CREATED KHOA, SINHVIEN, MONHOC, KETQUA')

# 4. KHOA from distinct MALOP
malops = []
for r in sv_rows:
    v = r.get('MALOP')
    if v is not None and v not in malops:
        malops.append(v)
for m in malops:
    con.Execute('INSERT INTO KHOA (MAKH, TENKH) VALUES (%s, %s)' % (q(m), q('Khoa ' + m)))
log('KHOA rows=%d  codes=%s' % (len(malops), ','.join(malops)))

# 5. SINHVIEN  (NOISINH=DIACHI, MAKH=MALOP, HOCBONG=NULL)
for r in sv_rows:
    masv=r.get('MASV'); hosv=r.get('HOSV'); tensv=r.get('TENSV')
    phai=r.get('PHAI'); ngnh=r.get('NGAYSINH')
    diachi=r.get('DIACHI'); malop=r.get('MALOP')
    phai_sql = '-1' if truthy(phai) else '0'
    ngnh_sql = ('#' + date_str(ngnh) + '#') if ngnh is not None else 'NULL'
    nois_sql = q(diachi) if diachi is not None else 'NULL'
    makh_sql = q(malop) if malop is not None else 'NULL'
    con.Execute('INSERT INTO SINHVIEN (MASV,HOSV,TENSV,PHAI,NGAYSINH,NOISINH,MAKH,HOCBONG) VALUES ('
        + q(masv)+', '+(q(hosv) if hosv is not None else 'NULL')+', '+(q(tensv) if tensv is not None else 'NULL')
        +', '+phai_sql+', '+ngnh_sql+', '+nois_sql+', '+makh_sql+', NULL)')
log('SINHVIEN inserted=%d' % len(sv_rows))

# 6. MONHOC  (SOTIET=SOTC)
for r in mh_rows:
    mamh=r.get('MAMH'); tenmh=r.get('TENMH'); sotec=r.get('SOTC')
    sotec_sql = str(int(sotec)) if sotec is not None else 'NULL'
    con.Execute('INSERT INTO MONHOC (MAMH,TENMH,SOTIET) VALUES ('
        + q(mamh)+', '+(q(tenmh) if tenmh is not None else 'NULL')+', '+sotec_sql+')')
log('MONHOC inserted=%d' % len(mh_rows))

# 7. KETQUA  (DIEM = LAN2>0 ? LAN2 : LAN1)
for r in kq_rows:
    masv=r.get('MASV'); mamh=r.get('MAMH')
    try: l1f = float(r.get('DIEMLAN1')) if r.get('DIEMLAN1') is not None else 0.0
    except: l1f = 0.0
    try: l2f = float(r.get('DIEMLAN2')) if r.get('DIEMLAN2') is not None else 0.0
    except: l2f = 0.0
    diem = l2f if l2f > 0 else l1f
    con.Execute('INSERT INTO KETQUA (MASV,MAMH,DIEM) VALUES ('+q(masv)+', '+q(mamh)+', '+str(diem)+')')
log('KETQUA inserted=%d' % len(kq_rows))

# 8. VERIFY
def count(name):
    rs.Open('SELECT COUNT(*) AS C FROM ' + name, con, 3)
    c = rs.Fields(0).Value; rs.Close(); return c
log('VERIFY counts:  KHOA=%d  SINHVIEN=%d  MONHOC=%d  KETQUA=%d'
    % (count('KHOA'), count('SINHVIEN'), count('MONHOC'), count('KETQUA')))

rs.Open('SELECT TOP 3 MASV,HOSV,TENSV,NOISINH,MAKH FROM SINHVIEN', con, 3)
while not rs.EOF:
    log('  SV: %s | %s %s | noisinh=%s | makH=%s' % (rs.Fields(0).Value, rs.Fields(1).Value, rs.Fields(2).Value, rs.Fields(3).Value, rs.Fields(4).Value))
    rs.MoveNext()
rs.Close()
rs.Open('SELECT TOP 3 MAMH,TENMH,SOTIET FROM MONHOC', con, 3)
while not rs.EOF:
    log('  MH: %s | %s | %s' % (rs.Fields(0).Value, rs.Fields(1).Value, rs.Fields(2).Value))
    rs.MoveNext()
rs.Close()
rs.Open('SELECT TOP 3 MASV,MAMH,DIEM FROM KETQUA', con, 3)
while not rs.EOF:
    log('  KQ: %s | %s | %s' % (rs.Fields(0).Value, rs.Fields(1).Value, rs.Fields(2).Value))
    rs.MoveNext()
rs.Close()

con.Close()
with open(LOG, 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))
print('REBUILD DONE')
