import win32com.client as wc, shutil, subprocess, os

DB  = r"D:\Users\Dkeiz_Dao\Desktop\access\QL_SINH_VIEN_21.accdb"
SCR = r"D:\Users\Dkeiz_Dao\Desktop\access\_scratch_test.accdb"
OUT = r"D:\Users\Dkeiz_Dao\Desktop\access\_mech_test.txt"

try:
    subprocess.run(['taskkill','/F','/IM','ACCESS.EXE'], capture_output=True, timeout=15)
except Exception:
    pass
try:
    if os.path.exists(SCR): os.remove(SCR)
    shutil.copy(DB, SCR)
except Exception as e:
    print('copy fail', repr(e)); raise

lines = []

# ---- Test 1: ADO CREATE TABLE + INSERT ----
try:
    con = wc.Dispatch('ADODB.Connection')
    con.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source='+SCR+';')
    con.Execute("CREATE TABLE TESTPARENT (PID TEXT(10) PRIMARY KEY, NAME TEXT(20))")
    lines.append('T1 ADO CREATE TABLE (PK): OK')
    con.Execute("CREATE TABLE TESTCHILD (CID TEXT(10) PRIMARY KEY, PID TEXT(10))")
    con.Execute("INSERT INTO TESTPARENT VALUES ('P1','alpha')")
    con.Execute("INSERT INTO TESTCHILD VALUES ('C1','P1')")
    lines.append('T1 ADO CREATE+INSERT: OK')
    con.Close()
except Exception as e:
    lines.append('T1 ADO CREATE TABLE: FAIL %r' % (e,))

# ---- Test 2: DAO.DBEngine variants from Python ----
dao_db = None
for ident in ['DAO.DBEngine','DAO.DBEngine.36','DAO.DBEngine.16','DAO.DBEngine.12.0']:
    try:
        dbeng = wc.Dispatch(ident)
        db = dbeng.OpenDatabase(SCR)
        lines.append('T2 DAO.OpenDatabase via %s: OK  TableDefs=%d' % (ident, db.TableDefs.Count))
        dao_db = db
        dao_ident = ident
        break
    except Exception as e:
        lines.append('T2 DAO %s: FAIL %r' % (ident, e))

if dao_db is not None:
    # list existing tables
    try:
        names = [dao_db.TableDefs(i).Name for i in range(dao_db.TableDefs.Count)]
        lines.append('T2 tables: %r' % names)
    except Exception as e:
        lines.append('T2 list tables FAIL %r' % e)
    # Test 3: CreateRelation with referential integrity
    try:
        rel = dao_db.CreateRelation('RelParentChild')
        rel.Type = 1  # dbRelationOneToMany
        rel.Table = 'TESTPARENT'
        rel.ForeignTable = 'TESTCHILD'
        rel.Fields('PID').Name = 'PID'
        rel.Updatable = True
        rel.Implied = False
        dao_db.Relations.Append(rel)
        lines.append('T3 DAO CreateRelation + Append: OK')
        # read back
        rr = dao_db.Relations('RelParentChild')
        lines.append('T3 readback: Table=%s ForeignTable=%s Fields=%d' % (rr.Table, rr.ForeignTable, rr.Fields.Count))
        # cleanup
        dao_db.Relations.Delete('RelParentChild')
        lines.append('T3 DeleteRelation: OK')
    except Exception as e:
        lines.append('T3 DAO CreateRelation: FAIL %r' % (e,))
    dao_db.Close()

# ---- Test 4: Access.Application + DoCmd.RunSQL + VBE access ----
try:
    app = wc.Dispatch('Access.Application')
    app.Visible = False
    app.OpenCurrentDatabase(SCR)
    lines.append('T4 Access.OpenCurrentDatabase: OK')
    # DoCmd.RunSQL
    try:
        app.DoCmd.RunSQL("CREATE TABLE TESTACMD (AID TEXT(5) PRIMARY KEY)")
        lines.append('T4 DoCmd.RunSQL CREATE TABLE: OK')
    except Exception as e:
        lines.append('T4 DoCmd.RunSQL: FAIL %r' % (e,))
    # VBE access (needs 'Trust access to VBA project object model')
    try:
        app.Visible = True
        vbe = app.VBE
        lines.append('T4 app.VBE: OK, VBProjects=%d' % vbe.VBProjects.Count)
    except Exception as e:
        lines.append('T4 app.VBE: FAIL %r' % (e,))
    app.CloseCurrentDatabase()
    lines.append('T4 Access.CloseCurrentDatabase: OK')
except Exception as e:
    lines.append('T4 Access app: FAIL %r' % (e,))

# cleanup scratch
try:
    subprocess.run(['taskkill','/F','/IM','ACCESS.EXE'], capture_output=True, timeout=15)
    if os.path.exists(SCR): os.remove(SCR)
    lines.append('CLEANUP scratch removed')
except Exception as e:
    lines.append('CLEANUP fail %r' % e)

with open(OUT,'w',encoding='utf-8') as f:
    f.write('\n'.join(lines))
print('done, wrote', OUT)
