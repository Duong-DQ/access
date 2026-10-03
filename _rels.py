# -*- coding: utf-8 -*-
import time
import win32com.client as wc

DB = r"D:\Users\Dkeiz_Dao\Desktop\access\QL_SINH_VIEN_21.accdb"
LOG = r"D:\Users\Dkeiz_Dao\Desktop\access\_rels_log.txt"

app = wc.Dispatch('Access.Application')
app.Visible = False
app.OpenCurrentDatabase(DB, False, False)
try:
    app.DisplayWarnings = False
except Exception:
    pass

vbe = app.VBE
vbproj = vbe.VBProjects(1)

# --- clean up leftover mRelBot* modules from prior hung runs ---
comps = vbproj.VBComponents
for i in range(comps.Count, 0, -1):
    try:
        if comps(i).Name.startswith("mRelBot"):
            comps.Remove(comps(i))
    except Exception:
        pass

mod = comps.Add(1)
mod.Name = "mRelBot%d" % int(time.time())

code = (
"Public Sub CreateRels()\r\n"
"    Dim fso As Object, logFile As Object, db As Object, i As Long\r\n"
"    Set fso = CreateObject(\"Scripting.FileSystemObject\")\r\n"
"    Set logFile = fso.CreateTextFile(\"D:\\Users\\Dkeiz_Dao\\Desktop\\access\\_rels_log.txt\", True, True)\r\n"
"    On Error GoTo ErrH\r\n"
"    logFile.WriteLine \"START \" & Now\r\n"
"    On Error Resume Next\r\n"
"    DoCmd.DeleteRelation \"KHOA_SINHVIEN\"\r\n"
"    DoCmd.DeleteRelation \"SINHVIEN_KETQUA\"\r\n"
"    DoCmd.DeleteRelation \"MONHOC_KETQUA\"\r\n"
"    On Error GoTo ErrH\r\n"
"    DoCmd.AddRelation \"KHOA_SINHVIEN\", 1, \"KHOA\", \"SINHVIEN\", \"MAKH\", \"MAKH\", True\r\n"
"    logFile.WriteLine \"rel1 KHOA_SINHVIEN ok\"\r\n"
"    DoCmd.AddRelation \"SINHVIEN_KETQUA\", 1, \"SINHVIEN\", \"KETQUA\", \"MASV\", \"MASV\", True\r\n"
"    logFile.WriteLine \"rel2 SINHVIEN_KETQUA ok\"\r\n"
"    DoCmd.AddRelation \"MONHOC_KETQUA\", 1, \"MONHOC\", \"KETQUA\", \"MAMH\", \"MAMH\", True\r\n"
"    logFile.WriteLine \"rel3 MONHOC_KETQUA ok\"\r\n"
"    On Error Resume Next\r\n"
"    Set db = CurrentDb()\r\n"
"    logFile.WriteLine \"VERIFY rels count=\" & db.Relations.Count\r\n"
"    For i = 0 To db.Relations.Count - 1\r\n"
"        logFile.WriteLine \"  \" & db.Relations(i).Name & \"  [\" & db.Relations(i).Table & \" -> \" & db.Relations(i).ForeignTable & \"]  type=\" & db.Relations(i).Type\r\n"
"    Next i\r\n"
"    logFile.WriteLine \"ALL DONE\"\r\n"
"    logFile.Close\r\n"
"    Exit Sub\r\n"
"ErrH:\r\n"
"    On Error Resume Next\r\n"
"    logFile.WriteLine \"ERROR num=\" & Err.Number & \" desc=\" & Err.Description\r\n"
"    logFile.Close\r\n"
"End Sub"
)
mod.CodeModule.AddFromString(code)

try:
    app.Run("CreateRels")
    print("app.Run returned OK")
except Exception as e:
    print("RUN ERR: " + str(e))
finally:
    try:
        vbproj.VBComponents.Remove(mod)
    except Exception:
        pass
    app.CloseCurrentDatabase()
print("RELS SCRIPT DONE")
