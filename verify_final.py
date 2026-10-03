# -*- coding: utf-8 -*-
import time
import win32com.client as wc

DB = r"D:\Users\Dkeiz_Dao\Desktop\access\QL_SINH_VIEN_21.accdb"
OUT = r"D:\Users\Dkeiz_Dao\Desktop\access\_verify_final.txt"

VBA = (
    "Option Explicit\r\n"
    "Public Sub Verify()\r\n"
    "  On Error GoTo EH\r\n"
    "  Dim fso As Object, lf As Object, i As Long, r As Object\r\n"
    "  Set fso = CreateObject(\"Scripting.FileSystemObject\")\r\n"
    "  Set lf = fso.CreateTextFile(\"" + OUT + "\", True, True)\r\n"
    "  lf.WriteLine \"START\"\r\n"
    "  For i = 0 To CurrentDb.Relations.Count - 1\r\n"
    "    Set r = CurrentDb.Relations(i)\r\n"
    "    lf.WriteLine \"REL: \" & r.Name & \" | \" & r.Table & \" > \" & r.ForeignTable & \" | type=\" & r.Type\r\n"
    "  Next i\r\n"
    "  For i = 0 To CurrentProject.AllForms.Count - 1\r\n"
    "    lf.WriteLine \"FORM: \" & CurrentProject.AllForms(i).Name\r\n"
    "  Next i\r\n"
    "  lf.WriteLine \"DONE\"\r\n"
    "  lf.Close\r\n"
    "  Exit Sub\r\n"
    "EH:\r\n"
    "  On Error Resume Next\r\n"
    "  If Not lf Is Nothing Then lf.WriteLine \"ERR \" & Err.Number & \" \" & Err.Description\r\n"
    "  If Not lf Is Nothing Then lf.Close\r\n"
    "End Sub\r\n"
)

app = wc.Dispatch('Access.Application')
app.Visible = False
app.OpenCurrentDatabase(DB, False, False)
try:
    vbe = app.VBE
    vbproj = app.VBE.VBProjects(1)
    ts = str(int(time.time()))
    for m in list(vbproj.VBComponents):
        if m.Name.startswith('mVerify'):
            vbproj.VBComponents.Remove(m)
    mod = vbproj.VBComponents.Add(1)
    mod.Name = 'mVerify' + ts
    mod.CodeModule.AddFromString(VBA)
    app.Run('Verify')
    try:
        vbproj.VBComponents.Remove(mod)
    except Exception:
        pass
finally:
    try:
        app.CloseCurrentDatabase()
    except Exception:
        pass
print("verify run complete")
