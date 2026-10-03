# -*- coding: utf-8 -*-
import time, win32com.client as wc
DB = r"D:\Users\Dkeiz_Dao\Desktop\access\QL_SINH_VIEN_21.accdb"
OUT = r"D:\Users\Dkeiz_Dao\Desktop\access\_verify2.txt"
app = wc.Dispatch('Access.Application'); app.Visible = False
app.OpenCurrentDatabase(DB, False, False)
try: app.DisplayWarnings = False
except: pass
vbe = app.VBE; vbproj = vbe.VBProjects(1)
for n in list(vbproj.VBComponents):
    if n.Name.startswith("mVrfy2"):
        vbproj.VBComponents.Remove(n)
mod = vbproj.VBComponents.Add(1); mod.Name = "mVrfy2"
code = '''Public Sub Vrfy2()
  On Error GoTo EH
  Dim fso As Object, lf As Object, i As Long
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set lf = fso.CreateTextFile("''' + OUT + '''", True, True)
  lf.WriteLine "START"
  On Error Resume Next
  Dim fc As Long: fc = CurrentProject.AllForms.Count
  lf.WriteLine "FORMS_COUNT=" & fc
  For i = 0 To fc - 1
    If Err.Number = 0 Then lf.WriteLine "FORM: " & CurrentProject.AllForms(i).Name
  Next i
  lf.WriteLine "FORMS_ERR=" & Err.Number
  On Error GoTo EH
  On Error Resume Next
  Dim rc As Long: rc = CurrentDb.Relations.Count
  lf.WriteLine "RELS_COUNT=" & rc
  For i = 0 To rc - 1
    If Err.Number = 0 Then lf.WriteLine "REL: " & CurrentDb.Relations(i).Name & " | " & CurrentDb.Relations(i).Table & " > " & CurrentDb.Relations(i).ForeignTable
  Next i
  lf.WriteLine "RELS_ERR=" & Err.Number
  On Error GoTo 0
  lf.WriteLine "DONE"
  lf.Close
  Exit Sub
EH:
  On Error Resume Next
  If Not lf Is Nothing Then
    lf.WriteLine "FATAL num=" & Err.Number & " desc=" & Err.Description
    lf.Close
  End If
End Sub'''
mod.CodeModule.AddFromString(code)
app.Run("Vrfy2")
try:
    for n in list(vbproj.VBComponents):
        if n.Name.startswith("mVrfy2"):
            vbproj.VBComponents.Remove(n)
except: pass
try: app.CloseCurrentDatabase()
except: pass
print("done")
