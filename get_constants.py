# -*- coding: utf-8 -*-
import time
import win32com.client as wc

DB = r"D:\Users\Dkeiz_Dao\Desktop\access\_formtest.accdb"
app = wc.Dispatch('Access.Application')
app.Visible = False
app.OpenCurrentDatabase(DB, False, False)
vbe = app.VBE
vbproj = vbe.VBProjects(1)

# remove any leftover mAcCmd* modules from prior runs
for comp in list(vbproj.VBComponents):
    if comp.Name.startswith('mAcCmd'):
        try:
            vbproj.VBComponents.Remove(comp)
        except Exception:
            pass

mod = vbproj.VBComponents.Add(1)
try:
    mod.Name = 'mAcCmd' + str(int(time.time()))
except Exception:
    pass

code = (
    'Public Sub DumpAcCmds()\r\n'
    '  Dim fso As Object, f As Object\r\n'
    '  Set fso = CreateObject("Scripting.FileSystemObject")\r\n'
    '  Set f = fso.CreateTextFile("D:\\Users\\Dkeiz_Dao\\Desktop\\access\\_accmd.txt", True, True)\r\n'
    '  f.WriteLine "createForm=" & acCmdCreateForm\r\n'
    '  f.WriteLine "createQuery=" & acCmdCreateQuery\r\n'
    '  f.WriteLine "createReport=" & acCmdCreateReport\r\n'
    '  f.WriteLine "createTable=" & acCmdCreateTable\r\n'
    '  f.WriteLine "makeMDE=" & acCmdMakeMDE\r\n'
    '  f.WriteLine "delete=" & acCmdDelete\r\n'
    '  f.WriteLine "dataSheetView=" & acCmdDataSheetView\r\n'
    '  f.WriteLine "formView=" & acCmdFormView\r\n'
    '  f.WriteLine "done"\r\n'
    '  f.Close\r\n'
    'End Sub\r\n'
)
mod.CodeModule.AddFromString(code)
app.Run("DumpAcCmds")
try:
    vbproj.VBComponents.Remove(mod)
except Exception:
    pass
app.CloseCurrentDatabase()
print("OK constants dumped")
