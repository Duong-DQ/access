# -*- coding: utf-8 -*-
import os, time
import win32com.client as wc

DB = r"D:\Users\Dkeiz_Dao\Desktop\access\_formtest.accdb"
OUT = r"D:\Users\Dkeiz_Dao\Desktop\access\_wiz_err.txt"
ts = int(time.time())
if os.path.exists(OUT):
    os.remove(OUT)

vba = (
    "Sub OpenFWizard()\r\n"
    "  On Error GoTo EH\r\n"
    "  DoCmd.RunCommand acCmdCreateForm\r\n"
    "  Exit Sub\r\n"
    "EH:\r\n"
    "  Dim f As Object, tf As Object\r\n"
    "  Set f = CreateObject(\"Scripting.FileSystemObject\")\r\n"
    "  Set tf = f.CreateTextFile(\"D:\\Users\\Dkeiz_Dao\\Desktop\\access\\_wiz_err.txt\", True, True)\r\n"
    "  tf.WriteLine \"ERR num=\" & Err.Number & \" desc=\" & Err.Description\r\n"
    "  tf.Close\r\n"
    "End Sub\r\n"
)

app = wc.Dispatch('Access.Application')
app.Visible = True
app.OpenCurrentDatabase(DB, False, False)
try:
    app.DisplayWarnings = False
except Exception:
    pass
vbe = app.VBE
vbproj = vbe.VBProjects(1)
for m in list(vbproj.VBComponents):
    if m.Name.lower().startswith('mfw'):
        try:
            vbproj.VBComponents.Remove(m)
        except Exception:
            pass
mod = vbproj.VBComponents.Add(1)
mod.Name = "mFW" + str(ts)
mod.CodeModule.AddFromString(vba)
app.Run("OpenFWizard")
time.sleep(3)
try:
    vbproj.VBComponents.Remove(mod)
except Exception:
    pass
try:
    app.CloseCurrentDatabase()
except Exception:
    pass
print("DONE")
