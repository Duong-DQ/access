# -*- coding: utf-8 -*-
import time, os, threading
import win32com.client as wc
import win32gui

DB = r"D:\Users\Dkeiz_Dao\Desktop\access\_formtest.accdb"
LOG = r"D:\Users\Dkeiz_Dao\Desktop\access\_wiz_recon7.txt"

def recon():
    time.sleep(6)
    lines = []
    tops = []
    def tcb(h, _):
        tops.append(h); return True
    win32gui.EnumWindows(tcb, None)
    for h in tops:
        if win32gui.IsWindowVisible(h):
            lines.append("TOP hwnd=%d cls=%r title=%r" % (h, win32gui.GetClassName(h), win32gui.GetWindowText(h)))
    for h in tops:
        if not win32gui.IsWindowVisible(h):
            continue
        t = win32gui.GetWindowText(h)
        if 'Wizard' in t or t == 'Microsoft Visual Basic':
            lines.append("=== TREE of %r cls=%s ===" % (t, win32gui.GetClassName(h)))
            children = []
            def ccb(ch, _):
                children.append(ch); return True
            win32gui.EnumChildWindows(h, ccb, None)
            for ch in children:
                lines.append("  hwnd=%d cls=%r txt=%r rect=%s" % (ch, win32gui.GetClassName(ch), win32gui.GetWindowText(ch), win32gui.GetWindowRect(ch)))
    with open(LOG, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

t = threading.Thread(target=recon, daemon=True)
t.start()

app = wc.Dispatch('Access.Application')
app.Visible = True
app.OpenCurrentDatabase(DB, False, False)
try:
    app.DisplayWarnings = False
except Exception:
    pass

ts = str(int(time.time()))
vbe = app.VBE
vbproj = vbe.VBProjects(1)
for comp in list(vbproj.VBComponents):
    if comp.Name.startswith("mFW2") or comp.Name.startswith("mDiag"):
        vbproj.VBComponents.Remove(comp)
mod = vbproj.VBComponents.Add(1)
mod.Name = "mFW2" + ts
vba = ('Public Sub OpenFWizard()\n'
       '  Dim fso As Object, f As Object\n'
       '  On Error GoTo EH\n'
       '  Set fso = CreateObject("Scripting.FileSystemObject")\n'
       '  Set f = fso.CreateTextFile("D:\\Users\\Dkeiz_Dao\\Desktop\\access\\_wiz_err2.txt", True, True)\n'
       '  f.WriteLine "STEP OpenTable MONHOC"\n'
       '  DoCmd.OpenTable "MONHOC"\n'
       '  f.WriteLine "STEP OpenTable OK -> RunCommand"\n'
       '  DoCmd.RunCommand acCmdCreateForm\n'
       '  f.WriteLine "RunCommand returned (non-blocking)"\n'
       '  f.Close\n'
       '  Exit Sub\n'
       'EH:\n'
       '  f.WriteLine "ERR " & Err.Number & " " & Err.Description\n'
       '  f.Close\n'
       'End Sub\n')
mod.CodeModule.AddFromString(vba)
app.Run("OpenFWizard")
time.sleep(6)
try:
    for comp in list(vbproj.VBComponents):
        if comp.Name.startswith("mFW2") or comp.Name.startswith("mDiag"):
            vbproj.VBComponents.Remove(comp)
except Exception as e:
    print("cleanup:", e)
os._exit(0)
