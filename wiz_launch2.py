# -*- coding: utf-8 -*-
import time, threading, win32com.client as wc, win32gui

DB  = r"D:\Users\Dkeiz_Dao\Desktop\access\_formtest.accdb"
LOG = r"D:\Users\Dkeiz_Dao\Desktop\access\_wiz_recon6.txt"
ts  = int(time.time())
lines = []

app = wc.Dispatch('Access.Application')
app.Visible = True
app.OpenCurrentDatabase(DB, False, False)
try:
    app.DisplayWarnings = False
except Exception:
    pass

# Inject a tiny VBA sub that opens the Form Wizard BY NAME (no integer needed)
vbe = app.VBE
vbproj = vbe.VBProjects(1)
for name in list(vbproj.VBComponents):
    if name.name.startswith("mFW"):
        try:
            vbproj.VBComponents.Remove(name)
        except Exception:
            pass
mod = vbproj.VBComponents.Add(1)
mod.Name = "mFW%d" % ts
mod.CodeModule.AddFromString('Sub OpenFWizard()\n\tDoCmd.RunCommand acCmdCreateForm\nEnd Sub')

def dump_wizard():
    time.sleep(5)
    try:
        tops = []
        def tcb(h, _):
            if win32gui.IsWindowVisible(h):
                tops.append((win32gui.GetWindowText(h), win32gui.GetClassName(h), h))
            return True
        win32gui.EnumWindows(tcb, None)
        for title, cls, h in tops:
            lines.append("TOP: %r cls=%r hwnd=%d" % (title, cls, h))
            if 'Wizard' in title:
                kids = []
                def kcb(hh, _):
                    kids.append(hh)
                    return True
                win32gui.EnumChildWindows(h, kcb, None)
                def tree(hh, d):
                    lines.append("  " * d + "[%d] %s | %r | rect=%s" % (
                        hh, win32gui.GetClassName(hh), win32gui.GetWindowText(hh), win32gui.GetWindowRect(hh)))
                    for cc in kids:
                        if win32gui.GetParent(cc) == hh:
                            tree(cc, d + 1)
                lines.append("=== WIZARD TREE (%s) ===" % title)
                for cc in kids:
                    if win32gui.GetParent(cc) == h:
                        tree(cc, 1)
    except Exception as e:
        lines.append("dump error: %s" % e)
    finally:
        try:
            with open(LOG, 'w', encoding='utf-8') as f:
                f.write("\n".join(lines) + "\n")
        except Exception:
            pass

threading.Thread(target=dump_wizard, daemon=True).start()

try:
    app.Run("OpenFWizard")
    lines.append("app.Run OpenFWizard returned (non-blocking)")
except Exception as e:
    lines.append("app.Run error: %s" % e)

time.sleep(6)
try:
    with open(LOG, 'a', encoding='utf-8') as f:
        f.write("\n".join(lines) + "\n")
except Exception:
    pass
try:
    vbproj.VBComponents.Remove(mod)
except Exception:
    pass
os._exit(0)
