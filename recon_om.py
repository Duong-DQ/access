# -*- coding: utf-8 -*-
import time
from collections import Counter
import win32gui, win32con
import win32com.client as wc

DB  = r"D:\Users\Dkeiz_Dao\Desktop\access\_formtest.accdb"
LOG = r"D:\Users\Dkeiz_Dao\Desktop\access\_om_recon.txt"

app = wc.Dispatch('Access.Application')
app.Visible = True
app.OpenCurrentDatabase(DB, False, False)
try:
    app.DisplayWarnings = False
except Exception:
    pass

omain = None
for _ in range(60):
    omain = win32gui.FindWindow("OMain", None)
    if omain:
        break
    time.sleep(0.5)

L = []
L.append("OMain=%s" % (omain,))
desc = []

def cb(h, _):
    desc.append(h)
    return True

if omain:
    win32gui.SetForegroundWindow(omain)
    time.sleep(1)
    win32gui.EnumChildWindows(omain, cb, None)
    L.append("descendants=%d" % len(desc))
    hist = Counter(win32gui.GetClassName(h) for h in desc)
    L.append("CLASSES: " + " ".join("%s:%d" % (k, v) for k, v in hist.most_common()))
    L.append("--- windows with text ---")
    for h in desc:
        txt = win32gui.GetWindowText(h)
        if txt:
            L.append("hwnd=%d cls=%s rect=%s text=%r" % (h, win32gui.GetClassName(h), str(win32gui.GetWindowRect(h)), txt[:50]))
    L.append("--- interactive (tree/list/combo/button/edit) ---")
    for h in desc:
        cls = win32gui.GetClassName(h)
        if cls in ("SysTreeView32", "SysListView32", "ComboBox", "ListBox", "Button", "Edit", "RICHEDIT50W", "NetUIHWND"):
            L.append("hwnd=%d cls=%s rect=%s text=%r" % (h, cls, str(win32gui.GetWindowRect(h)), win32gui.GetWindowText(h)[:40]))
open(LOG, "w", encoding="utf-8").write("\n".join(L))
time.sleep(900)
