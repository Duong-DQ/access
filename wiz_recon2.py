# -*- coding: utf-8 -*-
import threading, time
import win32com.client as wc
import win32gui

DB = r"D:\Users\Dkeiz_Dao\Desktop\access\_formtest.accdb"
LOG = r"D:\Users\Dkeiz_Dao\Desktop\access\_wiz_recon2.txt"
loglines = []

def log(s):
    loglines.append(s)
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(s + "\n")
    except Exception:
        pass

def top_windows():
    res = []
    win32gui.EnumWindows(lambda h, _: (res.append(h), True)[1], None)
    return res

def enum_children(parent):
    res = []
    win32gui.EnumChildWindows(parent, lambda h, _: (res.append(h), True)[1], None)
    return res

def dump_tree(h, depth, out, maxdepth=5):
    cls = win32gui.GetClassName(h)
    txt = win32gui.GetWindowText(h)
    r = win32gui.GetWindowRect(h)
    out.append("  " * depth + "cls=%s text=%r rect=%s" % (cls, txt, r))
    if depth >= maxdepth:
        return
    for ch in enum_children(h):
        dump_tree(ch, depth + 1, out, maxdepth)

def driver():
    wiz = None
    for i in range(150):  # up to 75s
        for h in top_windows():
            if win32gui.IsWindowVisible(h) and 'Wizard' in win32gui.GetWindowText(h):
                wiz = h
                break
        if wiz:
            break
        time.sleep(0.5)
    if not wiz:
        log("WIZARD NOT FOUND after polling")
        # dump all visible top-level for diagnosis
        for h in top_windows():
            if win32gui.IsWindowVisible(h):
                log("  TOP %s %r" % (win32gui.GetClassName(h), win32gui.GetWindowText(h)))
        return
    log("WIZARD FOUND hwnd=%d title=%r" % (wiz, win32gui.GetWindowText(wiz)))
    tree = []
    dump_tree(wiz, 0, tree, 5)
    for line in tree:
        log("TREE| " + line)
    log("RECON_DONE")
    time.sleep(600)

# fresh log
open(LOG, "w", encoding="utf-8").close()

app = wc.Dispatch('Access.Application')
app.Visible = True
app.OpenCurrentDatabase(DB, False, False)
try:
    app.DisplayWarnings = False
except Exception:
    pass
log("access opened, launching wizard")
t = threading.Thread(target=driver, daemon=True)
t.start()
app.DoCmd.RunCommand(233)  # acCmdCreateForm; may block
time.sleep(600)
