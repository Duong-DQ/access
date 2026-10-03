# -*- coding: utf-8 -*-
import threading, time
import win32com.client as wc
import win32gui

DB = r"D:\Users\Dkeiz_Dao\Desktop\access\_formtest.accdb"
LOG = r"D:\Users\Dkeiz_Dao\Desktop\access\_wiz_recon4.txt"

def log(s):
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

def dump_tree(h, depth, out, maxdepth=6):
    out.append("  " * depth + "hwnd=%d cls=%s text=%r rect=%s" % (h, win32gui.GetClassName(h), win32gui.GetWindowText(h), win32gui.GetWindowRect(h)))
    if depth >= maxdepth:
        return
    for ch in enum_children(h):
        dump_tree(ch, depth + 1, out, maxdepth)

def driver():
    wiz = None
    for i in range(160):
        for h in top_windows():
            if win32gui.IsWindowVisible(h) and 'Wizard' in win32gui.GetWindowText(h):
                wiz = h
                break
        if wiz:
            break
        time.sleep(0.5)
    if not wiz:
        log("WIZARD NOT FOUND; visible top-level:")
        for h in top_windows():
            if win32gui.IsWindowVisible(h):
                log("  TOP hwnd=%d %s %r" % (h, win32gui.GetClassName(h), win32gui.GetWindowText(h)))
        return
    log("WIZARD FOUND hwnd=%d title=%r" % (wiz, win32gui.GetWindowText(wiz)))
    tree = []
    dump_tree(wiz, 0, tree, 6)
    for line in tree:
        log("TREE| " + line)
    log("RECON_DONE")
    time.sleep(600)

open(LOG, "w", encoding="utf-8").close()

app = wc.Dispatch('Access.Application')
app.Visible = True
app.OpenCurrentDatabase(DB, False, False)
try:
    app.DisplayWarnings = False
except Exception:
    pass
log("access opened; applying acCmdCreateForm(14)")
t = threading.Thread(target=driver, daemon=True)
t.start()
try:
    app.DoCmd.RunCommand(14)
    log("RunCommand returned (non-blocking)")
except Exception as e:
    log("ApplyCommand error: %r" % e)
time.sleep(600)
