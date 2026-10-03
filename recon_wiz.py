# -*- coding: utf-8 -*-
import win32gui

out = []

def info(h):
    cls = win32gui.GetClassName(h)
    txt = win32gui.GetWindowText(h)
    r = win32gui.GetWindowRect(h)
    v = win32gui.IsWindowVisible(h)
    return cls, txt, r, v

def enum_children(parent):
    res = []
    def cb(h, _):
        res.append(h)
        return True
    win32gui.EnumChildWindows(parent, cb, None)
    return res

tops = []
def topcb(h, _):
    tops.append(h)
    return True
win32gui.EnumWindows(topcb, None)

out.append("=== TOP-LEVEL (visible) ===")
for h in tops:
    if win32gui.IsWindowVisible(h):
        c, t, r, v = info(h)
        out.append("hwnd=%d cls=%s title=%r rect=%s" % (h, c, t, r))

def find_wizard():
    for h in tops:
        if not win32gui.IsWindowVisible(h):
            continue
        if 'Wizard' in win32gui.GetWindowText(h):
            return h
    return win32gui.GetForegroundWindow()

def dump_tree(h, depth, maxdepth):
    c, t, r, v = info(h)
    out.append("  " * depth + "hwnd=%d cls=%s text=%r rect=%s vis=%d" % (h, c, t, r, v))
    if depth >= maxdepth:
        return
    for ch in enum_children(h):
        dump_tree(ch, depth + 1, maxdepth)

wiz = find_wizard()
out.append("=== WIZARD TREE (hwnd=%d) ===" % (wiz or 0))
if wiz:
    dump_tree(wiz, 0, 4)

with open(r"D:\Users\Dkeiz_Dao\Desktop\access\_wiz_recon.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("done %d lines; wizard hwnd=%s" % (len(out), wiz))
