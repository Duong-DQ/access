# -*- coding: utf-8 -*-
import time, os, sys
import win32gui
import win32com.client as wc

DB = r"D:\Users\Dkeiz_Dao\Desktop\access\_formtest.accdb"
AC = int(sys.argv[1]) if len(sys.argv) > 1 else 14
LOG = r"D:\Users\Dkeiz_Dao\Desktop\access\_wiz_recon5.txt"

lines = ["AC_CMD=%d" % AC, "start"]

app = wc.Dispatch('Access.Application')
app.Visible = True
app.OpenCurrentDatabase(DB, False, False)
try:
    app.DisplayWarnings = False
except Exception:
    pass
try:
    app.DoCmd.RunCommand(AC)
    lines.append("RunCommand returned (non-blocking)")
except Exception as e:
    lines.append("RunCommand EXC: %r" % (e,))
time.sleep(6)


def enum_children(parent):
    out = []

    def cb(h, _):
        out.append(h)
        return True
    win32gui.EnumChildWindows(parent, cb, None)
    return out


def info(h):
    return (win32gui.GetClassName(h), win32gui.GetWindowText(h), win32gui.GetWindowRect(h))


def dump_tree(h, depth, out, maxdepth=7):
    if depth > maxdepth:
        return
    for ch in enum_children(h):
        try:
            cls, t, r = info(ch)
            out.append(("  " * depth) + "hwnd=%d %s | %r | %s" % (ch, cls, t, r))
            dump_tree(ch, depth + 1, out, maxdepth)
        except Exception:
            pass


tops = []


def tc(h, _):
    tops.append(h)
    return True


win32gui.EnumWindows(tc, None)
lines.append("=== TOP-LEVEL (%d) ===" % len(tops))
interesting = []
for h in tops:
    try:
        if not win32gui.IsWindowVisible(h):
            continue
        t = win32gui.GetWindowText(h)
        c = win32gui.GetClassName(h)
        lines.append("TOP hwnd=%d class=%r title=%r" % (h, c, t))
        tl = (t or "").lower()
        if any(k in tl for k in ("wizard", "create", "form")) or "access" in c.lower() or "mswin" in c.lower():
            interesting.append(h)
    except Exception:
        pass

for th in interesting:
    lines.append("=== TREE hwnd=%d title=%r ===" % (th, win32gui.GetWindowText(th)))
    dump_tree(th, 0, lines)

with open(LOG, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))
os._exit(0)
