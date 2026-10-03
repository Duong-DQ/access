# -*- coding: utf-8 -*-
import time, os
import win32com.client as wc

LOG = r"D:\Users\Dkeiz_Dao\Desktop\access\_uia_recon.txt"
open(LOG, "w", encoding="utf-8").close()
def log(m):
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(m + "\n")

DB = r"D:\Users\Dkeiz_Dao\Desktop\access\_formtest.accdb"
app = wc.Dispatch('Access.Application')
app.Visible = True
app.OpenCurrentDatabase(DB, False, False)
time.sleep(4)

try:
    from pywinauto import Desktop
    log("pywinauto OK")
    d = Desktop(backend="uia")
    acc = None
    for w in d.windows():
        try:
            t = w.window_text()
        except Exception:
            t = ""
        if "_formtest" in str(t) or (str(t).startswith("Access")):
            log("WIN: %r" % t)
            if acc is None:
                acc = w
    if acc is None:
        log("NO Access window")
    else:
        # enumerate all descendants, report those whose name looks like a ribbon button
        kw = ("Form", "Wizard", "Report", "Query", "Table", "Macro", "Create", "Relationship")
        try:
            descs = acc.descendants()
        except Exception as e:
            descs = []
            log("descendants err: %r" % e)
        log("total descendants: %d" % len(descs))
        hits = 0
        for c in descs:
            try:
                ei = c.element_info
                nm = ei.name or ""
                ct = ei.control_type
            except Exception:
                continue
            if any(k in str(nm) for k in kw):
                log("  [%s] %r" % (ct, nm))
                hits += 1
                if hits > 80:
                    log("  ...truncated")
                    break
        log("hits: %d" % hits)
except Exception as e:
    log("EXC: %r" % e)

os._exit(0)
