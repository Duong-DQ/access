# -*- coding: utf-8 -*-
import time, os
import win32com.client as wc

DB = r"D:\Users\Dkeiz_Dao\Desktop\access\_formtest.accdb"
LOG = r"D:\Users\Dkeiz_Dao\Desktop\access\_uia_recon2.txt"
log = []
try:
    app = wc.Dispatch('Access.Application'); app.Visible = True
    app.OpenCurrentDatabase(DB, False, False)
    log.append("opened db")
    time.sleep(10)
    from pywinauto import Desktop
    log.append("pywinauto OK")
    d = Desktop(backend="uia")
    try:
        allws = d.windows()
        log.append("ALLWINDOWS(%d):" % len(allws))
        for w in allws:
            try:
                log.append("  W: " + (w.element_info.title_text() or ""))
            except Exception:
                pass
    except Exception as e:
        log.append("enum-all failed: %s" % e)
    acc = None
    for w in d.windows():
        try:
            t = w.element_info.title_text() or ""
            if "_formtest" in t or t.startswith("Access"):
                acc = w; log.append("ACC: " + t); break
        except Exception:
            pass
    if acc is not None:
        kws = ["Form", "Wizard", "Report", "Query", "Table", "Macro", "Create", "Relationship"]
        try:
            for c in acc.descendants():
                n = ""
                try: n = c.element_info.name_text() or ""
                except Exception: pass
                if any(k.lower() in n.lower() for k in kws):
                    ct = ""
                    try: ct = c.element_info.control_type() or ""
                    except Exception: pass
                    log.append("[%s] %s" % (ct, n))
        except Exception as e:
            log.append("desc failed: %s" % e)
    else:
        log.append("NO Access window")
finally:
    with open(LOG, "w", encoding="utf-8") as f:
        f.write("\n".join(log))
    try: app.CloseCurrentDatabase()
    except Exception: pass
