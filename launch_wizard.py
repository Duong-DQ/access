# -*- coding: utf-8 -*-
import time
import win32com.client as wc

DB = r"D:\Users\Dkeiz_Dao\Desktop\access\_formtest.accdb"
app = wc.Dispatch('Access.Application')
app.Visible = True
app.OpenCurrentDatabase(DB, False, False)
try:
    app.DisplayWarnings = False
except Exception:
    pass
# acCmdCreateForm = 233
app.DoCmd.RunCommand(233)
# keep process + Access alive for inspection (if RunCommand returned)
for _ in range(600):
    time.sleep(1)
