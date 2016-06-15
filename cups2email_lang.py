#!/usr/bin/env python
# -*- coding: utf-8 -*-

from ooop import OOOP
import sys
import configdb
from openpyxl import load_workbook

def cups2email(template=None,filename=None):
    if not filename:
        sys.exit("No filename defined")
    if not template:
        sys.exit("No template defined")
    wb = load_workbook(filename)
    ws = wb.active
    O = OOOP(**configdb.ooop)
    pol_obj = O.GiscedataPolissa
    cups_obj = O.GiscedataCupsPs
    templ_obj=O.PoweremailTemplates
    templ_id=templ_obj.search([('name','=',template)])
#Skip header row
    row=2
    while ws.cell(row=row,column=1).value:
        cups=ws.cell(row=row,column=1).value
        print cups
        try:
            cups_id = cups_obj.search([('name','=',cups)])[0]
        except:
            print "El CUPS no existe"
            continue
        try:
            pol_id = pol_obj.search([('cups','=',cups_id)])[0]
        except:
            print "No hay pólizas activas asociadas al CUPS"
            continue
        pol=pol_obj.get(pol_id)
        print pol.pagador.lang
        print pol.pagador.www_email
        ctx = {'active_ids': [pol.pagador.id],
              'active_id': pol.pagador.id,
              'template_id': templ_id,
              'src_model': 'giscedata.polissa',
              'src_rec_ids': [pol.pagador.id],
              'from': 1}
        params = {'state': 'single',
                 'priority':0,
                 'from': ctx['from']}
        try:
           wz_id = O.PoweremailSendWizard.create(params, ctx)
           O.PoweremailSendWizard.send_mail([wz_id], ctx)
        except Exception:
           raise
        row+=1
