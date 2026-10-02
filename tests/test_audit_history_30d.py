"""Pruebas del registro de FNR y MC de los 30 días anteriores."""
import ast
import unittest
from datetime import datetime, timedelta, date
from pathlib import Path

import numpy as np
import pandas as pd
import re

source=(Path(__file__).resolve().parents[1]/"app.py").read_text(encoding="utf-8")
tree=ast.parse(source)
wanted={"_audit_incident_date","_audit_incident_fingerprint",
        "audit_dedupe_30day_versions","_audit_item_key","norm","person_key",
        "token_key","name_tokens","email_key","extract_email_address"}
nodes=[n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name in wanted]
env={"pd":pd,"np":np,"re":re,"datetime":datetime,"timedelta":timedelta}
exec(compile(ast.Module(body=nodes,type_ignores=[]),"app.py","exec"),env)


def item(date_string,number="123",sku="s1",picker="JUAN PEREZ",quantity=1):
    return {"FECHA":date_string,"ORDER_NUMBER":number,"SKU":sku,
            "PICKER":picker,"CORREO":"","PRODUCTO":"manzana","INCIDENCIAS":quantity,
            "AREA":"Picking","TIPO":"FNR","TURNO":"Nocturno"}


class HistoricIncidenceTests(unittest.TestCase):
    def setUp(self):
        self.today=date(2026,10,2)
        self.start=self.today-timedelta(days=30)

    def dedup(self,*versions):
        return env["audit_dedupe_30day_versions"](
            [pd.DataFrame(v) for v in versions],self.today
        )

    def test_mx_and_iso_dates(self):
        f=env["_audit_incident_date"]
        self.assertEqual(f("01/10/2026"),date(2026,10,1))
        self.assertEqual(f("2026-09-30"),date(2026,9,30))
        self.assertEqual(f("20261001"),date(2026,10,1))
        self.assertEqual(f(""),None)
        self.assertEqual(f(float("nan")),None)
        self.assertEqual(f(46296),date(2026,10,1))

    def test_30_full_days_no_current_day(self):
        frame,stats=self.dedup([
            item(self.start.isoformat(),"primer_dia"),
            item((self.start-timedelta(days=1)).isoformat(),"antes"),
            item((self.today-timedelta(days=1)).isoformat(),"ayer"),
            item(self.today.isoformat(),"hoy"),
            item("","sin_fecha"),
        ])
        self.assertEqual(set(frame.ORDER_NUMBER),{"primer_dia","ayer"})
        self.assertEqual(stats["fechas_ausentes"],1)
        self.assertEqual(stats["fuera_de_periodo"],2)
        self.assertEqual(stats["registros_unicos"],2)

    def test_reuploaded_file_never_doubles(self):
        report=[item("2026-10-01")]
        frame,stats=self.dedup(report,report,report)
        self.assertEqual(len(frame),1)
        self.assertEqual(stats["versiones"],3)

    def test_two_same_rows_inside_export_are_preserved(self):
        same=item("2026-10-01")
        frame,stats=self.dedup([same,same],[same],[same,same])
        self.assertEqual(len(frame),2)
        self.assertEqual(stats["registros_unicos"],2)

    def test_incremental_different_orders_are_both_kept(self):
        frame,stats=self.dedup(
            [item("2026-10-01","A")],
            [item("2026-10-01","B")],
        )
        self.assertEqual(set(frame.ORDER_NUMBER),{"A","B"})

    def test_different_date_same_order_are_distinct(self):
        frame,_=self.dedup(
            [item("2026-09-30","A")],
            [item("2026-10-01","A")],
        )
        self.assertEqual(len(frame),2)


if __name__=="__main__":
    unittest.main()
