"""Pruebas de la cobertura por SLOT y del máximo crítico por hora."""
import ast
import unittest
from pathlib import Path

import pandas as pd

SOURCE=Path(__file__).resolve().parents[1].joinpath("app.py").read_text(encoding="utf-8")
TREE=ast.parse(SOURCE)
FUNCS={"audit_coverage_filter","audit_hourly_critical_queue"}
NODES=[n for n in TREE.body if isinstance(n,ast.FunctionDef) and n.name in FUNCS]
ENV={"pd":pd}
exec(compile(ast.Module(body=NODES,type_ignores=[]),"app.py","exec"),ENV)

def sample():
    return pd.DataFrame([
        {"Número de pedido":"nueve","_slot_min":9*60+59,"_slot_max":9*60+59},
        {"Número de pedido":"diez","_slot_min":10*60,"_slot_max":10*60},
        {"Número de pedido":"medio","_slot_min":9*60+30,"_slot_max":10*60+5},
        {"Número de pedido":"medianoche","_slot_min":0,"_slot_max":0},
        {"Número de pedido":"doce","_slot_min":12*60,"_slot_max":12*60},
    ])

class AuditCoverageTests(unittest.TestCase):
    def test_slot_prior_to_ten(self):
        result=ENV["audit_coverage_filter"](sample(),"early")
        self.assertEqual(set(result["Número de pedido"]),{"nueve","medianoche"})

    def test_day_team_starts_slot_ten(self):
        result=ENV["audit_coverage_filter"](sample(),"day")
        self.assertEqual(set(result["Número de pedido"]),{"diez","doce"})

    def test_mixed_slot_not_claimed_by_either_team(self):
        for mode in ("early","day"):
            self.assertNotIn("medio",set(ENV["audit_coverage_filter"](sample(),mode)["Número de pedido"]))
        self.assertEqual(len(ENV["audit_coverage_filter"](sample(),"all")),5)

    def test_highest_four_per_slot_hour(self):
        items=[]
        for i in range(6):
            items.append({
                "Número de pedido":str(i),"Hora auditoría":"08:00–08:59",
                "Candidato crítico":True,"Puntaje alarma":20-i,
                "Puntaje selección":20-i + (2 if i==2 else 0),
                "Máx. historial picker+artículo":1,
                "Diferencias cantidad":0,
            })
        items.append({
            "Número de pedido":"nine","Hora auditoría":"09:00–09:59",
            "Candidato crítico":True,"Puntaje alarma":22,"Puntaje selección":22,
            "Máx. historial picker+artículo":1,"Diferencias cantidad":0,
        })
        result=ENV["audit_hourly_critical_queue"](pd.DataFrame(items),4)
        self.assertEqual(len(result),5)
        self.assertEqual(sum(result["Hora auditoría"]=="08:00–08:59"),4)
        self.assertIn("nine",set(result["Número de pedido"]))
        eight=result[result["Hora auditoría"]=="08:00–08:59"]
        self.assertEqual(int(eight["Ranking crítico hora"].max()),4)

    def test_never_fill_quota_with_non_critical(self):
        items=pd.DataFrame([
            {"Número de pedido":str(i),"Hora auditoría":"07:00–07:59",
             "Candidato crítico":i==0,"Puntaje alarma":20-i,
             "Puntaje selección":20-i,
             "Máx. historial picker+artículo":1,"Diferencias cantidad":0}
            for i in range(8)
        ])
        self.assertEqual(len(ENV["audit_hourly_critical_queue"](items,4)),1)
        self.assertEqual(len(ENV["audit_hourly_critical_queue"](items,3)),1)

if __name__=="__main__":
    unittest.main()
