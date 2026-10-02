"""Pruebas de cobertura por SLOT y cupo porcentual de auditoría."""
import ast
import unittest
from pathlib import Path

import pandas as pd

SOURCE=Path(__file__).resolve().parents[1].joinpath("app.py").read_text(encoding="utf-8")
TREE=ast.parse(SOURCE)
FUNCS={"audit_coverage_filter","audit_hourly_quota_table","audit_hourly_critical_queue"}
NODES=[n for n in TREE.body if isinstance(n,ast.FunctionDef) and n.name in FUNCS]
ENV={"pd":pd}
exec(compile(ast.Module(body=NODES,type_ignores=[]),"app.py","exec"),ENV)


def coverage_samples():
    return pd.DataFrame([
        {"Número de pedido":"nueve","_slot_min":9*60+59,"_slot_max":9*60+59},
        {"Número de pedido":"diez","_slot_min":10*60,"_slot_max":10*60},
        {"Número de pedido":"mixto","_slot_min":9*60+30,"_slot_max":10*60+5},
        {"Número de pedido":"medianoche","_slot_min":0,"_slot_max":0},
        {"Número de pedido":"doce","_slot_min":12*60,"_slot_max":12*60},
    ])


def sample_population(per_hour):
    items=[]
    for slot_hour,num_orders in per_hour.items():
        for i in range(num_orders):
            items.append({
                "Número de pedido":f"{slot_hour}-{i:03d}",
                "Hora auditoría":f"{slot_hour:02d}:00–{slot_hour:02d}:59",
                "Candidato crítico":True,
                "Puntaje selección":float(num_orders-i)+float(i==2),
                "Puntaje alarma":num_orders-i,
                "Máx. historial picker+artículo":2.0,
                "Diferencias cantidad":0,
            })
    return pd.DataFrame(items)


class AuditCoverageTests(unittest.TestCase):
    def test_slot_prior_to_ten(self):
        result=ENV["audit_coverage_filter"](coverage_samples(),"early")
        self.assertEqual(set(result["Número de pedido"]),{"nueve","medianoche"})

    def test_day_team_starts_slot_ten(self):
        result=ENV["audit_coverage_filter"](coverage_samples(),"day")
        self.assertEqual(set(result["Número de pedido"]),{"diez","doce"})

    def test_mixed_slot_not_claimed_by_either_team(self):
        for mode in ("early","day"):
            self.assertNotIn("mixto",set(ENV["audit_coverage_filter"](coverage_samples(),mode)["Número de pedido"]))
        self.assertEqual(len(ENV["audit_coverage_filter"](coverage_samples(),"all")),5)

    def test_fifty_orders_twenty_percent_selects_ten(self):
        population=sample_population({8:50})
        result=ENV["audit_hourly_critical_queue"](population,20,population=population)
        self.assertEqual(len(result),10)
        self.assertEqual(int(result["Ranking crítico hora"].max()),10)
        self.assertEqual(result["Cupo máximo hora"].unique().tolist(),[10])
        self.assertEqual(int(ENV["audit_hourly_quota_table"](population,20)["Cupo máximo"].iloc[0]),10)

    def test_fifty_orders_sixteen_percent_selects_eight(self):
        population=sample_population({8:50})
        self.assertEqual(len(ENV["audit_hourly_critical_queue"](population,16,population=population)),8)

    def test_each_hour_gets_its_own_percentage(self):
        population=sample_population({8:50,9:20})
        result=ENV["audit_hourly_critical_queue"](population,20,population=population)
        self.assertEqual(result["Hora auditoría"].value_counts().to_dict(),
                         {"08:00–08:59":10,"09:00–09:59":4})

    def test_not_filled_with_non_critical(self):
        population=sample_population({8:50})
        available=population.copy()
        available.loc[6:,"Candidato crítico"]=False
        result=ENV["audit_hourly_critical_queue"](available,20,population=population)
        self.assertEqual(len(result),6)

    def test_hiding_audited_does_not_increase_quota(self):
        population=sample_population({8:50})
        available=population.iloc[3:].copy()
        result=ENV["audit_hourly_critical_queue"](available,20,population=population)
        self.assertEqual(len(result),10)
        self.assertNotIn("8-000",set(result["Número de pedido"]))

    def test_round_up_for_small_hours(self):
        population=sample_population({8:3})
        quota=ENV["audit_hourly_quota_table"](population,20)
        self.assertEqual(int(quota["Cupo máximo"].iloc[0]),1)
        result=ENV["audit_hourly_critical_queue"](population,20,population=population)
        self.assertEqual(len(result),1)

    def test_no_critical_orders(self):
        population=sample_population({8:50})
        available=population.copy()
        available["Candidato crítico"]=False
        self.assertEqual(len(ENV["audit_hourly_critical_queue"](available,20,population=population)),0)

    def test_empty_coverage(self):
        empty=sample_population({8:50}).head(0)
        self.assertTrue(ENV["audit_hourly_quota_table"](empty,20).empty)
        self.assertTrue(ENV["audit_hourly_critical_queue"](empty,20,population=empty).empty)

    def test_percentage_bounds(self):
        population=sample_population({8:10})
        with self.assertRaises(ValueError):
            ENV["audit_hourly_quota_table"](population,0)
        with self.assertRaises(ValueError):
            ENV["audit_hourly_quota_table"](population,101)


if __name__=="__main__":
    unittest.main()
