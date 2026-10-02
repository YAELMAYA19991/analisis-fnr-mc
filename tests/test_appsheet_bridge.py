"""AppSheet: pruebas de exportación y reimportación sin datos personales."""
import ast
import hashlib
import io
import unittest
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

SOURCE=(Path(__file__).resolve().parents[1]/"app.py").read_text(encoding="utf-8")
tree=ast.parse(SOURCE)
names={"appsheet_audit_frames","appsheet_audit_workbook","import_appsheet_audit_records"}
nodes=[
    node for node in tree.body
    if (
        isinstance(node,ast.FunctionDef) and node.name in names
    ) or (
        isinstance(node,ast.Assign) and any(
            isinstance(target,ast.Name) and target.id=="APPSHEET_AUDIT_HEADERS"
            for target in node.targets
        )
    )
]
env={"pd":pd,"io":io,"hashlib":hashlib,"datetime":datetime,"ZoneInfo":ZoneInfo}
exec(compile(ast.Module(body=nodes,type_ignores=[]),"app.py","exec"),env)


def source_data():
    selected=pd.DataFrame([
        {"Número de pedido":"101","Slot":"08:00","Turno picker":"Nocturno",
         "Pickers asignados":"Picker Ejemplo","Prioridad":"Foco alto",
         "Alarma":"Muy alarmante","Puntaje alarma":17,
         "Factores alarma":"Cruce FNR+MC","Cobertura equipo":"Madrugada"},
        {"Número de pedido":"102","Slot":"09:00","Turno picker":"Mañana",
         "Pickers asignados":"Picker Dos","Prioridad":"Revisar",
         "Alarma":"Alerta","Puntaje alarma":9,
         "Factores alarma":"FNR","Cobertura equipo":"Madrugada"},
    ])
    detail=pd.DataFrame([
        {"Número de pedido":"101","SKU":"S1","Artículo":"Manzana","Picker relacionado":"Picker Ejemplo",
         "Cantidad pedida":"2","Cantidad pickeada":"2"},
        {"Número de pedido":"101","SKU":"S2","Artículo":"Pera","Picker relacionado":"Picker Ejemplo",
         "Cantidad pedida":"1","Cantidad pickeada":"1"},
        {"Número de pedido":"102","SKU":"S3","Artículo":"Plátano","Picker relacionado":"Picker Dos",
         "Cantidad pedida":"4","Cantidad pickeada":"3"},
        {"Número de pedido":"103","SKU":"S4","Artículo":"Excluido","Picker relacionado":"Sin nombre",
         "Cantidad pedida":"1","Cantidad pickeada":"1"},
    ])
    return detail,selected


class AppSheetBridgeTests(unittest.TestCase):
    def test_export_four_sheets_and_order_relations(self):
        detail,selected=source_data()
        frames=env["appsheet_audit_frames"](detail,selected)
        self.assertEqual(set(frames),{"Pedidos","Articulos","Auditorias","Diferencias"})
        self.assertEqual(len(frames["Pedidos"]),2)
        self.assertEqual(len(frames["Articulos"]),3)
        self.assertEqual(set(frames["Articulos"]["ID_Pedido"]),{"101","102"})
        self.assertEqual(len(set(frames["Articulos"]["ID_Articulo"])),3)

    def test_export_is_repeatable(self):
        detail,selected=source_data()
        a=env["appsheet_audit_frames"](detail,selected)
        b=env["appsheet_audit_frames"](detail,selected)
        self.assertEqual(a["Articulos"]["ID_Articulo"].tolist(),b["Articulos"]["ID_Articulo"].tolist())

    def test_workbook_has_headers(self):
        detail,selected=source_data()
        data=env["appsheet_audit_workbook"](detail,selected)
        tabs=pd.read_excel(io.BytesIO(data),sheet_name=None,dtype=str)
        self.assertEqual(set(tabs),{"Pedidos","Articulos","Auditorias","Diferencias"})
        self.assertIn("ID_Auditoria",tabs["Auditorias"].columns)
        self.assertIn("ID_Diferencia",tabs["Diferencias"].columns)

    def test_import_records_only_when_validated_and_no_dupes(self):
        detail,selected=source_data()
        raw=env["appsheet_audit_workbook"](detail,selected)
        book=pd.read_excel(io.BytesIO(raw),sheet_name=None,dtype=str)
        book["Auditorias"]=pd.DataFrame([
            {"ID_Auditoria":"A-1","ID_Pedido":"101","Auditor":"Auditor Ejemplo",
             "Fecha_Auditoria":"2026-10-02 10:00","Validado":"TRUE",
             "Resultado":"Pedido correcto","Observaciones":"","Evidencia":""},
            {"ID_Auditoria":"A-2","ID_Pedido":"102","Auditor":"Auditor Ejemplo",
             "Fecha_Auditoria":"2026-10-02 10:02","Validado":"FALSE",
             "Resultado":"Pendiente","Observaciones":"","Evidencia":""},
        ])
        book["Diferencias"]=pd.DataFrame([
            {"ID_Diferencia":"D-1","ID_Auditoria":"A-1","ID_Articulo":"id_1",
             "Diferencia_Encontrada":"TRUE","Correccion":"Corregir cantidad","Observaciones":"1 pieza"},
        ])
        out=io.BytesIO()
        with pd.ExcelWriter(out,engine="openpyxl") as writer:
            for name,frame in book.items():
                frame.to_excel(writer,sheet_name=name,index=False)
        rows,ignored=env["import_appsheet_audit_records"](out.getvalue(),[])
        self.assertEqual(len(rows),1)
        self.assertEqual(rows[0]["id"],"AS-A-1")
        self.assertEqual(rows[0]["diferencias"],1)
        self.assertEqual(ignored,1)
        again,ignored_again=env["import_appsheet_audit_records"](out.getvalue(),rows)
        self.assertEqual(len(again),0)
        self.assertEqual(ignored_again,2)


if __name__=="__main__":
    unittest.main()
