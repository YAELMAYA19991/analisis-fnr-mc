import ast
from pathlib import Path
import unittest

APP=Path(__file__).resolve().parents[1]/"app.py"

class AppStaticTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.text=APP.read_text(encoding="utf-8")
        cls.tree=ast.parse(cls.text)

    def test_core_functions_exist(self):
        names={n.name for n in ast.walk(self.tree) if isinstance(n,(ast.FunctionDef,ast.AsyncFunctionDef))}
        required={
            "prepare_source_frames","summary","order_risk_analysis",
            "monthly_picker_comparison","build_attendance_summary",
            "save_store","load_store",
        }
        self.assertTrue(required.issubset(names),required-names)

    def test_modern_lazy_tabs_enabled(self):
        self.assertIn('st.tabs(_tab_labels,on_change="rerun"',self.text)
        self.assertNotIn("except TypeError:\n    a,o,b,i,h,j,x=st.tabs",self.text)

    def test_secure_smtp(self):
        self.assertIn('if not smtp.has_extn("starttls")',self.text)

    def test_stable_identity_present(self):
        self.assertIn("person_context_key",self.text)
        self.assertIn("CORREO_KEY",self.text)

    def test_store_shards_present(self):
        self.assertIn("STORE_SHARDS={",self.text)
        self.assertIn('"order_audits"',self.text)
        self.assertIn('"monthly_history"',self.text)

    def test_compact_order_audit_workflow(self):
        self.assertIn("Auditoría rápida de pedidos",self.text)
        self.assertIn("1. Cargar pedidos (Excel o CSV)",self.text)
        self.assertIn("3. Seleccionar pedido",self.text)
        self.assertIn("Guardar revisión",self.text)
        self.assertIn("📚 Historial de auditorías",self.text)

    def test_hourly_critical_audit_selection(self):
        self.assertIn("def audit_hourly_critical_queue",self.text)
        self.assertIn("Candidato crítico",self.text)
        self.assertIn("Ranking crítico hora",self.text)
        self.assertIn("Puntaje alarma",self.text)
        self.assertIn("Sugeridos primero",self.text)

    def test_pending_audit_logic_kept(self):
        self.assertIn("def audit_record_is_complete",self.text)
        self.assertIn("Motivo por el que quedó pendiente",self.text)
        self.assertIn("30 días",self.text)
        self.assertIn('"Estado":"Validado" if audit_record_is_complete(rec) else "Pendiente"',self.text)

if __name__=="__main__":
    unittest.main()
