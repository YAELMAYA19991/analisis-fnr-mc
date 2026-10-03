"""Regresión: revisiones incompletas nunca cierran pedidos pendientes."""
import ast
import unittest
from pathlib import Path

SOURCE=Path(__file__).resolve().parents[1].joinpath("app.py").read_text(encoding="utf-8")
TREE=ast.parse(SOURCE)
NODE=next(n for n in TREE.body if isinstance(n,ast.FunctionDef) and n.name=="audit_record_is_complete")
ENV={}
exec(compile(ast.Module(body=[NODE],type_ignores=[]),"app.py","exec"),ENV)
complete=ENV["audit_record_is_complete"]


class AuditPendingTests(unittest.TestCase):
    def test_pending_even_if_legacy_record_was_approved(self):
        self.assertFalse(complete({"resultado":"No se pudo validar","pedido_validado":True}))

    def test_pending_when_checkbox_not_marked(self):
        self.assertFalse(complete({"resultado":"Pedido correcto","pedido_validado":False}))

    def test_finished_correct_order(self):
        self.assertTrue(complete({"resultado":"Pedido correcto","pedido_validado":True}))

    def test_finished_with_differences(self):
        self.assertTrue(complete({"resultado":"Con diferencias","pedido_validado":True}))

    def test_old_confirmed_record_without_flag(self):
        self.assertTrue(complete({"resultado":"Pedido correcto"}))

    def test_old_failed_record_without_flag(self):
        self.assertFalse(complete({"resultado":"No se pudo validar"}))

    def test_bool_as_text_from_exports(self):
        self.assertFalse(complete({"resultado":"Con diferencias","pedido_validado":"false"}))
        self.assertTrue(complete({"resultado":"Con diferencias","pedido_validado":"sí"}))

    def test_unknown_draft_never_marks_complete(self):
        self.assertFalse(complete({}))
        self.assertFalse(complete(None))
        self.assertFalse(complete({"resultado":"Pendiente","pedido_validado":True}))

if __name__=="__main__":
    unittest.main()
