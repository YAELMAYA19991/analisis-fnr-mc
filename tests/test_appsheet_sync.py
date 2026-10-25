"""Regresión del ciclo de sincronización de auditorías de AppSheet."""
import ast
import copy
import json
import unittest
import uuid
from pathlib import Path
from types import SimpleNamespace

APP=Path(__file__).resolve().parents[1]/"app.py"
TREE=ast.parse(APP.read_text(encoding="utf-8"))
NAMES={
    "order_audit_upload_changed",
    "persist_completed_appsheet_audits",
    "merge_order_audit_records",
    "appsheet_publish_uploaded_batch",
}
NODES=[node for node in TREE.body if isinstance(node,ast.FunctionDef) and node.name in NAMES]
if {node.name for node in NODES}!=NAMES:
    raise AssertionError("Falta una función de sincronización requerida")
COMPILED=compile(ast.Module(body=NODES,type_ignores=[]),str(APP),"exec")


def base_audit(result="Pedido correcto"):
    return {
        "id":"appsheet-audit-123","appsheet_audit_id":"audit-123",
        "pedido":"P-123","slot":"12:00","fecha_auditoria":"2026-10-25T10:00:00Z",
        "auditor":"Supervisora","resultado":result,"pedido_validado":True,
        "observaciones":"Revisión completa","diferencias":0,
        "lineas":[{"Código":"7501","Cantidad pedida":"1"}],
        "fuente_archivo":"AppSheet","appsheet":True,
    }


class SleepRecorder:
    def __init__(self):
        self.calls=[]
    def sleep(self,seconds):
        self.calls.append(seconds)


def environment(conflicts=0,incoming=None):
    saved=[]
    timer=SleepRecorder()
    namespace={
        "json":json,"uuid":uuid,
        "pd":SimpleNamespace(DataFrame=lambda:None),
        "st":SimpleNamespace(session_state={}),
        "appsheet_rows_to_audits":lambda tables,detail:[dict(x) for x in tables.get("_incoming",[])],
        "audit_record_is_complete":lambda row:True,
        "save_store":lambda store:saved.append(copy.deepcopy(store.get("order_audits",[]))),
        "time":timer,
    }
    current={name:[] for name in ("Pedidos","Articulos","Validaciones","Auditorias","Diferencias")}
    payload={"Pedidos":[{"ID_Pedido":"P-123"}],"Articulos":[{"ID_Articulo":"A-123"}]}
    calls={"publish":0}
    def bridge(action,**data):
        if action=="health":
            return {"capabilities":{"replaceBatch":True,"serverFingerprint":True}}
        if action=="read":
            return {"data":copy.deepcopy(current),"fingerprint":"server-fingerprint"}
        if action=="publish":
            calls["publish"]+=1
            if calls["publish"]<=conflicts:
                raise RuntimeError("El lote cambió desde la lectura; no se borró ningún dato.")
            current["Pedidos"]=copy.deepcopy(data["Pedidos"])
            current["Articulos"]=copy.deepcopy(data["Articulos"])
            return {"ok":True}
        raise AssertionError(action)
    namespace.update({
        "appsheet_bridge_call":bridge,
        "appsheet_batch_fingerprint":lambda tables:"incoming-fingerprint",
        "appsheet_order_payload":lambda detail,summary,existing:copy.deepcopy(payload),
        "appsheet_archive_preflight":lambda tables,detail:([],[]),
    })
    exec(COMPILED,namespace)
    return namespace,saved,timer,calls,current


class AppSheetSyncTests(unittest.TestCase):
    def test_archive_metadata_does_not_trigger_repeat_updates(self):
        incoming=base_audit()
        previous={
            **incoming,"lote_id":"archived:abc",
            "appsheet_validaciones_originales":[{"ID_Validacion":"V-1"}],
        }
        ns,saved,_,_,_=environment()
        store={"order_audits":[previous]}
        result=ns["persist_completed_appsheet_audits"](store,{"_incoming":[incoming]})
        self.assertEqual(result,(0,0))
        self.assertEqual(saved,[])
        self.assertEqual(store["order_audits"][0]["lote_id"],"archived:abc")

    def test_changed_appsheet_fields_are_still_persisted(self):
        previous=base_audit()
        previous["lote_id"]="archived:abc"
        incoming=base_audit("Con diferencias")
        ns,saved,_,_,_=environment()
        store={"order_audits":[previous]}
        result=ns["persist_completed_appsheet_audits"](store,{"_incoming":[incoming]})
        self.assertEqual(result,(0,1))
        self.assertEqual(len(saved),1)
        self.assertEqual(store["order_audits"][0]["resultado"],"Con diferencias")
        self.assertEqual(store["order_audits"][0]["lote_id"],"archived:abc")

    def test_upload_callback_pauses_only_while_a_file_is_selected(self):
        ns,_,_,_,_=environment()
        state=ns["st"].session_state
        state["order_audit_upload"]=object()
        ns["order_audit_upload_changed"]()
        self.assertTrue(state["_appsheet_upload_pending"])
        state["order_audit_upload"]=None
        ns["order_audit_upload_changed"]()
        self.assertFalse(state["_appsheet_upload_pending"])

    def test_transient_concurrent_writes_retry_and_complete(self):
        ns,_,timer,calls,_=environment(conflicts=2)
        session={"_appsheet_upload_pending":True}
        result=ns["appsheet_publish_uploaded_batch"](
            None,None,{},force_replace=True,upload_identity="event-1",session_state=session
        )
        self.assertEqual(result["status"],"replaced")
        self.assertEqual(result["attempts"],3)
        self.assertEqual(calls["publish"],3)
        self.assertEqual(timer.calls,[0.75,1.5])
        self.assertEqual(session["_appsheet_last_upload_identity"],"event-1")
        self.assertFalse(session["_appsheet_upload_pending"])


if __name__=="__main__":
    unittest.main()
