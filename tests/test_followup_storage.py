"""Pruebas de persistencia de seguimientos sin conexión a Supabase real."""
import ast
import json
import re
import uuid
import unittest
from pathlib import Path

APP=Path(__file__).resolve().parents[1]/"app.py"
SOURCE=APP.read_text(encoding="utf-8")
TREE=ast.parse(SOURCE)
NAMES={
    "norm","person_key","name_tokens","token_key","picker_record",
    "_unique_records","_merge_picker_store","_followup_event_objects",
    "_load_followup_journal","save_followup_entry",
}
NODES=[node for node in TREE.body if isinstance(node,ast.FunctionDef) and node.name in NAMES]
COMPILED=compile(ast.Module(body=NODES,type_ignores=[]),str(APP),"exec")

def environment(*,fail_index=False,cloud_on=True):
    objects={}
    index_calls=[]
    namespace={
        "json":json,"re":re,"uuid":uuid,
        "FOLLOWUP_JOURNAL_PREFIX":"seguimientos_registros/",
        "cloud_enabled":lambda:cloud_on,
    }
    def upload(key,data,content_type):
        assert key.startswith("seguimientos_registros/")
        objects[key]=data
    def list_events(prefix,limit=500,offset=0):
        names=[{"name":key[len(prefix):]} for key in sorted(objects) if key.startswith(prefix)]
        return names[offset:offset+limit]
    def get_json(key,default):
        return json.loads(objects[key].decode("utf-8")) if key in objects else default
    def save_index(store):
        index_calls.append(True)
        if fail_index:
            raise IOError("simulación: índice temporalmente no disponible")
    namespace.update({
        "cloud_upload":upload,
        "cloud_list":list_events,
        "_cloud_json_read":get_json,
        "save_store":save_index,
    })
    exec(COMPILED,namespace)
    return namespace,objects,index_calls

class FollowupPersistenceTests(unittest.TestCase):
    def test_followups_saved_individually(self):
        ns,objects,calls=environment()
        state={"pickers":{}}
        for i in range(2):
            ns["save_followup_entry"](
                state,"PEREZ GARDUÑO ALDAIR","acciones",
                {"id":f"seg-{i}","fecha":"2026-10-01 20:30","motivo":str(i)},
            )
        self.assertEqual(len(objects),2)
        self.assertEqual(len(calls),2)
        self.assertEqual(len(state["pickers"]["PEREZ GARDUÑO ALDAIR"]["acciones"]),2)

    def test_recovery_from_journal_after_index_failure(self):
        ns,objects,_=environment(fail_index=True)
        state={"pickers":{}}
        with self.assertRaisesRegex(RuntimeError,"respaldado individualmente"):
            ns["save_followup_entry"](
                state,"PEREZ GARDUÑO ALDAIR","documentos",
                {"id":"doc-1","tipo":"2. Acta","path":"cloud://seguimiento_pdfs/documento.pdf"},
            )
        self.assertEqual(len(objects),1)
        restored={"pickers":{}}
        self.assertEqual(ns["_load_followup_journal"](restored),1)
        docs=restored["pickers"]["PEREZ GARDUÑO ALDAIR"]["documentos"]
        self.assertEqual(len(docs),1)
        self.assertEqual(docs[0]["id"],"doc-1")
        self.assertEqual(ns["_load_followup_journal"](restored),0)

    def test_concurrent_picker_changes_keep_both_actions(self):
        ns,_,_=environment()
        remote={"PEDRO":{"acciones":[{"id":"a","motivo":"uno"}],"documentos":[]}}
        local={"PEDRO":{"acciones":[{"id":"b","motivo":"dos"}],"documentos":[]}}
        merged=ns["_merge_picker_store"](remote,local)
        self.assertEqual({x["id"] for x in merged["PEDRO"]["acciones"]},{"a","b"})

    def test_updates_use_id_not_full_record_for_dedup(self):
        ns,_,_=environment()
        rows=ns["_unique_records"]([
            {"id":"a","email_estado":"Pendiente"},
            {"id":"a","email_estado":"Enviado"},
        ])
        self.assertEqual(len(rows),1)
        self.assertEqual(rows[0]["email_estado"],"Enviado")

    def test_no_cloud_means_no_success(self):
        ns,_,calls=environment(cloud_on=False)
        with self.assertRaisesRegex(RuntimeError,"falta configurar la nube"):
            ns["save_followup_entry"]({"pickers":{}},"PEDRO","acciones",{"id":"x"})
        self.assertEqual(calls,[])

    def test_journal_pagination_over_500(self):
        ns,objects,_=environment()
        for i in range(501):
            objects[f"seguimientos_registros/r{i:04d}.json"]=b'{}'
        self.assertEqual(len(ns["_followup_event_objects"]()),501)

if __name__=="__main__":
    unittest.main()
