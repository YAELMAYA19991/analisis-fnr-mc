"""Pruebas de seguridad para la miniaplicación de validación móvil."""
import unittest
import mobile_audit_backend as m


class MobileAuditTests(unittest.TestCase):
    def test_verified_identity(self):
        user={"is_logged_in":True,"email_verified":True,"email":"picker@example.com"}
        self.assertEqual(m.verified_identity(user),"picker@example.com")
        self.assertEqual(m.verified_identity({**user,"email_verified":False}),"")
        self.assertEqual(m.verified_identity({**user,"is_logged_in":False}),"")

    def test_supervisor_pin_validates_without_email(self):
        self.assertTrue(m.supervisor_pin_valid("pin-de-prueba", "pin-de-prueba"))
        self.assertFalse(m.supervisor_pin_valid("incorrecto", "pin-de-prueba"))
        self.assertFalse(m.supervisor_pin_valid("123", "123"))

    def test_domain_must_be_configured(self):
        self.assertFalse(m.domain_allowed("picker@example.com",""))
        self.assertTrue(m.domain_allowed("picker@example.com","example.com"))
        self.assertFalse(m.domain_allowed("picker@example.com","otro.com"))

    def test_autoverification_requires_owner(self):
        self.assertTrue(m.validated_assignment("autoverificacion","picker@example.com",["picker@example.com"]))
        with self.assertRaises(ValueError):
            m.validated_assignment("autoverificacion","other@example.com",["picker@example.com"])

    def test_cross_audit_forbids_owner(self):
        self.assertTrue(m.validated_assignment("auditoria_cruzada","other@example.com",["picker@example.com"]))
        with self.assertRaises(ValueError):
            m.validated_assignment("auditoria_cruzada","picker@example.com",["picker@example.com"])

    def test_ambiguous_email_fails_closed(self):
        import pandas as pd
        frame=pd.DataFrame({"Picker relacionado":["Alex"]})
        people={
            "one@example.com":{"name":"Alex","names":{"alex"}},
            "two@example.com":{"name":"Alex","names":{"alex"}},
        }
        self.assertEqual(m.owners_for_lines(frame,people),([],["Alex"]))


if __name__=="__main__":
    unittest.main()
