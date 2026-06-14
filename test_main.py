import configparser
import base64
from datetime import date
import tempfile
import unittest
from pathlib import Path

import main


class ConfigTests(unittest.TestCase):
    def test_load_config_values_reads_default_section(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "config.ini"
            path.write_text(
                "\n".join(
                    [
                        "[DEFAULT]",
                        "ProductKey = PK-001",
                        "RegistrationKeySuffix = SUFFIX",
                        "RegistrationKeyStart = 00042",
                        "RegistrationKeyCount = 7",
                    ]
                ),
                encoding="utf-8",
            )

            values = main.load_config_values(path)

        self.assertEqual(values["ProductKey"], "PK-001")
        self.assertEqual(values["RegistrationKeySuffix"], "SUFFIX")
        self.assertEqual(values["RegistrationKeyStart"], "00042")
        self.assertEqual(values["RegistrationKeyCount"], "7")

    def test_load_config_values_accepts_lowercase_keys(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "config.ini"
            path.write_text(
                "\n".join(
                    [
                        "[DEFAULT]",
                        "productkey = PK-LOWER",
                        "registrationkeysuffix = LOWER-SUFFIX",
                        "registrationkeystart = 00009",
                        "registrationkeycount = 22",
                    ]
                ),
                encoding="utf-8",
            )

            values = main.load_config_values(path)

        self.assertEqual(values["ProductKey"], "PK-LOWER")
        self.assertEqual(values["RegistrationKeySuffix"], "LOWER-SUFFIX")
        self.assertEqual(values["RegistrationKeyStart"], "00009")
        self.assertEqual(values["RegistrationKeyCount"], "22")

    def test_save_config_values_writes_default_section(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "config.ini"
            main.save_config_values(
                {
                    "ProductKey": "PK-002",
                    "RegistrationKeySuffix": "TAIL",
                    "RegistrationKeyStart": "00003",
                    "RegistrationKeyCount": "12",
                },
                path,
            )

            parser = configparser.ConfigParser()
            parser.read(path, encoding="utf-8")
            text = path.read_text(encoding="utf-8")

        self.assertEqual(parser["DEFAULT"]["ProductKey"], "PK-002")
        self.assertEqual(parser["DEFAULT"]["RegistrationKeySuffix"], "TAIL")
        self.assertEqual(parser["DEFAULT"]["RegistrationKeyStart"], "00003")
        self.assertEqual(parser["DEFAULT"]["RegistrationKeyCount"], "12")
        self.assertIn("ProductKey = PK-002", text)
        self.assertIn("RegistrationKeySuffix = TAIL", text)

    def test_get_system_info_lines_contains_author_and_email(self):
        lines = dict(main.get_system_info_lines())

        self.assertEqual(lines["Author"], main.AUTHOR_NAME)
        self.assertEqual(lines["Email"], main.AUTHOR_EMAIL)
        self.assertNotIn("Expiration Date", lines)

    def test_validate_expiration_allows_expiration_date(self):
        errors = main.validate_expiration(date.today(), date.today())

        self.assertEqual(errors, [])

    def test_validate_expiration_rejects_date_after_expiration(self):
        errors = main.validate_expiration(date(2026, 12, 31), date(2027, 1, 1))

        self.assertEqual(errors, [main.get_expiration_log_message(date(2026, 12, 31))])
        self.assertNotIn("expired", errors[0].lower())
        self.assertNotIn("2026-12-31", errors[0])

    def test_get_expiration_log_message_encodes_expiration_date(self):
        encoded = base64.b64encode(b"2026-12-31").decode("ascii")

        self.assertEqual(main.get_expiration_log_message(date(2026, 12, 31)), f">>> {encoded}")

    def test_validate_expiration_allows_date_but_log_message_is_still_available(self):
        errors = main.validate_expiration(date(2026, 12, 31), date(2026, 1, 1))

        self.assertEqual(errors, [])
        self.assertTrue(main.get_expiration_log_message(date(2026, 12, 31)).startswith(">>> "))


if __name__ == "__main__":
    unittest.main()
