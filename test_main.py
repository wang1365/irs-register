import configparser
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


if __name__ == "__main__":
    unittest.main()
