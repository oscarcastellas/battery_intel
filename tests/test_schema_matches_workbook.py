"""The SQL schema and the workbook must describe the same tables, columns and allowed values."""

import re
import unittest

from helpers import ROOT, bt

SCHEMA = (ROOT / "sql" / "0001_schema.sql").read_text()


def table_ddl(table):
    m = re.search(rf"CREATE TABLE {table} \((.*?)\n\);", SCHEMA, re.S)
    return m.group(1) if m else None


class TestSchemaMatchesWorkbook(unittest.TestCase):
    def test_every_tab_has_a_table(self):
        for table in bt.TABLES:
            self.assertIsNotNone(table_ddl(table), f"no CREATE TABLE for {table}")

    def test_every_input_column_exists(self):
        for table, columns in bt.TABLES.items():
            ddl = table_ddl(table)
            for name, _, kind, _ in columns:
                if not kind.startswith("auto:"):
                    self.assertRegex(ddl, rf"(?m)^\s+{name}\s", f"{table}.{name} missing from SQL")

    def test_dropdown_values_match_check_constraints(self):
        for table, columns in bt.TABLES.items():
            ddl = table_ddl(table)
            for name, _, kind, _ in columns:
                if kind.startswith("list:"):
                    for value in bt.LISTS[kind[5:]]:
                        self.assertIn(f"'{value}'", ddl, f"{table}.{name}: '{value}' not allowed by SQL CHECK")

    def test_id_prefixes_match(self):
        for table, prefix in bt.ID_PREFIX.items():
            self.assertIn(f"LIKE '{prefix}%'", table_ddl(table), f"{table}: ID prefix check missing")


if __name__ == "__main__":
    unittest.main()
