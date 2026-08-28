from decimal import Decimal
from pathlib import Path

import yaml

GOLDENS = Path(__file__).parent / "golden"


def load(name: str) -> dict:
    return yaml.safe_load((GOLDENS / name).read_text(encoding="utf-8"))


def test_four_owner_approved_goldens_exist():
    files = sorted(p.name for p in GOLDENS.glob("*.yaml"))
    assert files == [
        "credit_79_2_AT.yaml",
        "credit_79_3_AT.yaml",
        "debit_79_2_AT.yaml",
        "debit_79_3_AT.yaml",
    ]
    for name in files:
        case = load(name)
        assert case["status"] == "APPROVED"
        assert case["input"]["source_account"] in {"79.2", "79.3"}
        assert len(case["expected"]) == 2
        assert all(Decimal(row["amount"]) == Decimal("84272.40") for row in case["expected"])


def test_credit_is_exact_mirror_of_debit_for_each_source_account():
    for account_token in ("79_2", "79_3"):
        debit = load(f"debit_{account_token}_AT.yaml")
        credit = load(f"credit_{account_token}_AT.yaml")
        assert debit["input"]["organization"] == credit["input"]["organization"] == "АТ"
        assert debit["input"]["department"] == credit["input"]["department"]
        assert debit["input"]["supplier_rvp"] == credit["input"]["supplier_rvp"]

        debit_source, debit_gk = debit["expected"]
        credit_source, credit_gk = credit["expected"]

        assert credit_source["debit_account"] == debit_source["credit_account"]
        assert credit_source["credit_account"] == debit_source["debit_account"]
        assert credit_source["debit_supplier_rvp"] == debit_source["credit_supplier_rvp"]
        assert credit_source["credit_supplier_rvp"] == debit_source["debit_supplier_rvp"]
        assert credit_source["debit_department"] == debit_source["credit_department"]
        assert credit_source["credit_department"] == debit_source["debit_department"]

        assert credit_gk["debit_account"] == debit_gk["credit_account"] == "79.1"
        assert credit_gk["credit_account"] == debit_gk["debit_account"] == "79.1"
        assert credit_gk["debit_supplier_rvp"] == debit_gk["credit_supplier_rvp"]
        assert credit_gk["credit_supplier_rvp"] == debit_gk["debit_supplier_rvp"]
        assert credit_gk["debit_department"] == debit_gk["credit_department"]
        assert credit_gk["credit_department"] == debit_gk["debit_department"]


def test_79_2_and_79_3_have_identical_rules_except_source_account():
    for side in ("debit", "credit"):
        a = load(f"{side}_79_2_AT.yaml")
        b = load(f"{side}_79_3_AT.yaml")
        assert a["input"] | {"source_account": "79.x"} == b["input"] | {"source_account": "79.x"}

        def normalize(rows: list[dict]) -> list[dict]:
            normalized = []
            for row in rows:
                item = dict(row)
                if item["debit_account"] in {"79.2", "79.3"}:
                    item["debit_account"] = "79.x"
                if item["credit_account"] in {"79.2", "79.3"}:
                    item["credit_account"] = "79.x"
                normalized.append(item)
            return normalized

        assert normalize(a["expected"]) == normalize(b["expected"])
