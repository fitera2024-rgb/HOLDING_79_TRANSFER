"""Reference values used by the 79.x transfer engine.

Business-domain code receives these values through ``TransferConfig`` so the
organization and department names are not duplicated in posting construction.
"""

from __future__ import annotations

from dataclasses import dataclass

from .models import RULES_VERSION

MANAGER_ORGANIZATION = "ГК"
MANAGER_FINANCIAL_DEPARTMENT = "Б_ГК Финансовый отдел"


@dataclass(frozen=True)
class TransferConfig:
    """Reference/configuration values for one transfer-engine run."""

    manager_organization: str = MANAGER_ORGANIZATION
    manager_financial_department: str = MANAGER_FINANCIAL_DEPARTMENT
    rules_version: str = RULES_VERSION

    def __post_init__(self) -> None:
        for name in (
            "manager_organization",
            "manager_financial_department",
            "rules_version",
        ):
            value = getattr(self, name)
            if not isinstance(value, str):
                raise TypeError(f"{name} must be a string")
            value = value.strip()
            if not value:
                raise ValueError(f"{name} must not be empty")
            object.__setattr__(self, name, value)


TransferEngineConfig = TransferConfig
