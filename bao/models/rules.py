"""校验规则定义"""

from dataclasses import dataclass, field
from enum import Enum


class Severity(str, Enum):
    ERROR = "error"      # 阻断性错误，必须修正
    WARNING = "warning"  # 警告，建议修正
    INFO = "info"        # 提示信息


@dataclass
class ValidationResult:
    """单条校验结果"""
    rule: str
    severity: Severity
    message: str
    field: str = ""
    suggestion: str = ""  # 知识库修复建议


@dataclass
class ValidationReport:
    """校验报告"""
    results: list[ValidationResult] = field(default_factory=list)

    @property
    def error_count(self) -> int:
        return sum(1 for r in self.results if r.severity == Severity.ERROR)

    @property
    def warning_count(self) -> int:
        return sum(1 for r in self.results if r.severity == Severity.WARNING)

    @property
    def info_count(self) -> int:
        return sum(1 for r in self.results if r.severity == Severity.INFO)

    @property
    def is_valid(self) -> bool:
        """无阻断性错误即为通过"""
        return self.error_count == 0

    @property
    def total(self) -> int:
        return len(self.results)
