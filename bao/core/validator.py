"""校验引擎 — FBA 装箱单数据校验"""
from typing import Optional

from ..core.weaver import weave_fba


class Severity:
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


class ValidationResult:
    def __init__(self, rule, severity, message, field="", suggestion=""):
        self.rule = rule
        self.severity = severity
        self.message = message
        self.field = field
        self.suggestion = suggestion


class ValidationReport:
    def __init__(self):
        self.results = []

    @property
    def error_count(self):
        return sum(1 for r in self.results if r.severity == Severity.ERROR)

    @property
    def warning_count(self):
        return sum(1 for r in self.results if r.severity == Severity.WARNING)

    @property
    def info_count(self):
        return sum(1 for r in self.results if r.severity == Severity.INFO)

    @property
    def is_valid(self):
        return self.error_count == 0

    @property
    def total(self):
        return len(self.results)


def check_fba_data(data: dict) -> ValidationReport:
    """校验 FBA 货件数据

    Args:
        data: FBAParser.parse() 返回的字典

    Returns:
        ValidationReport
    """
    report = ValidationReport()
    meta = data.get("meta", {})
    items = data.get("items", [])

    # 必填字段
    if not meta.get("shipment_id"):
        report.results.append(ValidationResult(
            "必填字段", Severity.ERROR, "货件编号缺失",
            field="shipment_id"))

    if not items:
        report.results.append(ValidationResult(
            "必填字段", Severity.ERROR, "未解析到 SKU 数据"))

    # 箱号段格式检查
    for item in items:
        box_range = item.get("box_range", "")
        if box_range and "~" in box_range:
            report.results.append(ValidationResult(
                "箱号段格式", Severity.INFO,
                f"SKU {item.get('msku', '?')[:20]} 箱号段含 ~，将自动转为 -",
                field="box_range"))

    # 重量/尺寸合理性
    for item in items:
        weight = item.get("weight", 0)
        if weight <= 0:
            report.results.append(ValidationResult(
                "重量检查", Severity.WARNING,
                f"SKU {item.get('msku', '?')[:20]} 单箱重量为 0",
                field="weight"))
        length = item.get("length", 0)
        if length <= 0:
            report.results.append(ValidationResult(
                "尺寸检查", Severity.WARNING,
                f"SKU {item.get('msku', '?')[:20]} 箱子尺寸缺失",
                field="length"))

    return report
