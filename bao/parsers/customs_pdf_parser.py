"""报关单 PDF 解析器 — 提取海关出口货物报关单关键字段并按规则映射"""
import re
import PyPDF2
from pathlib import Path
from typing import Dict, List, Optional


# 抬头映射：单号前缀 → 公司名
TITLE_MAP = {
    "QX": "沁香",
    "MS": "蔓莎",
    "LJ": "六觉",
}


class CustomsPDFParser:
    """解析海关出口货物报关单 PDF，提取关键字段并按业务规则映射"""

    # ── 正则模式 ──
    RE_ENTRY_NO = re.compile(r'\*(\d{18,})\*')               # 条形码号码 / 海关编号
    RE_DECL_DATE = re.compile(r'(\d{8})\s*申报日期')          # 申报日期: 20260403
    RE_TOTAL_PRICE = re.compile(r'总价.*?([\d,]+\.?\d*)')    # 总价: 5153.40
    RE_CURRENCY = re.compile(r'(美元|人民币|欧元|日元|英镑|港币)')  # 币制
    RE_DATE_8DIGIT = re.compile(r'(20\d{2})(\d{2})(\d{2})')  # 8位日期提取
    RE_PRICE_LINE = re.compile(r'^\s*([\d,]+\.?\d*)\s*$', re.MULTILINE)

    HEADER_FIELDS = [
        "单号", "下单日期", "报单订单单号", "报关单号", "报关金额",
        "币种", "兑成RMB", "报关单是否收到", "平台", "上传政府平台",
        "申报日期", "提单报关单上传一达通平台", "物流", "报关抬头",
    ]

    def parse(self, file_path: str) -> Dict:
        """解析单个 PDF 报关单，返回映射后的字段字典"""
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"PDF 文件不存在: {file_path}")

        text = self._extract_text(str(path))
        return self._parse_text(text, path.name)

    def parse_batch(self, file_paths: List[str]) -> List[Dict]:
        """批量解析多个 PDF，返回结果列表"""
        results = []
        for fp in file_paths:
            try:
                results.append(self.parse(fp))
            except Exception as e:
                results.append({"error": str(e), "file": Path(fp).name})
        return results

    # ────────── 内部方法 ──────────

    def _extract_text(self, file_path: str) -> str:
        """使用 PyPDF2 提取 PDF 全文"""
        reader = PyPDF2.PdfReader(file_path)
        lines = []
        for page in reader.pages:
            t = page.extract_text()
            if t:
                lines.append(t)
        return "\n".join(lines)

    def _extract_contract_no(self, text: str) -> str:
        """从文本中提取合同协议号"""
        m = re.search(r'([A-Z]{2,4}\d{8,}[\-\d]*)\s*合同协议号', text)
        if m:
            return m.group(1).strip()
        return ""

    def _extract_entry_no(self, text: str) -> str:
        """提取海关编号 / 预录入编号"""
        m = self.RE_ENTRY_NO.search(text)
        if m:
            return m.group(1)
        m = re.search(r'预录入编号[：:]\s*(\d{18,})', text)
        if m:
            return m.group(1)
        m = re.search(r'海关编号[：:]\s*(\d{18,})', text)
        if m:
            return m.group(1)
        return ""

    def _extract_decl_date(self, text: str) -> str:
        """提取申报日期"""
        m = self.RE_DECL_DATE.search(text)
        if m:
            return self._format_date(m.group(1))
        return ""

    def _extract_total_price(self, text: str) -> Optional[float]:
        """提取报关总价 — 在币种关键字附近定位"""
        lines = text.split('\n')
        for i, line in enumerate(lines):
            stripped = line.strip()
            m_cur = re.search(r'(美元|人民币|欧元|日元|英镑|港币)', stripped)
            if m_cur:
                before = stripped[:m_cur.start()]
                m_price = re.search(r'([\d,]+\.?\d*)\s*$', before)
                if m_price:
                    return self._to_float(m_price.group(1))
                if i > 0:
                    prev = lines[i - 1].strip()
                    m_price = self.RE_PRICE_LINE.match(prev)
                    if m_price:
                        return self._to_float(m_price.group(1))
        m = self.RE_TOTAL_PRICE.search(text)
        if m:
            return self._to_float(m.group(1))
        return None

    def _extract_currency(self, text: str) -> str:
        """提取币制"""
        m = self.RE_CURRENCY.search(text)
        return m.group(1) if m else ""

    def _parse_order_date(self, contract_no: str) -> str:
        """从合同协议号中提取 20xx/xx/xx 格式日期"""
        m = self.RE_DATE_8DIGIT.search(contract_no)
        if m:
            return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
        return ""

    def _parse_title(self, contract_no: str) -> str:
        """单号前缀 → 报关抬头"""
        if not contract_no:
            return ""
        prefix = contract_no[:2].upper()
        return TITLE_MAP.get(prefix, "")

    @staticmethod
    def _format_date(raw: str) -> str:
        raw = raw.strip()
        if len(raw) == 8 and raw.isdigit():
            return f"{raw[:4]}/{raw[4:6]}/{raw[6:8]}"
        return raw

    @staticmethod
    def _to_float(s: str) -> Optional[float]:
        s = s.replace(",", "").strip()
        try:
            return float(s)
        except (ValueError, TypeError):
            return None

    def _parse_text(self, text: str, filename: str = "") -> Dict:
        """从 PDF 全文解析所有字段并映射到汇总表结构"""
        contract_no = self._extract_contract_no(text)
        result = {
            "单号": contract_no,
            "下单日期": self._parse_order_date(contract_no),
            "报单订单单号": "",
            "报关单号": self._extract_entry_no(text),
            "报关金额": self._extract_total_price(text) or "",
            "币种": self._extract_currency(text),
            "兑成RMB": "",
            "报关单是否收到": "",
            "平台": "综保平台",
            "上传政府平台": "",
            "申报日期": self._extract_decl_date(text),
            "提单报关单上传一达通平台": "",
            "物流": "",
            "报关抬头": self._parse_title(contract_no),
            "_source": filename,
        }
        return result
