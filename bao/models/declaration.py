"""出口报关单 Pydantic 数据模型

字段参照海关进出口货物报关单填制规范设计。
"""

from datetime import date, datetime
from enum import Enum
from typing import Optional

from pydantic import BaseModel, Field, field_validator


# ── 枚举 ───────────────────────────────────────────────────

class Incoterm(str, Enum):
    FOB = "FOB"
    CIF = "CIF"
    CFR = "CFR"
    EXW = "EXW"
    FCA = "FCA"
    CPT = "CPT"
    CIP = "CIP"
    DAP = "DAP"
    DPU = "DPU"
    DDP = "DDP"


class TradeMode(str, Enum):
    GENERAL = "一般贸易"
    PROCESSING_IMPORT = "进料加工"
    PROCESSING_SUPPLY = "来料加工"
    BONDED = "保税监管场所进出境货物"
    BONDED_LOGISTICS = "海关特殊监管区域物流货物"
    LOW_VALUE = "低值简易通关"


class TransportMode(str, Enum):
    SEA = "水路运输"
    AIR = "航空运输"
    RAIL = "铁路运输"
    ROAD = "公路运输"
    POST = "邮递运输"


class Currency(str, Enum):
    USD = "USD"
    EUR = "EUR"
    CNY = "CNY"
    JPY = "JPY"
    HKD = "HKD"
    GBP = "GBP"


class PackageType(str, Enum):
    CARTON = "纸箱"
    PALLET = "托盘"
    BAG = "袋"
    BALE = "捆"
    DRUM = "桶"
    CASE = "木箱"
    BULK = "散装"


# ── 商品项 ─────────────────────────────────────────────────

class DeclarationItem(BaseModel):
    """出口报关单-商品项"""
    seq: int = Field(ge=1, description="项号，从1开始")
    hs_code: str = Field(
        default="00000000",
        description="HS编码",
        min_length=8,
        max_length=12,
    )
    name_zh: str = Field(min_length=1, max_length=255, description="品名（中文）")
    name_en: str = Field(default="", max_length=255, description="品名（英文）")
    specs: str = Field(default="", max_length=500, description="规格型号")
    quantity: float = Field(gt=0, description="成交数量")
    unit: str = Field(description="成交单位", examples=["个", "件", "套", "千克"])
    unit_price: float = Field(ge=0, description="单价")
    total_price: float = Field(ge=0, description="总价")
    currency: Currency = Field(default=Currency.USD, description="币制")
    net_weight: float = Field(default=0.0, ge=0, description="净重（KG）")
    gross_weight: float = Field(default=0.0, ge=0, description="毛重（KG）")
    origin: str = Field(default="中国", description="原产国")
    destination_country: str = Field(default="", description="运抵国")
    final_destination: str = Field(default="", description="最终目的国")
    tax_rebate_rate: float = Field(
        default=0.0, ge=0, le=17, description="退税率（%）"
    )
    declaration_elements: str = Field(
        default="",
        max_length=2000,
        description="申报要素（品牌/型号/用途/材质等逗号分隔）",
    )
    item_no_in_contract: str = Field(default="", description="合同项号")
    hs_suggestions: list = Field(
        default_factory=list,
        description="知识库HS编码匹配建议（hs_code 缺失时自动查询）",
    )


# ── 出口报关单主体 ──────────────────────────────────────────

class ExportDeclaration(BaseModel):
    """出口报关单"""

    # ── 表头
    declaration_id: str = Field(
        default="",
        description="报关单号（系统自动生成或手工填入）",
    )
    trade_mode: TradeMode = Field(default=TradeMode.GENERAL, description="贸易方式")

    # ── 境内信息
    domestic_shipper: str = Field(min_length=1, description="境内发货人（企业名称）")
    shipper_code: str = Field(
        default="",
        description="统一社会信用代码（18位）",
        min_length=18,
        max_length=18,
    )
    shipper_contact: str = Field(default="", description="联系人")
    shipper_phone: str = Field(default="", description="联系电话")

    # ── 境外信息
    overseas_consignee: str = Field(min_length=1, description="境外收货人")
    overseas_consignee_en: str = Field(default="", description="境外收货人（英文）")
    destination_country: str = Field(min_length=1, description="运抵国")
    final_destination: str = Field(default="", description="最终目的国")
    country_of_dispatch: str = Field(default="", description="启运国")

    # ── 运输
    transport_mode: TransportMode = Field(
        default=TransportMode.SEA, description="运输方式"
    )
    vessel_name: str = Field(default="", description="船名/航次")
    voyage_no: str = Field(default="", description="航次号")
    bill_of_lading: str = Field(default="", description="提运单号")
    port_of_loading: str = Field(default="", description="装货港（境内）")
    port_of_discharge: str = Field(default="", description="指运港（境外）")
    port_of_transshipment: str = Field(default="", description="中转港")

    # ── 贸易条款
    incoterm: Incoterm = Field(default=Incoterm.FOB, description="成交方式")
    currency: Currency = Field(default=Currency.USD, description="结算币制")
    freight: float = Field(default=0.0, ge=0, description="运费")
    freight_currency: Currency = Field(default=Currency.USD)
    insurance: float = Field(default=0.0, ge=0, description="保费")
    insurance_currency: Currency = Field(default=Currency.USD)
    miscellaneous: float = Field(default=0.0, description="杂费")

    # ── 包装
    package_type: PackageType = Field(
        default=PackageType.CARTON, description="包装种类"
    )
    package_count: int = Field(ge=0, description="件数")

    # ── 合同/发票
    contract_no: str = Field(default="", description="合同号")
    invoice_no: str = Field(default="", description="发票号")
    declaration_date: date = Field(default_factory=date.today, description="申报日期")
    export_date: Optional[date] = Field(default=None, description="出口日期")
    release_date: Optional[date] = Field(default=None, description="放行日期")

    # ── 商品明细
    items: list[DeclarationItem] = Field(min_length=1, description="商品项列表")

    # ── 杂项
    remarks: str = Field(default="", max_length=1000, description="备注")
    customs_code: str = Field(default="", description="申报地海关代码")

    # ── 计算字段
    @property
    def total_packages(self) -> int:
        return self.package_count

    @property
    def total_gross_weight(self) -> float:
        return sum(item.gross_weight for item in self.items)

    @property
    def total_net_weight(self) -> float:
        return sum(item.net_weight for item in self.items)

    @property
    def total_amount(self) -> float:
        return sum(item.total_price for item in self.items)

    @property
    def fob_amount(self) -> float:
        """FOB价（出口退税以此为基数）"""
        if self.incoterm == Incoterm.FOB:
            return self.total_amount
        elif self.incoterm == Incoterm.CIF:
            return self.total_amount - self.freight - self.insurance
        elif self.incoterm in (Incoterm.CFR, Incoterm.CPT):
            return self.total_amount - self.freight
        elif self.incoterm in (Incoterm.EXW, Incoterm.FCA):
            return self.total_amount + self.freight
        return self.total_amount

    @property
    def estimated_tax_rebate(self) -> float:
        """估算退税额（基于FOB价）"""
        total = 0.0
        for item in self.items:
            rebate = item.total_price * item.tax_rebate_rate / 100
            total += rebate
        return round(total, 2)

    @property
    def item_count(self) -> int:
        return len(self.items)

    # ── 校验
    @field_validator("shipper_code")
    @classmethod
    def check_shipper_code(cls, v: str) -> str:
        if v and len(v) != 18:
            raise ValueError("统一社会信用代码必须为18位")
        return v

    def validate_weights(self) -> list[str]:
        """重量逻辑校验，返回警告信息列表"""
        warnings = []
        for item in self.items:
            if item.net_weight > item.gross_weight:
                warnings.append(
                    f"项号{item.seq} 净重({item.net_weight})不能大于毛重({item.gross_weight})"
                )
        return warnings

    def validate_incoterm_fields(self) -> list[str]:
        """成交方式相关字段校验"""
        warnings = []
        if self.incoterm == Incoterm.FOB:
            if self.freight > 0:
                warnings.append("FOB成交方式下运费应为0（由买方支付）")
            if self.insurance > 0:
                warnings.append("FOB成交方式下保费应为0（由买方支付）")
        if self.incoterm in (Incoterm.CIF, Incoterm.CFR):
            if self.freight <= 0:
                warnings.append(f"{self.incoterm.value}成交方式下运费必须填写")
        if self.incoterm == Incoterm.CIF:
            if self.insurance <= 0:
                warnings.append("CIF成交方式下保费必须填写")
        return warnings
