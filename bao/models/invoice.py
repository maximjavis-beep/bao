"""商业发票（Invoice）数据模型"""

from datetime import date
from typing import Optional

from pydantic import BaseModel, Field

from .declaration import Currency, Incoterm


class InvoiceItem(BaseModel):
    """发票行项目"""
    seq: int = Field(ge=1, description="行号")
    name_zh: str = Field(min_length=1, description="品名（中文）")
    name_en: str = Field(default="", description="品名（英文）")
    specs: str = Field(default="", description="规格型号")
    hs_code: str = Field(default="", max_length=10, description="HS编码")
    quantity: float = Field(gt=0, description="数量")
    unit: str = Field(default="PCS", description="单位")
    unit_price: float = Field(ge=0, description="单价")
    total_price: float = Field(ge=0, description="总价")
    net_weight: float = Field(default=0.0, ge=0, description="净重(KG)")
    gross_weight: float = Field(default=0.0, ge=0, description="毛重(KG)")
    origin: str = Field(default="中国", description="原产国")


class CommercialInvoice(BaseModel):
    """商业发票"""
    invoice_no: str = Field(min_length=1, description="发票号")
    invoice_date: date = Field(description="发票日期")

    # 卖方（出口商）
    seller: str = Field(min_length=1, description="卖方名称")
    seller_address: str = Field(default="", description="卖方地址")
    seller_contact: str = Field(default="", description="卖方联系人")
    seller_tel: str = Field(default="", description="卖方电话")

    # 买方（进口商）
    buyer: str = Field(min_length=1, description="买方名称")
    buyer_address: str = Field(default="", description="买方地址")
    buyer_en: str = Field(default="", description="买方名称（英文）")

    # 运输
    vessel_name: str = Field(default="", description="船名/航班号")
    voyage_no: str = Field(default="", description="航次")
    port_of_loading: str = Field(default="", description="起运港")
    port_of_discharge: str = Field(default="", description="目的港")
    destination_country: str = Field(default="", description="目的国")

    # 贸易
    incoterm: Incoterm = Field(default=Incoterm.FOB, description="贸易术语")
    currency: Currency = Field(default=Currency.USD, description="结算币制")

    # 合同
    contract_no: str = Field(default="", description="合同号")
    payment_terms: str = Field(default="T/T", description="付款方式")

    # 包装汇总
    package_count: int = Field(default=0, ge=0, description="总件数")
    package_type: str = Field(default="纸箱", description="包装种类")

    # 行项目
    items: list[InvoiceItem] = Field(min_length=1, description="发票明细")

    # 计算
    @property
    def total_amount(self) -> float:
        return sum(item.total_price for item in self.items)

    @property
    def total_net_weight(self) -> float:
        return sum(item.net_weight for item in self.items)

    @property
    def total_gross_weight(self) -> float:
        return sum(item.gross_weight for item in self.items)
