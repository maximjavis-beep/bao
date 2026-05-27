"""装箱单（Packing List）数据模型"""

from datetime import date as date_type

from pydantic import BaseModel, Field


class PackageItem(BaseModel):
    """装箱单-每箱/每托明细"""
    seq: int = Field(ge=1, description="箱号")
    description: str = Field(default="", description="品名概述")
    quantity: float = Field(gt=0, description="本箱数量")
    unit: str = Field(default="PCS", description="单位")
    net_weight: float = Field(gt=0, description="净重(KG)")
    gross_weight: float = Field(gt=0, description="毛重(KG)")
    dimensions: str = Field(default="", description="尺寸(L×W×H cm)")


class PackingList(BaseModel):
    """装箱单"""
    packing_no: str = Field(default="", description="装箱单号（通常等于发票号）")
    packing_date: date_type = Field(description="装箱日期")

    # 关联
    invoice_no: str = Field(default="", description="关联发票号")
    contract_no: str = Field(default="", description="关联合同号")

    # 发货方
    shipper: str = Field(default="", description="发货人")
    consignee: str = Field(default="", description="收货人")

    # 总体
    package_type: str = Field(default="纸箱", description="包装种类")
    marks: str = Field(default="", description="唛头")
    bill_of_lading: str = Field(default="", description="提单号")

    # 每箱明细
    packages: list[PackageItem] = Field(min_length=1, description="装箱明细")

    @property
    def total_packages(self) -> int:
        return len(self.packages)

    @property
    def total_quantity(self) -> float:
        return sum(p.quantity for p in self.packages)

    @property
    def total_net_weight(self) -> float:
        return sum(p.net_weight for p in self.packages)

    @property
    def total_gross_weight(self) -> float:
        return sum(p.gross_weight for p in self.packages)
