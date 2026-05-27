from bao.parsers.excel_parser import ExcelParser
from bao.core.weaver import weave
from bao.core.exporter import DeclarationExporter

p = ExcelParser()
inv = p.parse_invoice('data/示例发票.xlsx')
pkg = p.parse_packing('data/示例装箱单.xlsx')
decl = weave(inv, pkg, domestic_shipper='苏州恒通纺织进出口有限公司')
print('运抵国:', decl.destination_country)
print('总金额:', decl.total_amount)
print('件数:', decl.package_count)
print('项数:', decl.item_count)
print('HS:', [i.hs_code for i in decl.items])
print('FOB:', decl.fob_amount)

exp = DeclarationExporter()
exp.export(decl, 'data/报关单_示例.xlsx')
print('✅ 已导出')
