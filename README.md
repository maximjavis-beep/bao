# bao — 出口报关文件编织助手

从 FBA 货件 Excel 一键生成清关装箱单，支持调整明细自动填充发货计划、货件追踪码匹配、自定义模板上传、DISPIMG 嵌入图片。双面板 Web 界面 + CLI + Vercel 线上部署。

🌐 **线上版**：https://bao-sepia-delta.vercel.app/

---

## 快速开始

### 本地面板

```bash
cd customs/bao
source venv/bin/activate
bash restart.sh
```

浏览器打开 **http://127.0.0.1:8888**（或 `python start_web.py 9999` 指定端口）

### Vercel 线上部署

项目通过 GitHub + Vercel 自动部署，每次 push 自动上线。

---

## 使用流程

### Web 面板（上下分屏）

#### 上屏 — 调整明细 → 发货计划

1. 上传「调整明细」Excel + 「发货计划」Excel
2. 点击「🔀 生成调整后发货计划」
3. 按四步规则（店铺筛选 → 识别码+FNSKU 匹配 → 数量判断 → 清理黄色行）自动处理
4. 下载结果（匹配组数 / 删除行数统计）

#### 下屏 — FBA 装箱单生成

1. 上传 FBA 货件 Excel（packing sheet）— 支持多选拖拽
2. （可选）上传自定义模板 — 自动适配表头，不限品类
3. （可选）上传货件追踪码 / 出运数据 — 自动匹配 Reference ID，并填充仓库代码、渠道、总件数
4. 自动解析并预览 SKU 明细和货件信息
5. 点击「🧵 单件生成」或「📦 批量生成」
6. 单件直接下载，批量打包 ZIP

### CLI

```bash
bao build from-fba -i FBA货件.xlsx -o 装箱单.xlsx
bao build preview -i FBA货件.xlsx
```

---

## 自动规则

### 模板颜色规则

程序根据上传模板的表头填充色自动分类处理：

| 表头填充色 | 行为 |
|-----------|------|
| 🟡 黄色 (FFFF00) | 自动匹配 — 从 HS 知识库或 FBA 数据填入，匹配不到则留空 |
| ⬜ 灰色 (theme tint) | 模板复制 — 部分字段（是否带电/型号/单位/每套个数）从模板示例行复制 |

黄色列表头关键词自动识别（Shipment ID、箱号段、SKU、英文品名、中文品名、材质、用途、品牌、海关编码、每箱数量、单箱重量/长/宽/高），不依赖固定列位置。

### 品类匹配

通过 `data/hs_code.xlsx` 查表获取各品类的英文品名、中文品名、材质、用途、品牌等信息。支持上传自定义模板，表头自动推断（扫描前 30 行，匹配 25+ 个关键词字段），不限品类。

### Commercial Invoice 元数据自动填充

上传出运数据文件时，根据 FBA 编号自动匹配并填充模板上半区元数据：

| 模板字段 | 数据来源 | 规则 |
|---------|------|------|
| 参考号 | Shipment ID | 直接使用 FBA 货件编号 |
| 目的仓库代码 | 出运数据 仓库代码 | 按 FBA 编号匹配 |
| 渠道 | 出运数据 渠道 | 按 FBA 编号匹配 |
| 总件数 | 出运数据 箱数 | 按 FBA 编号匹配 |
| Reference ID | 出运数据 货件追踪码 | 按 FBA 编号匹配 |

出运数据文件按表头关键词自动定位列（FBA编号 / 货件追踪码 / 仓库代码 / 渠道 / 箱数），不依赖固定列位置。

### 装箱单生成规则

| 规则 | 说明 |
|------|------|
| 箱号段标准化 | `~` / `～` → `-`，单整数如 `16` 自动补全为 `16-16` |
| ASIN 自动带入 | packing sheet 的 ASIN 列自动填入 |
| 海关编码 | 优先匹配 HS 知识库，回退到输入文件的"进口海关编码" |
| 灰色字段 | 是否带电/型号/单位/每套个数 从模板第一行复制 |
| Sheet 自动检测 | 优先匹配「下单模板」，回退到首个 sheet |
| 自由列位置 | 不依赖固定 ABC 列，按表头关键词自动定位 |

### 图片处理

| 模板图片类型 | 检测方式 | 输出方式 |
|------------|----------|---------|
| openpyxl 嵌入图 (PNG) | `ws._images` 非空 | `TwoCellAnchor` 填满单元格 + `editAs="oneCell"` 自适应 |
| WPS DISPIMG 公式 | `xl/cellimages.xml` 存在 | `=_xlfn.DISPIMG("ID",1)` 公式 + cellimages 四组件注入 |

> **DISPIMG 注入四组件**：
> 1. `[Content_Types].xml` — 注册 cellimages 类型
> 2. `xl/_rels/workbook.xml.rels` — WPS cellImage 关系
> 3. `xl/cellimages.xml` + `xl/_rels/cellimages.xml.rels`
> 4. 单元格 `=_xlfn.DISPIMG("ID",1)` 公式
>
> 缺少任一项 → `#NAME?` 或空白。

---

## 调整明细 → 发货计划 映射规则

四步处理流程（适用调整明细与发货计划表格）：

| 步骤 | 规则 |
|------|------|
| 13 | 筛选同店铺 — 调整店铺 = 计划表店铺 |
| 14 | 识别码 + FNSKU 匹配 — 在已有计划行中匹配 |
| 15 | 数量判断 — 1:1 / 多:1 / 差异（仅标记，不覆盖） |
| 16 | 清理 — 删除已处理的黄色新增行，补全库存箱数 |

| 发货计划列 | 数据来源 | 规则 |
|-----------|------|------|
| C 识别码 | 调整明细 识别码 | 同名同义 |
| E 国家 | D 列 `-` 后代码 | 推导：MSCandle-JP → JP |
| F SKU | 调整明细 SKU | 同名同义 |
| G FNSKU | 调整明细 调整FNSKU | 调整后目标值 |
| J 计划发货量 | 调整明细 调整量 | 异名同义 |
| K 库存数 | 调整明细 调整量 | 异名同义 |
| Q 调拨单号/调拨量 | 调整明细 调整单号 | 多个用逗号拼接 |

> 无法对应的列留空，新增行整行黄色标注。

---

## 项目结构

```
bao/
├── api/index.py              # FastAPI Vercel 入口
├── bao/
│   ├── core/
│   │   ├── exporter.py       # 导出（颜色规则/DISPIMG/追踪码/元数据填充）
│   │   └── weaver.py         # 编织（HS 查表/箱号标准化/品类匹配）
│   ├── parsers/
│   │   └── fba_parser.py     # FBA packing sheet 解析器
│   └── web/
│       ├── server.py         # 本地 HTTP 服务
│       ├── index.html        # 本地 Web 面板
│       └── index_vercel.html # Vercel 在线面板
├── data/
│   └── hs_code.xlsx          # HS 编码知识库
├── templates/                # 模板文件
├── pyproject.toml            # 依赖配置
└── README.md
```

## 依赖

fastapi / python-multipart / openpyxl / pydantic / typer / rich / Pillow

## 故障排除

| 现象 | 解决 |
|------|------|
| 本地面板打不开 | 用 `127.0.0.1:8888`（默认端口） |
| 端口占用 | `lsof -i :8888 -t \| xargs kill` |
| Vercel 部署失败 | 查看 Vercel Build Log |
| 输出图片显示 `#NAME?` | 模板为 DISPIMG 格式但 cellimages 未正确注入 |
| 自定义模板列对不上 | 检查模板表头关键词是否在 PATTERNS 中 |
| 追踪码未匹配 | 确认出运数据文件 FBA编号 列与 Shipment ID 一致 |
| 渠道填错 | 确认模板元数据标签名称（如"渠道"）与 PATTERNS 精确匹配，注意"渠道"≠"渠道能力" |
