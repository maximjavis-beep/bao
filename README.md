# bao — 出口报关文件编织助手

从 FBA 货件 Excel 一键生成清关装箱单，同时支持调整明细自动填充发货计划。双面板 Web 界面 + CLI + Vercel 线上部署。

🌐 **线上版**：https://bao-sepia-delta.vercel.app/

---

## 快速开始

### 本地面板

```bash
cd customs/bao
source venv/bin/activate
bash restart.sh
```

浏览器打开 **http://127.0.0.1:7777**（或 `python start_web.py 9999` 指定端口）

### Vercel 线上部署

项目通过 GitHub + Vercel 自动部署，每次 push 自动上线。

## 使用流程

### Web 面板（上下分屏）

**上屏 — 调整明细 → 发货计划**

1. 上传调整明细 Excel（含识别码、店铺、SKU、调整FNSKU、调整量等）
2. 点击「🔀 生成发货计划」
3. 下载填充好的发货计划 Excel（新增行黄色标注）

**下屏 — FBA 装箱单生成**

1. 上传 FBA 货件 Excel（packing sheet）— 支持多选拖拽
2. 自动解析并预览 SKU 明细和货件信息
3. （可选）填写 HS 编码覆盖
4. 点击「🧵 单件生成」或「📦 批量生成」
5. 单件直接下载，批量打包 ZIP

### CLI

```bash
bao build from-fba -i FBA货件.xlsx -o 装箱单.xlsx
bao build preview -i FBA货件.xlsx
```

## 自动规则

程序生成装箱单时自动应用以下规则：

| 规则 | 说明 |
|------|------|
| 箱号段标准化 | `~` → `-`，单整数如 `16` 自动补全为 `16-16` |
| ASIN 自动带入 | packing sheet 的 ASIN 列自动填入 L 列（欧洲清关必填） |
| 模板图片清除 | 自动删除模板示例图片，避免残留到输出文件 |
| T 列不自动填入 | 图片列保持空白 |

## 调整明细 → 发货计划 映射规则

上传调整明细后，自动按以下规则填充发货计划模板：

| 发货计划列 | 数据来源 | 规则 |
|-----------|------|------|
| C 识别码 | 调整明细 识别码 | 同名同义 |
| D 店铺-国家 | 调整明细 店铺 | 异名同义 |
| E 国家 | D 列 `-` 后代码 | 推导：MSCandle-JP → JP |
| F SKU | 调整明细 SKU | 同名同义 |
| G FNSKU | 调整明细 调整FNSKU | 调整后目标值 |
| J 计划发货量 | 调整明细 调整量 | 异名同义 |
| K 库存数 | 调整明细 调整量 | 异名同义 |
| P 店铺 | 调整明细 调整店铺 | 调整后归属 |

> 无法对应的列（发货批次号、MSKU、单箱数量等）留空，新增行整行黄色标注。

## 安全机制（线上版）

| 措施 | 说明 |
|------|------|
| 上传即删 | 原始货件文件解析后立即从 `/tmp` 删除 |
| 一次性令牌下载 | 装箱单通过 `/api/download/{token}` 下载，用完销毁 |
| 无公开文件路径 | 不暴露 `/downloads/` 目录，无跨请求文件残留 |
| 批量 zip 内存打包 | zip 在内存中生成 base64，不落盘 |

## 列名兼容性

解析器自动识别 packing sheet 列名：MSKU、ASIN、商品申报量、箱子长/宽/高/重量、货件箱子编号等。

## 项目结构

```
bao/
├── api/index.py              # FastAPI Vercel 入口
├── bao/
│   ├── cli.py                # CLI 入口
│   ├── commands/             # build / check / serve / archive
│   ├── core/                 # weaver / exporter
│   ├── parsers/              # FBA Excel 解析器（纯 openpyxl）
│   └── web/                  # server.py + index.html + index_vercel.html
├── templates/                # 装箱单模板 + 发货计划模板
├── pyproject.toml            # 依赖配置（Vercel 安装来源）
└── restart.sh                # 本地面板启动脚本
```

## 依赖

fastapi / python-multipart / openpyxl / pydantic / typer / rich

## 故障排除

| 现象 | 解决 |
|------|------|
| 本地面板打不开 | 用 `127.0.0.1:7777`（默认端口） |
| 端口占用 | `lsof -i :7777 -t \| xargs kill` |
| 生成内容为空 | 确认上传的是 FBA packing sheet（含 MSKU 列） |
| Vercel 报错 | 查看 Vercel Build Log 或 Functions 日志 |
