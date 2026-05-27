# bao — 出口报关文件编织助手

从发票/装箱单 Excel 一键生成出口报关单，支持 Web 面板、CLI 校验和归档检索。

---

## 安装

```bash
cd customs/bao
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
pip install -e .
```

## 启动 Web 面板

```bash
bash restart.sh            # 杀旧进程 + 清缓存 + 后台启动
```

浏览器打开 **http://127.0.0.1:8888**

停止服务：`pkill -f start_web.py`

## Web 面板使用流程

1. 左侧拖入发票 Excel，自动解析并预览
2. （可选）拖入装箱单 Excel
3. 填写手动字段（发货人、贸易条款、运输信息等）
4. 点击「🧵 编织生成报关单」
5. 中间展示商品明细，右侧显示摘要 + 校验报告 + 下载按钮
6. 右侧面板切换至「📚 知识库」Tab，搜索法规、HS 归类、退单排查等参考资料
7. 点击「📚 归档历史」查看已归档记录

## 面板可编辑字段

| 区域 | 字段 |
|------|------|
| 🏢 发货人 | 名称、统一社会信用代码、报关单号 |
| 🚢 贸易条款 | 成交方式、贸易方式、币制、运费、保费 |
| ✈️ 运输信息 | 运输方式、船名/航次、装货港、指运港、运抵国 |
| 🗂️ 列名映射 | JSON 自定义列名映射（上传后自动提示） |

## CLI 命令

```bash
bao build preview -i 发票.xlsx                  # 预览
bao build from-files -i 发票.xlsx -p 装箱单.xlsx -o 报关单.xlsx -s "公司名"
bao check validate -i 发票.xlsx -p 装箱单.xlsx   # 校验（彩色报告）
bao archive list                                 # 归档列表
bao archive search "关键词"                       # 模糊搜索
bao archive stats                                # 月度统计
bao serve start                                  # Web 面板
bao version
```

## 列名兼容性

默认中英文自动识别：品名/Description、HS编码/HS Code、数量/Quantity 等。
上传后系统检测未识别列名并在映射框填入候选，补充后即可：

```json
{"name_zh": ["产品名称"], "hs_code": ["海关编码"], "quantity": ["件数"]}
```

## 校验规则

编织后自动执行：必填字段、重量逻辑、成交方式一致性、退税率提醒、金额对比、重量对比。

**知识库增强**（自动触发，无需配置）：

| 触发条件 | 增强动作 |
|---------|---------|
| FOB/CIF/CFR 术语不一致 | 查询「贸易术语速查」确认语义 |
| 出现 ERROR 级别结果 | 按「退单」标签搜索修复建议 |
| HS 编码缺失 (`00000000`) | 按品名搜索 `02-HS归类/` 给出归类建议 |

校验报告的 `suggestion` 字段会携带知识库返回的修复方向，在 CLI 和 Web 面板均可查看。

## 适配真实文件需改动的代码

| 场景 | 文件 | 改动 |
|------|------|------|
| 列名多了新叫法 | `parsers/excel_parser.py` | `INVOICE_DEFAULT_MAP` |
| 表头 label 不同 | `parsers/excel_parser.py` | `HEADER_FIELDS` |
| 报关单需要新字段 | `models/declaration.py` | Pydantic 字段 |
| 字段映射逻辑 | `core/weaver.py` | `weave()` |
| Excel 导出布局 | `core/exporter.py` | `export()` |
| 校验规则 | `core/validator.py` | `check_declaration()` |
| 知识库接口 | `core/knowledge.py` | `KnowledgeLookup` 类 |
| 历史案例回写 | `core/history.py` | `append_historical_case()` |

## 故障排除

| 现象 | 解决 |
|------|------|
| 打不开 | 用 `127.0.0.1:8888` 不是 `localhost` |
| 上传没反应 | `Cmd+Shift+R` 强制刷新 |
| 归档加载失败 | `bash restart.sh` |
| 知识库搜索失败 | 确认 `zhishiku.db` 已初始化：<br>`cd customs/zhishiku && python -c "from db.importer import import_all; import_all(clear=True)"` |
| 知识库搜索提示初始化失败 | Web 面板点击「🔄 重建索引」按钮手动触发 |
| 自动刷新不生效 | 重启 bao 后重试 |

## 📚 知识库（zhishiku）

`customs/zhishiku/` 为 `bao` 提供参考资料，在编织和校验报关单时按需检索。

### 知识分类

| 目录 | 内容 | bao 工作时的用途 |
|------|------|----------------|
| `01-法规税则/` | 出口关税税则、监管条件速查、FTA 优惠协定清单 | 确认监管条件和税率 |
| `02-HS归类/` | HS 编码规则、机械/电子/纺织品归类、疑难案例 | 发票品名 → HS 编码匹配 |
| `03-报关实务/` | 报关单填制规范、退单原因及处理、贸易术语速查、币制/运输方式代码表 | 字段语义校验、术语查询 |
| `04-单证模板/` | 发票、装箱单、报关单模板说明 | 理解单证结构 |
| `05-历史案例/` | 已完成报关单摘要索引 | 参考历史做法 |
| `06-常见问题/` | 发票、装箱单、报关单常见问题 FAQ | 快速排查异常 |

### 检索方式

```bash
# grep 按标签快速检索 Markdown 文件
grep -r "#HS归类" customs/zhishiku/
grep -ri "退单原因" customs/zhishiku/

# SQLite FTS5 全文搜索（推荐，支持中文分词）
python -c "
import sys; sys.path.insert(0,'customs/zhishiku')
from db import KnowledgeDB, KnowledgeRepository
repo = KnowledgeRepository(KnowledgeDB('customs/zhishiku/zhishiku.db'))
repo.search('HS编码 AND 机械')
"
```

### 自动索引刷新

`bao` 启动或首次查询知识库时，自动对比 `.md` 和 `.db` 的修改时间。
任一知识文件较新则自动重建 FTS5 索引，无需手动操作。

### 历史案例回写（闭环）

每次编织 + 归档成功后，自动将报关单摘要写入 `zhishiku/05-历史案例/`，并重建 FTS5。
下次编织同类商品时 weaver 可搜索到历史做法，形成「使用 → 积累 → 反哺」闭环。

- **CLI**：`bao build from-files --save` 时自动写入
- **Web 面板**：点击「编织生成报关单」后自动归档 + 写入

### 手动同步索引

Web 面板「📚 知识库」Tab 底部有「🔄 重建索引」按钮，可手动触发。
CLI 也可执行：

```bash
cd customs/zhishiku
python -c "from db.importer import import_all; import_all(clear=True)"
```

## 项目结构

```
bao/
├── restart.sh / start_web.py
├── bao/
│   ├── cli.py / commands/     # 5 条 CLI 命令
│   ├── core/                  # weaver / exporter / validator / knowledge / history
│   ├── models/                # 数据模型
│   ├── db/                    # SQLite
│   ├── parsers/               # Excel 解析器
│   └── web/                   # server.py + index.html + downloads/
├── templates/ / data/ / tests/
```

## 依赖

typer / rich / jinja2 / pandas / openpyxl / pydantic
