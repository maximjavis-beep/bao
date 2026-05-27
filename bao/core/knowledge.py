"""知识库接口 — 打通 zhishiku 调用链路

提供 HS 归类检索、全文搜索和分类浏览能力，
供 weaver（自动补全 HS 编码）和 Web 面板（知识库 Tab）使用。
"""

import sys
from pathlib import Path
from typing import Dict, List, Optional

# zhishiku 的 db 模块路径
_ZHISHIKU_ROOT = Path(__file__).parent.parent.parent.parent / "zhishiku"
if str(_ZHISHIKU_ROOT) not in sys.path:
    sys.path.insert(0, str(_ZHISHIKU_ROOT))

try:
    from db import KnowledgeDB, KnowledgeRepository
except ImportError:
    KnowledgeDB = None
    KnowledgeRepository = None


def _needs_refresh(db_path: Path, zhishiku_root: Path) -> bool:
    """检测知识文件是否比数据库更新（含 .md 和 Word/Excel/PDF）"""
    if not db_path.exists():
        return True
    try:
        db_mtime = db_path.stat().st_mtime
    except OSError:
        return True

    # 遍历分类目录下的 .md 文件
    for md in sorted(zhishiku_root.rglob("*.md")):
        parts = md.relative_to(zhishiku_root).parts
        if any(p.startswith(".") or p == "db" for p in parts):
            continue
        try:
            if md.stat().st_mtime > db_mtime:
                return True
        except OSError:
            continue

    # 检测 Word/Excel/PDF 文件是否有未转换的
    _SUPPORTED = {".xlsx", ".docx", ".pdf"}
    for ext in _SUPPORTED:
        for f in zhishiku_root.rglob(f"*{ext}"):
            parts = f.relative_to(zhishiku_root).parts
            if any(p.startswith(".") or p == "db" for p in parts):
                continue
            # 对应 .md 不存在 或 源文件比 .md 新 → 需要刷新
            md_path = f.with_suffix(".md")
            try:
                if not md_path.exists():
                    return True
                if f.stat().st_mtime > md_path.stat().st_mtime:
                    return True
            except OSError:
                continue

    return False


class KnowledgeLookup:
    """知识库查询封装 — 单次会话内复用连接"""

    _instance: Optional["KnowledgeLookup"] = None

    @classmethod
    def get(cls, db_path: Optional[str] = None) -> "KnowledgeLookup":
        """获取或创建全局查询实例"""
        if cls._instance is None:
            cls._instance = cls(db_path)
        return cls._instance

    @classmethod
    def reset(cls):
        """关闭并重置实例"""
        if cls._instance:
            cls._instance.close()
            cls._instance = None

    def __init__(self, db_path: Optional[str] = None):
        if KnowledgeDB is None:
            raise RuntimeError(
                "zhishiku db 模块未找到，请确认 customs/zhishiku/db/ 存在"
            )
        if db_path is None:
            db_path = str(_ZHISHIKU_ROOT / "zhishiku.db")
        db_path_obj = Path(db_path)

        if not db_path_obj.exists():
            raise FileNotFoundError(f"知识库数据库文件不存在: {db_path}")

        # 自动检测 Markdown 文件变更，按需重建索引
        if _needs_refresh(db_path_obj, _ZHISHIKU_ROOT):
            import sys as _sys
            print("[zhishiku] 检测到知识文件变更，自动重建索引…", file=_sys.stderr)
            try:
                KnowledgeLookup.refresh()
                print("[zhishiku] ✓ 索引重建完成", file=_sys.stderr)
            except Exception as e:
                print(f"[zhishiku] ⚠ 自动刷新失败: {e}", file=_sys.stderr)
                print("[zhishiku] 将继续使用现有索引", file=_sys.stderr)

        self.db = KnowledgeDB(str(db_path_obj))
        self.repo = KnowledgeRepository(self.db)

    # ── 全文搜索 ──────────────────────────────────

    def search(self, query: str, limit: int = 20) -> List[Dict]:
        """FTS5 全文搜索"""
        try:
            return self.repo.search(query, limit=limit)
        except Exception:
            return []

    # ── HS 编码归类检索 ────────────────────────────

    def search_hs(self, product_name: str, limit: int = 5) -> List[Dict]:
        """按品名搜索 HS 编码归类建议

        策略：
        1. 在标题中搜索品名关键词
        2. 优先返回 02-HS归类 分类中的结果
        3. 回退到全库文本搜索
        """
        if not product_name or not product_name.strip():
            return []

        name = product_name.strip()

        # 先按分类 + 品名关键词搜索
        results = self.repo.search(name, limit=limit * 2)

        # 优先保留 02-HS归类 分类的条目
        hs_results = [
            r for r in results
            if r.get("category", "").startswith("02-")
        ]
        if not hs_results:
            hs_results = results

        # 对每条结果提取可能相关的片段
        return self._format_hs_suggestions(hs_results[:limit], name)

    def _format_hs_suggestions(
        self, results: List[Dict], query: str
    ) -> List[Dict]:
        """将知识条目格式化为 HS 编码建议"""
        suggestions = []
        for r in results:
            body = r.get("body", "")
            title = r.get("title", "")
            suggestions.append({
                "title": title,
                "category": r.get("category", ""),
                "source": r.get("source", ""),
                "snippet": self._extract_snippet(body, query, 120),
                "file_path": r.get("file_path", ""),
            })
        return suggestions

    @staticmethod
    def _extract_snippet(text: str, query: str, max_len: int = 120) -> str:
        """从正文提取包含关键词的摘要片段"""
        if not text:
            return ""
        idx = text.find(query)
        if idx < 0:
            # 尝试前 3 个字符匹配
            idx = text.find(query[:3]) if len(query) >= 3 else -1
        if idx < 0:
            return text[:max_len] + ("..." if len(text) > max_len else "")
        start = max(0, idx - 30)
        end = min(len(text), idx + max_len - 30)
        snippet = text[start:end]
        if start > 0:
            snippet = "..." + snippet
        if end < len(text):
            snippet = snippet + "..."
        return snippet

    # ── 分类浏览 ──────────────────────────────────

    def search_by_category(self, category: str, limit: int = 50) -> List[Dict]:
        """按分类浏览"""
        try:
            return self.repo.list_by_category(category, limit=limit)
        except Exception:
            return []

    def search_by_tag(self, tag: str, limit: int = 50) -> List[Dict]:
        """按标签搜索"""
        try:
            return self.repo.search_by_tag(tag, limit=limit)
        except Exception:
            return []

    def list_categories(self) -> List[Dict]:
        """列出所有分类及条目数"""
        try:
            rows = self.db.execute(
                """SELECT category, COUNT(*) as cnt
                   FROM knowledge_entries
                   GROUP BY category
                   ORDER BY category"""
            ).fetchall()
            return [dict(r) for r in rows]
        except Exception:
            return []

    # ── 条目获取与索引维护 ─────────────────────────

    def get_entry(self, file_path: str) -> Optional[Dict]:
        """按文件路径获取完整知识条目"""
        try:
            return self.repo.get_by_path(file_path)
        except Exception:
            return None

    @staticmethod
    def refresh():
        """重建 FTS5 全文搜索索引

        新增或修改 Markdown 知识文件后调用，将文件内容重新导入数据库。
        """
        import subprocess
        import sys as _sys
        zhishiku_dir = str(_ZHISHIKU_ROOT)
        try:
            result = subprocess.run(
                [_sys.executable, "-c",
                 "from db.importer import import_all; import_all(clear=True)"],
                cwd=zhishiku_dir,
                capture_output=True,
                text=True,
                timeout=30,
            )
            if result.returncode != 0:
                raise RuntimeError(result.stderr.strip() or "索引重建失败")
        except subprocess.TimeoutExpired:
            raise RuntimeError("索引重建超时")

    def close(self):
        if self.db:
            self.db.close()
        KnowledgeLookup._instance = None
