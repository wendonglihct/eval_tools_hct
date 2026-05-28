"""
评测报告配置加载与 KQL 风格 DSL 解释器（通用）。

参考 Auto_loop_sdk/astra_sdk/eval/sim_report_summary.py 的设计思路：
把"数据集检索条件 / 路径取值"抽象为可配置 DSL，业务侧（hct_report 等）只关心聚合与渲染。

DSL 语法:
    <field.path> : "value"           # 字段精确匹配（路径下任一值等于）
    <field.path> in ["v1","v2"]      # 任一值在集合中
    cond and cond / cond or cond / not cond / (cond)

路径前缀（scenario_object. / raw. / meta. / SaturnVScenario_Info.）会被自动剥离。
"""

import ast
import json
import os
import re
from typing import Any, Callable, Dict, List


# -------------------- 路径取值 --------------------

_STRIP_PREFIXES = (
    "scenario_object.",
    "raw.",
    "meta.",
    "SaturnVScenario_Info.",
)


def normalize_path(field_path: str) -> List[str]:
    """剥离公共前缀并按 '.' 切分。多次剥离以兼容 SaturnVScenario_Info.meta.xxx。"""
    path = str(field_path or "").strip()
    changed = True
    while changed:
        changed = False
        for prefix in _STRIP_PREFIXES:
            if path.startswith(prefix):
                path = path[len(prefix):]
                changed = True
    return [p for p in path.split(".") if p]


def collect_path_values(obj: Any, parts: List[str]) -> List[Any]:
    """沿 parts 递归取值；遇到 list 自动展开。"""
    if not parts:
        return [obj]
    if isinstance(obj, list):
        result: List[Any] = []
        for item in obj:
            result.extend(collect_path_values(item, parts))
        return result
    if not isinstance(obj, dict):
        return []
    head = parts[0]
    if head not in obj:
        return []
    return collect_path_values(obj.get(head), parts[1:])


def make_selector(field_path: str) -> Callable[[dict], List[str]]:
    """根据 selector 字符串生成 extractor: meta -> List[str]（去重去空）。"""
    parts = normalize_path(field_path)

    def _extract(meta: dict) -> List[str]:
        if not parts:
            return []
        values = collect_path_values(meta, parts)
        return [str(v) for v in values if v is not None and str(v) != ""]

    return _extract


# -------------------- KQL DSL --------------------

_IN_PATTERN = re.compile(r'([A-Za-z0-9_\.]+)\s+in\s+\[([^\]]*)\]', flags=re.IGNORECASE)
_COND_PATTERN = re.compile(r'([A-Za-z0-9_\.]+)\s*:\s*"([^"]*)"')
# 通配符：field : *  → 字段存在任意非空值即真
_WILDCARD_PATTERN = re.compile(r'([A-Za-z0-9_\.]+)\s*:\s*\*')


def _parse_in_values(text: str) -> List[str]:
    values: List[str] = []
    for m in re.finditer(r'"([^"]*)"|\'([^\']*)\'', text or ""):
        values.append(m.group(1) if m.group(1) is not None else m.group(2))
    return values


def _match_cond(meta: dict, field_path: str, expected: str) -> bool:
    parts = normalize_path(field_path)
    if not parts:
        return False
    expected_str = str(expected)
    return any(str(v) == expected_str for v in collect_path_values(meta, parts))


def _match_wildcard(meta: dict, field_path: str) -> bool:
    """字段路径上存在任意非空值即真。"""
    parts = normalize_path(field_path)
    if not parts:
        return False
    return any(v is not None and str(v) != "" for v in collect_path_values(meta, parts))


def _match_in(meta: dict, field_path: str, expected_values: List[str]) -> bool:
    parts = normalize_path(field_path)
    if not parts:
        return False
    values = {str(v) for v in collect_path_values(meta, parts)}
    expected = {str(v) for v in expected_values}
    return bool(values & expected)


def _safe_eval_bool(expr: str) -> bool:
    """只接受 True/False、and/or/not、括号；其余一律拒绝。"""
    node = ast.parse(expr, mode="eval")

    def _ev(n: ast.AST) -> bool:
        if isinstance(n, ast.Expression):
            return _ev(n.body)
        if isinstance(n, ast.Constant) and isinstance(n.value, bool):
            return bool(n.value)
        if isinstance(n, ast.Name):
            if n.id in ("True", "False"):
                return n.id == "True"
            raise ValueError(f"非法名称: {n.id}")
        if isinstance(n, ast.UnaryOp) and isinstance(n.op, ast.Not):
            return not _ev(n.operand)
        if isinstance(n, ast.BoolOp) and isinstance(n.op, ast.And):
            return all(_ev(v) for v in n.values)
        if isinstance(n, ast.BoolOp) and isinstance(n.op, ast.Or):
            return any(_ev(v) for v in n.values)
        raise ValueError(f"不支持的节点: {type(n).__name__}")

    return _ev(node)


def compile_query(query: str) -> Callable[[dict], bool]:
    """把 KQL 风格 query 编译为 predicate(meta)->bool。"""
    text = str(query or "").strip()
    if not text:
        raise ValueError("数据集检索为空")

    # 静态校验：把所有条件替换为 True 后，剩下的应只能是 bool 逻辑
    check = re.sub(r"\bAND\b", "and", text, flags=re.IGNORECASE)
    check = re.sub(r"\bOR\b", "or", check, flags=re.IGNORECASE)
    check = re.sub(r"\bNOT\b", "not", check, flags=re.IGNORECASE)
    check = _IN_PATTERN.sub("True", check)
    check = _COND_PATTERN.sub("True", check)
    check = _WILDCARD_PATTERN.sub("True", check)
    residue = re.sub(r"\b(True|False|and|or|not)\b", "", check).replace("(", "").replace(")", "").strip()
    if residue:
        raise ValueError(f"数据集检索存在无法解析的内容: {residue}")

    def predicate(meta: dict) -> bool:
        expr = text
        expr = re.sub(r"\bAND\b", "and", expr, flags=re.IGNORECASE)
        expr = re.sub(r"\bOR\b", "or", expr, flags=re.IGNORECASE)
        expr = re.sub(r"\bNOT\b", "not", expr, flags=re.IGNORECASE)
        expr = _IN_PATTERN.sub(
            lambda m: "True" if _match_in(meta, m.group(1), _parse_in_values(m.group(2))) else "False",
            expr,
        )
        expr = _COND_PATTERN.sub(
            lambda m: "True" if _match_cond(meta, m.group(1), m.group(2)) else "False",
            expr,
        )
        expr = _WILDCARD_PATTERN.sub(
            lambda m: "True" if _match_wildcard(meta, m.group(1)) else "False",
            expr,
        )
        return _safe_eval_bool(expr)

    return predicate


# -------------------- 配置加载 --------------------

CONFIG_DIR = os.path.join(os.path.dirname(__file__), "configs")


def get_config_path(name: str) -> str:
    """按名称定位 configs 目录下的 JSON 配置。"""
    return os.path.join(CONFIG_DIR, name)


def load_hct_feature_config(json_path: str = None) -> List[Dict[str, Any]]:
    """
    加载 HCT 特性配置并预编译。返回结构与 hct_report 旧版 _feature_configs() 兼容：

    [
      {
        "name": "...",
        "sql": "...",
        "predicate": Callable[[meta], bool],
        "primary_result": int,
        "results": [{name, agg, numerator, denominator, value_format,
                     rollback_threshold, optimize_threshold, note}, ...],
        "small_feature": {"field_sql": "...", "extractor": Callable[[meta], List[str]]}
      },
      ...
    ]
    """
    path = json_path or get_config_path("hct_report_config.json")
    with open(path, "r", encoding="utf-8") as f:
        raw = json.load(f)
    if not isinstance(raw, dict):
        raise ValueError(f"配置根节点必须是对象: {path}")

    features: List[Dict[str, Any]] = []
    for feature_name, item in raw.items():
        if not isinstance(item, dict):
            raise ValueError(f"{feature_name} 配置必须是对象")
        # 空对象视为占位（未配置），跳过
        if not item:
            continue
        query = str(item.get("数据集检索", "")).strip()
        if not query:
            raise ValueError(f"{feature_name} 缺少 数据集检索")
        metric_list = item.get("指标配置", [])
        if not isinstance(metric_list, list) or not metric_list:
            raise ValueError(f"{feature_name} 指标配置必须是非空数组")

        # 过滤 enabled=false 的指标项（缺省视为启用）
        original_primary = int(item.get("primary_result", 0))
        kept: List[Dict[str, Any]] = []
        new_primary = 0
        for idx, m in enumerate(metric_list):
            if not isinstance(m, dict):
                continue
            if not bool(m.get("enabled", True)):
                continue
            if idx == original_primary:
                new_primary = len(kept)
            kept.append(m)
        if not kept:
            raise ValueError(f"{feature_name} 指标配置全部 enabled=false，无可用项")
        # 若原 primary_result 被禁用，回退到首项
        if original_primary >= len(metric_list) or not bool(
            metric_list[original_primary].get("enabled", True)
            if isinstance(metric_list[original_primary], dict)
            else False
        ):
            new_primary = 0
        metric_list = kept

        # 小特性：仅支持 list（多小特性）。每项 {enabled, selector}。
        raw_small = item.get("小特性")
        if raw_small is None:
            raw_small_list = []
        elif isinstance(raw_small, list):
            raw_small_list = raw_small
        else:
            raise ValueError(f"{feature_name} 小特性必须是数组")

        small_features: List[Dict[str, Any]] = []
        for sf in raw_small_list:
            if not isinstance(sf, dict):
                raise ValueError(f"{feature_name} 小特性条目必须是对象")
            sel = str(sf.get("selector", "") or "").strip()
            en = bool(sf.get("enabled", False)) and bool(sel)
            small_features.append({
                "field_sql": sel if en else "",
                "extractor": make_selector(sel) if en else (lambda _meta: []),
                "enabled": en,
            })
        # 至少占位一个，保持下游遍历稳定
        if not small_features:
            small_features = [{"field_sql": "", "extractor": (lambda _meta: []), "enabled": False}]

        features.append(
            {
                "name": str(feature_name),
                "sql": query,
                "predicate": compile_query(query),
                "primary_result": new_primary,
                "results": list(metric_list),
                "small_features": small_features,
            }
        )
    if not features:
        raise ValueError(f"配置为空: {path}")
    return features
