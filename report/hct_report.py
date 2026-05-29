import os
import re
import sys
import operator
import subprocess
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# 允许直接 `python report/hct_report.py` 运行（debugger 场景）
_PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from report.config_loader import load_hct_feature_config, get_config_path


PASS_STATUS = {"Pass", "EVAL FAIL"}

# 运行时产物根目录（task 缓存 + 报告产出）
OUTPUTS_ROOT = os.path.join(_PROJECT_ROOT, "outputs")
TASKS_DIR = os.path.join(OUTPUTS_ROOT, "tasks")
REPORTS_DIR = os.path.join(OUTPUTS_ROOT, "reports")


def _ensure_output_dirs():
    os.makedirs(TASKS_DIR, exist_ok=True)
    os.makedirs(REPORTS_DIR, exist_ok=True)


def _load_task_json(task_id):
    _ensure_output_dirs()
    file_path = os.path.join(TASKS_DIR, f"task_{task_id}.json")
    if not os.path.exists(file_path):
        cmd = ["pncops", "task-detail", "--task_id", str(task_id), "-o", file_path]
        subprocess.run(cmd, check=True)
    import json
    with open(file_path, "r", encoding="utf-8") as f:
        return json.load(f)


def _normalize_ver_map(ver_map, target_len):
    ver_map = list(ver_map or [])
    if len(ver_map) >= target_len:
        return ver_map[:target_len]
    idx = 0
    while len(ver_map) < target_len:
        q = idx
        label = ""
        while True:
            label = chr(ord("M") + (q % 26)) + label
            q = q // 26 - 1
            if q < 0:
                break
        ver_map.append(label)
        idx += 1
    return ver_map


def _extract_sim_records(task_id):
    data = _load_task_json(task_id)
    records = []
    # sim 顶层这些字段也注入 meta，使 DSL 中 `scenario_object.<field>` 写法可解析
    # （DSL 会剥离 scenario_object./meta. 前缀，最终在 meta 里查找）
    SIM_LEVEL_FIELDS = (
        "tenant", "task_id", "task_hardware_type", "task_billing_pool",
        "data_billing_type", "agent_name", "agent_queue_name",
        "scenario_start_time", "scenario_execution_duration", "task_create_time",
    )
    for sim in data.get("es_sim", []):
        if sim.get("scenario_execution_result") not in PASS_STATUS:
            continue
        raw = sim.get("scenario_object", {}).get("raw", {})
        meta = dict(raw.get("meta", {}) or {})
        for k in SIM_LEVEL_FIELDS:
            if k in sim and k not in meta:
                meta[k] = sim[k]
        records.append(
            {
                "scenario_id": sim.get("scenario_id"),
                "meta": meta,
                "evaluation_result": sim.get("evaluation_result", {}) or {},
            }
        )
    return records


def _build_task_records_map(task_ids):
    task_to_records = {}
    for task_id in task_ids:
        task_to_records[str(task_id)] = _extract_sim_records(str(task_id))
    return task_to_records


def _align_by_base_pairwise(task_to_records, base_task):
    # 每个版本与 base 两两取交，避免“全量交集”导致统计偏小
    aligned_task_records = {}
    aligned_base_records = {}
    base_records = task_to_records.get(base_task, [])
    base_ids = {r.get("scenario_id") for r in base_records}
    aligned_task_records[base_task] = list(base_records)
    aligned_base_records[base_task] = list(base_records)
    for task_id, records in task_to_records.items():
        if task_id == base_task:
            continue
        task_ids = {r.get("scenario_id") for r in records}
        pair_common = base_ids & task_ids
        aligned_task_records[task_id] = [r for r in records if r.get("scenario_id") in pair_common]
        aligned_base_records[task_id] = [r for r in base_records if r.get("scenario_id") in pair_common]
    return aligned_task_records, aligned_base_records


def _metric_value(rec, metric_key):
    metric = rec.get("evaluation_result", {}).get(metric_key, {})
    value = metric.get("metric_value", 0)
    try:
        return float(value)
    except Exception:
        return 0.0


def _parse_result_agg(agg_mode):
    """agg='metric_result:<bool>' 时返回目标布尔值；否则返回 None。
    支持 true/false、1/0（大小写不敏感）。"""
    text = str(agg_mode or "").strip().lower()
    prefix = "metric_result:"
    if not text.startswith(prefix):
        return None
    target = text[len(prefix):].strip()
    if target in ("true", "1"):
        return True
    if target in ("false", "0"):
        return False
    return None


def _metric_result_equals(rec, metric_key, target):
    """metric_result 为目标布尔值时返回 1.0，否则 0.0。"""
    metric = rec.get("evaluation_result", {}).get(metric_key, {})
    raw = metric.get("metric_result", None)
    if isinstance(raw, bool):
        val = raw
    elif isinstance(raw, str):
        s = raw.strip().lower()
        if s in ("true", "1"):
            val = True
        elif s in ("false", "0"):
            val = False
        else:
            return 0.0
    else:
        return 0.0
    return 1.0 if val == target else 0.0


def _record_term_value(rec, metric_key, agg_mode):
    """逐条记录的指标取值：
    - agg='metric_result:<bool>' → 1/0 计数
    - 其他 agg（sum/ave）→ metric_value
    """
    target = _parse_result_agg(agg_mode)
    if target is not None:
        return _metric_result_equals(rec, metric_key, target)
    return _metric_value(rec, metric_key)


# ------ 记录级 filter ------
_FILTER_VALUE_RE = re.compile(
    r'^\s*([A-Za-z0-9_]+)\s*\.\s*metric_value\s*(<=|>=|==|!=|<|>)\s*(-?\d+(?:\.\d+)?)\s*$'
)
_FILTER_RESULT_RE = re.compile(
    r'^\s*([A-Za-z0-9_]+)\s*\.\s*metric_result\s*:\s*(true|false|0|1)\s*$',
    flags=re.IGNORECASE,
)
_FILTER_OPS = {
    "<": operator.lt, "<=": operator.le,
    ">": operator.gt, ">=": operator.ge,
    "==": operator.eq, "!=": operator.ne,
}


def _compile_record_filter(filter_str):
    """把 filter 表达式编译为 Callable[[rec], bool]。空/None → 返回 None。"""
    if not filter_str:
        return None
    text = str(filter_str).strip()
    if not text:
        return None

    m = _FILTER_VALUE_RE.match(text)
    if m:
        key, op_text, num_text = m.group(1), m.group(2), m.group(3)
        cmp = _FILTER_OPS[op_text]
        threshold = float(num_text)

        def _f(rec):
            metric = rec.get("evaluation_result", {}).get(key)
            if not isinstance(metric, dict) or "metric_value" not in metric:
                return False
            v = metric.get("metric_value")
            if v is None:
                return False
            try:
                return cmp(float(v), threshold)
            except Exception:
                return False
        return _f

    m = _FILTER_RESULT_RE.match(text)
    if m:
        key, tgt = m.group(1), m.group(2).lower()
        target = tgt in ("true", "1")

        def _f(rec):
            return _metric_result_equals(rec, key, target) > 0.5
        return _f

    raise ValueError(
        f"无法解析 filter: {filter_str!r}（支持 '<key>.metric_value <op> <num>' 或 '<key>.metric_result:<bool>'）"
    )


def _parse_metric_expression(expr):
    text = str(expr or "").strip()
    if not text:
        return []
    norm = text.replace(" ", "")
    parts = re.findall(r"[+\-]?[^+\-]+", norm)
    terms = []
    for idx, part in enumerate(parts):
        if not part:
            continue
        sign = 1
        token = part
        if token[0] == "+":
            token = token[1:]
        elif token[0] == "-":
            sign = -1
            token = token[1:]
        elif idx > 0:
            sign = 1
        if token:
            terms.append((sign, token))
    return terms


def _calc_expr_on_record(rec, expr, agg_mode="sum"):
    terms = _parse_metric_expression(expr)
    if not terms:
        return 0.0
    value = 0.0
    for sign, metric_key in terms:
        value += sign * _record_term_value(rec, metric_key, agg_mode)
    return value


def _calc_expr_value_or_none(rec, expr, agg_mode):
    """ave 模式专用：任一 term 的 metric_value 缺失/不可解析 → 返回 None；否则返回 float。
    metric_result:<bool> 模式不视为缺测，沿用计数语义。"""
    terms = _parse_metric_expression(expr)
    if not terms:
        return 0.0
    if _parse_result_agg(agg_mode) is not None:
        return _calc_expr_on_record(rec, expr, agg_mode)
    value = 0.0
    for sign, metric_key in terms:
        m = rec.get("evaluation_result", {}).get(metric_key)
        if not isinstance(m, dict) or "metric_value" not in m:
            return None
        raw = m.get("metric_value")
        if raw is None:
            return None
        try:
            value += sign * float(raw)
        except Exception:
            return None
    return value


def _aggregate_selected_metric(records, expr, agg_mode):
    if not records:
        return 0.0
    if str(agg_mode).strip().lower() == "ave":
        # ave 模式跳过缺测：仅对 metric_value 实际存在的记录求平均
        values = [v for v in (_calc_expr_value_or_none(r, expr, agg_mode) for r in records) if v is not None]
        return sum(values) / len(values) if values else 0.0
    values = [_calc_expr_on_record(r, expr, agg_mode) for r in records]
    return sum(values)


def _expr_label(expr, agg_mode, suffix):
    terms = _parse_metric_expression(expr)
    if not terms:
        return f"{agg_mode}()_{suffix}"
    pieces = []
    for idx, (sign, metric_key) in enumerate(terms):
        token = f"{agg_mode}({metric_key})_{suffix}"
        if idx == 0:
            pieces.append(token if sign > 0 else f"-{token}")
        else:
            op = "+" if sign > 0 else "-"
            pieces.append(f"{op}{token}")
    return "".join(pieces)


def _build_formula_text(agg_mode, numerator_expr, denominator_expr):
    num_comp = _expr_label(numerator_expr, agg_mode, "comp")
    num_base = _expr_label(numerator_expr, agg_mode, "base")
    if denominator_expr:
        den_base = _expr_label(denominator_expr, agg_mode, "base")
        return f"(({num_comp})-({num_base}))/({den_base})"
    return f"({num_comp})-({num_base})"


def _aggregate_stats_from_selected(selected_records, agg_mode, metric_expr, record_filter=None):
    """record_filter: 可选 Callable[[rec], bool]，仅影响 metric_sum；count 保留全量。"""
    stats = {}
    for task_id, records in selected_records.items():
        filt_records = [r for r in records if record_filter(r)] if record_filter else records
        metric_sum = _aggregate_selected_metric(filt_records, metric_expr, agg_mode)
        stats[task_id] = {"count": len(records), "metric_sum": metric_sum, "records": records}
    return stats


def _feature_configs(mode=None):
    """按 mode 选择 configs/ 目录下的配置文件并加载。
    - 当前 DT_HCT → hct_report_config.json
    - 其他 mode（含 None） → 暂时回退到 hct_report_config.json
    新 mode 落地时只需在 MODE_CONFIG_MAP 中追加映射。
    返回结构：
      [{name, sql, predicate, primary_result, results,
        small_features:[{field_sql, extractor, enabled}, ...]}, ...]
    """
    name = MODE_CONFIG_MAP.get(mode or "", DEFAULT_CONFIG_FILE)
    return load_hct_feature_config(get_config_path(name))


# mode → configs/ 下的 JSON 文件名映射；未命中走 DEFAULT_CONFIG_FILE
DEFAULT_CONFIG_FILE = "hct_report_config.json"
MODE_CONFIG_MAP = {
    "DT_HCT": "hct_report_config.json",
    # 例: "DT_HCT_close": "hct_report_close_config.json",
}


def _format_value(value, value_format):
    fmt = str(value_format or "float").lower()
    if fmt == "%":
        return f"{value:.2f}%"
    if fmt == "int":
        return str(int(round(value)))
    # float 默认保留 3 位
    return f"{value:.3f}"


def _format_diff(delta, denominator, use_denominator, value_format):
    if not use_denominator:
        return _format_value(delta, value_format)
    if denominator == 0:
        if delta == 0:
            return _format_value(0.0, value_format)
        return "N/A"
    ratio_percent = (delta / denominator) * 100
    return _format_value(ratio_percent, value_format)


def _format_base_denominators(task_order, base_denominator_stats, use_denominator):
    if not use_denominator:
        return "-"
    compare_task_ids = task_order[1:] if len(task_order) > 1 else task_order
    values = []
    for task_id in compare_task_ids:
        den = base_denominator_stats.get(task_id, {}).get("metric_sum", 0)
        try:
            values.append(f"{int(den)}")
        except Exception:
            values.append("0")
    return "/".join(values) if values else "0"


def _format_per_version_raw(value, value_format):
    """单值展示格式：
    - value_format=='float' 时按 float 保留 3 位（如 score）
    - 其他（int/%）按 int 展示（如帧数）
    """
    fmt = str(value_format or "float").lower()
    try:
        v = float(value)
    except Exception:
        v = 0.0
    if fmt == "float":
        return f"{v:.3f}"
    return str(int(round(v)))


def _format_denominator_column(version_denominator_stats, task_order, use_denominator):
    """基数(总数)列：百分比指标显示各版本 denominator，'/' 拼接；否则占位 '-'。"""
    if not use_denominator:
        return "-"
    pieces = []
    for task_id in task_order:
        den = version_denominator_stats.get(task_id, {}).get("metric_sum", 0)
        try:
            pieces.append(f"{int(den)}")
        except Exception:
            pieces.append("0")
    return "/".join(pieces) if pieces else "-"


def _build_row(
    data_class,
    sql_text,
    formula_text,
    result_name,
    note,
    rollback_threshold,
    optimize_threshold,
    task_order,
    stats,
    base_ref_stats,
    base_denominator_stats,
    version_denominator_stats,
    use_denominator,
    value_format,
):
    base_task = task_order[0]
    base_sum = stats.get(base_task, {}).get("metric_sum", 0.0)
    ref_task_for_count = task_order[1] if len(task_order) > 1 else base_task
    base_count = base_ref_stats.get(ref_task_for_count, {}).get("count", stats.get(base_task, {}).get("count", 0))
    denominator_cell = _format_denominator_column(version_denominator_stats, task_order, use_denominator)
    row = [
        data_class,
        sql_text,
        base_count,
        formula_text,
        result_name,
        rollback_threshold,
        optimize_threshold,
        note,
        denominator_cell,
    ]
    for task_id in task_order:
        s = stats.get(task_id, {"count": 0, "metric_sum": 0.0})
        if task_id == base_task:
            # base 列：显示「base 全量 (与每个对比版本对齐后的 base 值, '/' 拼接)」
            # - 全量值放在括号外，作为参考基准
            # - 括号内每段对应一个对比版本，便于核对 A 列 = aligned_base + diff
            base_full = _format_per_version_raw(s.get("metric_sum", 0.0), value_format)
            compare_tasks = task_order[1:]
            aligned_pieces = []
            for cmp_task in compare_tasks:
                aligned_base = base_ref_stats.get(cmp_task, {"metric_sum": base_sum}).get("metric_sum", 0.0)
                aligned_pieces.append(_format_per_version_raw(aligned_base, value_format))
            if aligned_pieces:
                row.append(f"{base_full}({'/'.join(aligned_pieces)})")
            else:
                row.append(base_full)
            continue
        version_value = _format_per_version_raw(s.get("metric_sum", 0.0), value_format)
        base_ref = base_ref_stats.get(task_id, {"metric_sum": base_sum})
        base_den = base_denominator_stats.get(task_id, {"metric_sum": 0.0})
        diff = _format_diff(
            s["metric_sum"] - base_ref["metric_sum"],
            base_den["metric_sum"],
            use_denominator,
            value_format,
        )
        row.append(f"{version_value}({diff})")
    return row


def _select_records_by_predicate(task_records, predicate):
    selected = {}
    for task_id, records in task_records.items():
        selected[task_id] = [r for r in records if predicate(r.get("meta", {}))]
    return selected


def _init_task_stats(task_ids):
    return {task_id: {"count": 0, "metric_sum": 0.0} for task_id in task_ids}


def _aggregate_small_feature_stats(records_by_task, extractor, agg_mode, metric_expr, task_ids, record_filter=None):
    """
    单次遍历聚合小特性统计：
    - value_stats[value][task_id] = {"count", "metric_sum"}
    - other_stats[task_id] = {"count", "metric_sum"}  # extractor 无值时归档
    注：
    - 按 records_by_task 的 dict key 迭代，避免 task_ids 列表中存在重复 id 时双倍累加。
    - ave 模式跳过 metric_value 缺测的记录：count 仍计入（用于"数量"列），
      但 valid_count 只统计非缺测；归一化用 valid_count，与离线一致。
    - record_filter: 可选 Callable[[rec], bool]。未通过 filter 的记录仅计入 count，
      不参与 metric_sum/valid_count。
    """
    value_stats = {}
    other_stats = _init_task_stats(task_ids)
    is_ave = str(agg_mode).strip().lower() == "ave"

    def _accum(target, metric_val, valid):
        target["count"] += 1
        if is_ave:
            if valid:
                target["metric_sum"] += metric_val
                target["valid_count"] = target.get("valid_count", 0) + 1
        else:
            target["metric_sum"] += metric_val

    for task_id, recs in records_by_task.items():
        for rec in recs:
            passes = record_filter(rec) if record_filter else True
            if not passes:
                # 不通过 filter 的记录：仅 count，不进 metric_sum / valid_count
                valid = False
                metric_val = 0.0
            elif is_ave:
                opt = _calc_expr_value_or_none(rec, metric_expr, agg_mode)
                valid = opt is not None
                metric_val = opt if valid else 0.0
            else:
                metric_val = _calc_expr_on_record(rec, metric_expr, agg_mode)
                valid = True
            raw_values = extractor(rec.get("meta", {}))
            values = sorted({v for v in raw_values if v})
            if not values:
                _accum(other_stats[task_id], metric_val, valid)
                continue
            for value in values:
                if value not in value_stats:
                    value_stats[value] = _init_task_stats(task_ids)
                _accum(value_stats[value][task_id], metric_val, valid)
    if is_ave:
        for stats in value_stats.values():
            for tid in stats:
                vc = stats[tid].get("valid_count", 0)
                stats[tid]["metric_sum"] = (stats[tid]["metric_sum"] / vc) if vc else 0.0
        for tid in other_stats:
            vc = other_stats[tid].get("valid_count", 0)
            other_stats[tid]["metric_sum"] = (other_stats[tid]["metric_sum"] / vc) if vc else 0.0
    return value_stats, other_stats


def _sum_counts(stats_by_task):
    return sum(v["count"] for v in stats_by_task.values())


def _merge_columns_by_ranges(ws, merge_ranges, col_indices=(1, 2, 3)):
    for start_row, end_row in merge_ranges:
        if end_row <= start_row:
            continue
        for col_idx in col_indices:
            ws.merge_cells(start_row=start_row, start_column=col_idx, end_row=end_row, end_column=col_idx)
            cell = ws.cell(row=start_row, column=col_idx)
            cell.alignment = Alignment(horizontal="left", vertical="center")


def _stats_or_default(stats_map, key, task_ids):
    return stats_map.get(key, _init_task_stats(task_ids))


def _build_result_bundle(selected_records, base_selected_records, small_features, result_item, task_ids):
    """small_features: 列表，每项 {field_sql, extractor, enabled}。
    返回 bundle 中 small_stats 为同长度列表，按下标对应每个小特性。"""
    agg_mode = result_item["agg"]
    numerator_expr = result_item["numerator"]
    denominator_expr = result_item.get("denominator")
    use_denominator = bool(denominator_expr)
    value_format = result_item.get("value_format", "%" if use_denominator else "float")
    record_filter = _compile_record_filter(result_item.get("filter"))

    stats = _aggregate_stats_from_selected(selected_records, agg_mode, numerator_expr, record_filter)
    base_ref_stats = _aggregate_stats_from_selected(base_selected_records, agg_mode, numerator_expr, record_filter)
    if use_denominator:
        base_denominator_stats = _aggregate_stats_from_selected(base_selected_records, agg_mode, denominator_expr, record_filter)
        version_denominator_stats = _aggregate_stats_from_selected(selected_records, agg_mode, denominator_expr, record_filter)
    else:
        base_denominator_stats = _init_task_stats(task_ids)
        version_denominator_stats = _init_task_stats(task_ids)

    small_stats_list = []
    for sf in small_features:
        ex = sf["extractor"]
        value_stats, other_stats = _aggregate_small_feature_stats(
            selected_records, ex, agg_mode, numerator_expr, task_ids, record_filter
        )
        base_value_stats, base_other_stats = _aggregate_small_feature_stats(
            base_selected_records, ex, agg_mode, numerator_expr, task_ids, record_filter
        )
        if use_denominator:
            base_value_den_stats, base_other_den_stats = _aggregate_small_feature_stats(
                base_selected_records, ex, agg_mode, denominator_expr, task_ids, record_filter
            )
            version_value_den_stats, version_other_den_stats = _aggregate_small_feature_stats(
                selected_records, ex, agg_mode, denominator_expr, task_ids, record_filter
            )
        else:
            base_value_den_stats = {}
            base_other_den_stats = _init_task_stats(task_ids)
            version_value_den_stats = {}
            version_other_den_stats = _init_task_stats(task_ids)
        small_stats_list.append({
            "field_sql": sf.get("field_sql", ""),
            "enabled": sf.get("enabled", True),
            "value_stats": value_stats,
            "other_stats": other_stats,
            "base_value_stats": base_value_stats,
            "base_other_stats": base_other_stats,
            "base_value_den_stats": base_value_den_stats,
            "base_other_den_stats": base_other_den_stats,
            "version_value_den_stats": version_value_den_stats,
            "version_other_den_stats": version_other_den_stats,
        })

    return {
        "result_item": result_item,
        "use_denominator": use_denominator,
        "formula_text": _build_formula_text(agg_mode, numerator_expr, denominator_expr),
        "value_format": value_format,
        "stats": stats,
        "base_ref_stats": base_ref_stats,
        "base_denominator_stats": base_denominator_stats,
        "version_denominator_stats": version_denominator_stats,
        "small_stats": small_stats_list,
    }


def _apply_threshold_coloring(ws, rollback_col=5, optimize_col=6, first_result_col=8):
    # 参考 generate_execl_mix_sample_0919.py 的阈值着色策略
    better_colors = ["FFFFFFFF", "FFD9F5D6", "FF8EE085"]
    worse_colors = ["FFFFFFFF", "FFFAF1D1", "FFFAD355", "FFFBBFBC"]

    def between(value, low, high):
        return low < value < high

    op_handles = {
        ">": operator.gt,
        "<": operator.lt,
        ">=": operator.ge,
        "<=": operator.le,
        "between": between,
    }

    pattern = r"(<=|>=|<|>)\s*(-?\d+\.?\d*)(?:%)?"

    for row_idx in range(2, ws.max_row + 1):
        # 跳过真正空行；注意合并单元格行在 A 列可能为 None，不能据此跳过
        row_has_value = any(ws.cell(row=row_idx, column=col_idx).value not in (None, "") for col_idx in range(1, ws.max_column + 1))
        if not row_has_value:
            continue

        rollback_number, optimize_number = None, None
        rollback_op, optimize_op = None, None

        rollback_cell = ws.cell(row=row_idx, column=rollback_col).value
        if isinstance(rollback_cell, str):
            matches = re.match(pattern, rollback_cell.strip())
            if matches:
                rollback_op = matches.group(1)
                rollback_number = float(matches.group(2))

        optimize_cell = ws.cell(row=row_idx, column=optimize_col).value
        if isinstance(optimize_cell, str):
            matches = re.match(pattern, optimize_cell.strip())
            if matches:
                optimize_op = matches.group(1)
                optimize_number = float(matches.group(2))

        for col_idx in range(first_result_col, ws.max_column + 1):
            raw = ws.cell(row=row_idx, column=col_idx).value
            if raw is None or raw == "":
                continue
            # 单元格格式为 "value(diff)"，着色依据为括号内的 diff
            m = re.search(r"\(([^)]+)\)", str(raw))
            inner = m.group(1) if m else str(raw)
            try:
                column_value = float(inner.replace("%", "").replace("=", "").strip())
            except Exception:
                continue

            if rollback_number is not None and optimize_number is None:
                if not op_handles[rollback_op](column_value, rollback_number):
                    ws.cell(row=row_idx, column=col_idx).fill = PatternFill(fill_type="solid", start_color="FFF54A45")
                continue

            if rollback_number is None and optimize_number is not None:
                if op_handles[optimize_op](column_value, optimize_number):
                    ws.cell(row=row_idx, column=col_idx).fill = PatternFill(fill_type="solid", start_color="FF34C724")
                continue

            if rollback_number is None and optimize_number is None:
                continue

            low_value = min(rollback_number, optimize_number)
            high_value = max(rollback_number, optimize_number)

            if abs(column_value) < 0.001:
                ws.cell(row=row_idx, column=col_idx).fill = PatternFill(fill_type="solid", start_color="FFFFFFFF")
            elif op_handles["between"](column_value, low_value, high_value):
                if (optimize_number > rollback_number) == op_handles["between"](column_value, 0.0, high_value):
                    discrete = (optimize_number - 0.0) / len(better_colors) if optimize_number != 0 else 1
                    idx = int((column_value - 0.0) / discrete) if discrete else 0
                    idx = max(0, min(len(better_colors) - 1, idx))
                    ws.cell(row=row_idx, column=col_idx).fill = PatternFill(fill_type="solid", start_color=better_colors[idx])
                else:
                    discrete = (rollback_number - 0.0) / len(worse_colors) if rollback_number != 0 else 1
                    idx = int((column_value - 0.0) / discrete) if discrete else 0
                    idx = max(0, min(len(worse_colors) - 1, idx))
                    ws.cell(row=row_idx, column=col_idx).fill = PatternFill(fill_type="solid", start_color=worse_colors[idx])
            else:
                if column_value == rollback_number:
                    ws.cell(row=row_idx, column=col_idx).fill = PatternFill(fill_type="solid", start_color="FFF54A45")
                    continue
                if column_value == optimize_number:
                    ws.cell(row=row_idx, column=col_idx).fill = PatternFill(fill_type="solid", start_color="FF34C724")
                    continue
                if rollback_number <= optimize_number:
                    if op_handles[rollback_op](column_value, rollback_number):
                        ws.cell(row=row_idx, column=col_idx).fill = PatternFill(fill_type="solid", start_color="FF34C724")
                    else:
                        ws.cell(row=row_idx, column=col_idx).fill = PatternFill(fill_type="solid", start_color="FFF54A45")
                else:
                    if op_handles[optimize_op](column_value, optimize_number):
                        ws.cell(row=row_idx, column=col_idx).fill = PatternFill(fill_type="solid", start_color="FF34C724")
                    else:
                        ws.cell(row=row_idx, column=col_idx).fill = PatternFill(fill_type="solid", start_color="FFF54A45")


def _relation_and_score(metric_map):
    # HCT 指标是“越小越好”，统一转成“分数越大越好”用于版本关系排序
    score_map = {k: -float(v) for k, v in metric_map.items()}
    sorted_items = sorted(score_map.items(), key=lambda x: x[1], reverse=True)
    relation = ">".join([k for k, _ in sorted_items]) if sorted_items else ""
    return relation, score_map


def _append_hct_report(msg, txt_path, all_ver_com, tag_lv1_ver_com, tag_all_ver_com):
    lines = []
    lines.append("================= 整体 =================")
    lines.append(f"版本关系: {all_ver_com.get('ver', '')}")
    lines.append("各版本得分:")
    for ver, score in all_ver_com.get("score", {}).items():
        lines.append(f"{ver}: {score:.4f}")
    lines.append("")
    lines.append("================= 四大特性 =================")
    for tag_lv1, info in tag_lv1_ver_com.items():
        lines.append(f"{tag_lv1}: {info.get('ver', '')}")
        lines.append("各版本得分:")
        for ver, score in info.get("score", {}).items():
            lines.append(f"{ver}: {score:.4f}")
        lines.append("")
    for domain, domain_info in tag_all_ver_com.items():
        lines.append(f"================= {domain} =================")
        for tag_all, info in domain_info.items():
            lines.append(f"{tag_all}: {info.get('ver', '')}")
            lines.append("各版本得分:")
            for ver, score in info.get("score", {}).items():
                lines.append(f"{ver}: {score:.4f}")
            lines.append("")

    content = "\n".join(lines).strip() + "\n"
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(content)
    if msg and not msg.endswith("\n"):
        msg += "\n"
    msg += content
    return msg


def _safe_sheet_title(raw_name, used_names):
    base = re.sub(r'[\[\]\:\*\?\/\\]', "_", str(raw_name or "DT_HCT"))
    base = base.strip() or "DT_HCT"
    if len(base) > 31:
        base = base[:31]
    title = base
    idx = 1
    while title in used_names:
        suffix = f"_{idx}"
        title = f"{base[:31 - len(suffix)]}{suffix}"
        idx += 1
    used_names.add(title)
    return title


def run_compare_hct(tasks_dict, ver_map, msg, file_name_suffix, mode="DT_HCT"):
    wb = Workbook()
    ws_default = wb.active
    used_sheet_names = set()
    has_written_sheet = False

    overall_metric_sums = defaultdict(float)
    tag_lv1_metric_sums = defaultdict(lambda: defaultdict(float))
    tag_all_metric_sums = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    task_order_for_rank = []
    feature_configs = _feature_configs(mode)
    major_feature_names = {f["name"] for f in feature_configs}
    for compare_key, task_ids in tasks_dict.items():
        all_rows = []
        task_ids = [str(x) for x in task_ids]
        if not task_ids:
            continue
        if not has_written_sheet:
            ws = ws_default
            ws.title = _safe_sheet_title(compare_key, used_sheet_names)
            has_written_sheet = True
        else:
            ws = wb.create_sheet(title=_safe_sheet_title(compare_key, used_sheet_names))
        for task_id in task_ids:
            if task_id not in task_order_for_rank:
                task_order_for_rank.append(task_id)
        versions = _normalize_ver_map(ver_map, len(task_ids))
        header_task_ver = [f"{v}({t})" for v, t in zip(versions, task_ids)]
        header = ["数据分类", "Sql语句", "数量", "分析逻辑", "评测内容", "回退阈值", "优秀阈值", "注释", "基数(总数)"] + header_task_ver
        all_rows.append(header)
        merge_ranges = []
        major_feature_row_idx = set()
        next_row_idx = 2

        task_records_raw = _build_task_records_map(task_ids)
        task_records, base_ref_records = _align_by_base_pairwise(task_records_raw, task_ids[0])

        for feature in feature_configs:
            selected_records = _select_records_by_predicate(task_records, feature["predicate"])
            base_selected_records = _select_records_by_predicate(base_ref_records, feature["predicate"])
            # 数据集检索命中为 0：跳过整个特性，不写入任何行
            total_hits = sum(len(recs) for recs in selected_records.values()) + sum(
                len(recs) for recs in base_selected_records.values()
            )
            if total_hits == 0:
                continue
            small_features = feature["small_features"]
            result_bundles = [
                _build_result_bundle(
                    selected_records=selected_records,
                    base_selected_records=base_selected_records,
                    small_features=small_features,
                    result_item=result_item,
                    task_ids=task_ids,
                )
                for result_item in feature["results"]
            ]
            primary_bundle = result_bundles[feature.get("primary_result", 0)]
            primary_stats = primary_bundle["stats"]
            for task_id in task_ids:
                metric_sum = float(primary_stats[task_id]["metric_sum"])
                overall_metric_sums[task_id] += metric_sum
                tag_lv1_metric_sums[feature["name"]][task_id] += metric_sum

            major_feature_row_idx.add(next_row_idx)
            feature_start_row = next_row_idx
            for idx, bundle in enumerate(result_bundles):
                result_item = bundle["result_item"]
                all_rows.append(
                    _build_row(
                        feature["name"] if idx == 0 else "",
                        feature["sql"] if idx == 0 else "",
                        bundle["formula_text"],
                        result_item["name"],
                        result_item["note"],
                        result_item["rollback_threshold"],
                        result_item["optimize_threshold"],
                        task_ids,
                        bundle["stats"],
                        bundle["base_ref_stats"],
                        bundle["base_denominator_stats"],
                        bundle["version_denominator_stats"],
                        bundle["use_denominator"],
                        bundle["value_format"],
                    )
                )
                next_row_idx += 1
            merge_ranges.append((feature_start_row, next_row_idx - 1))

            # 遍历每个小特性：分别渲染 value 行 + --other 行
            for sf_idx, sf_cfg in enumerate(small_features):
                if not sf_cfg.get("enabled", True):
                    continue
                field_sql = sf_cfg.get("field_sql", "")
                primary_small = primary_bundle["small_stats"][sf_idx]
                primary_value_stats = primary_small["value_stats"]
                primary_other_stats = primary_small["other_stats"]

                # 小特性按命中规模从大到小排序（总命中数量降序）
                value_order = sorted(
                    primary_value_stats.keys(),
                    key=lambda value: (-_sum_counts(primary_value_stats[value]), value),
                )

                for value in value_order:
                    for task_id in task_ids:
                        metric_sum = primary_value_stats[value][task_id]["metric_sum"]
                        tag_all_metric_sums["DT_HCT"][f'{feature["name"]}|{field_sql}|{value}'][task_id] = float(metric_sum)
                    small_start_row = next_row_idx
                    for idx, bundle in enumerate(result_bundles):
                        result_item = bundle["result_item"]
                        small = bundle["small_stats"][sf_idx]
                        all_rows.append(
                            _build_row(
                                f"--{value}" if idx == 0 else "",
                                f'{field_sql} = "{value}"' if idx == 0 else "",
                                bundle["formula_text"],
                                result_item["name"],
                                result_item["note"],
                                result_item["rollback_threshold"],
                                result_item["optimize_threshold"],
                                task_ids,
                                _stats_or_default(small["value_stats"], value, task_ids),
                                _stats_or_default(small["base_value_stats"], value, task_ids),
                                _stats_or_default(small["base_value_den_stats"], value, task_ids),
                                _stats_or_default(small["version_value_den_stats"], value, task_ids),
                                bundle["use_denominator"],
                                bundle["value_format"],
                            )
                        )
                        next_row_idx += 1
                    merge_ranges.append((small_start_row, next_row_idx - 1))

                # 未命中小特性字段的 case 统一归档到 --other
                for task_id in task_ids:
                    tag_all_metric_sums["DT_HCT"][f'{feature["name"]}|{field_sql}|other'][task_id] = float(
                        primary_other_stats[task_id]["metric_sum"]
                    )
                if _sum_counts(primary_other_stats) > 0:
                    other_start_row = next_row_idx
                    for idx, bundle in enumerate(result_bundles):
                        result_item = bundle["result_item"]
                        small = bundle["small_stats"][sf_idx]
                        all_rows.append(
                            _build_row(
                                "--other" if idx == 0 else "",
                                f'{field_sql} = "other"' if idx == 0 else "",
                                bundle["formula_text"],
                                result_item["name"],
                                result_item["note"],
                                result_item["rollback_threshold"],
                                result_item["optimize_threshold"],
                                task_ids,
                                small["other_stats"],
                                small["base_other_stats"],
                                small["base_other_den_stats"],
                                small["version_other_den_stats"],
                                bundle["use_denominator"],
                                bundle["value_format"],
                            )
                        )
                        next_row_idx += 1
                    merge_ranges.append((other_start_row, next_row_idx - 1))

        for row in all_rows:
            ws.append(row)
        _merge_columns_by_ranges(ws, merge_ranges, col_indices=(1, 2, 3))

        # 表头样式：居中 + 加粗
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # 大特性行：红色、加粗、字号放大
        for row_idx in major_feature_row_idx:
            data_class = ws.cell(row=row_idx, column=1).value
            if data_class in major_feature_names:
                ws.cell(row=row_idx, column=1).font = Font(color="FF0000", bold=True, size=14)
            ws.cell(row=row_idx, column=1).alignment = Alignment(horizontal="left", vertical="center")

        # 分类相关单元格左对齐、垂直居中
        for row_idx in range(2, ws.max_row + 1):
            for col_idx in (1, 2, 3):
                ws.cell(row=row_idx, column=col_idx).alignment = Alignment(horizontal="left", vertical="center")

        ws.column_dimensions["A"].width = 26
        ws.column_dimensions["B"].width = 30
        ws.column_dimensions["C"].width = 6
        ws.column_dimensions["D"].width = 30
        ws.column_dimensions["E"].width = 16
        ws.column_dimensions["F"].width = 8
        ws.column_dimensions["G"].width = 8
        ws.column_dimensions["H"].width = 10
        ws.column_dimensions["I"].width = 20
        ws.column_dimensions["J"].width = 15
        # 基线列是参考列（数量/分母），不参与着色；从 base 之后的对比列开始
        _apply_threshold_coloring(ws, rollback_col=6, optimize_col=7, first_result_col=11)

    if not has_written_sheet:
        ws_default.title = _safe_sheet_title("DT_HCT", used_sheet_names)

    _ensure_output_dirs()
    xlsx_path = os.path.join(REPORTS_DIR, f"output_hct_{file_name_suffix}.xlsx")
    wb.save(xlsx_path)
    txt_path = os.path.join(REPORTS_DIR, f"output_hct_{file_name_suffix}.txt")
    all_ver_com = {}
    tag_lv1_ver_com = {}
    tag_all_ver_com = defaultdict(dict)

    if not task_order_for_rank:
        msg += "DT_HCT: 无有效任务可统计\n"
        return msg, xlsx_path, txt_path, all_ver_com, tag_lv1_ver_com, tag_all_ver_com

    ordered_overall = {tid: overall_metric_sums.get(tid, 0.0) for tid in task_order_for_rank}
    relation, score_map = _relation_and_score(ordered_overall)
    all_ver_com["ver"] = relation
    all_ver_com["score"] = score_map

    for tag_lv1, task_metric in tag_lv1_metric_sums.items():
        ordered = {tid: task_metric.get(tid, 0.0) for tid in task_order_for_rank}
        rel, score = _relation_and_score(ordered)
        tag_lv1_ver_com[tag_lv1] = {"ver": rel, "score": score}

    for domain, tag_map in tag_all_metric_sums.items():
        for tag_name, task_metric in tag_map.items():
            ordered = {tid: task_metric.get(tid, 0.0) for tid in task_order_for_rank}
            rel, score = _relation_and_score(ordered)
            tag_all_ver_com[domain][tag_name] = {"ver": rel, "score": score}

    msg = _append_hct_report(msg, txt_path, all_ver_com, tag_lv1_ver_com, tag_all_ver_com)
    return msg, xlsx_path, txt_path, all_ver_com, tag_lv1_ver_com, tag_all_ver_com


if __name__ == "__main__":
    import time
    ver_map = ["base", "A", "B", "C"]
    tasks_dict = {"tmp":[1193157, 1193159, 1197924]}
    file_name_suffix = time.strftime("%Y%m%d%H%M%S", time.localtime())  
    msg, xlsx_path, txt_path, all_ver_com, tag_lv1_ver_com, tag_all_ver_com = run_compare_hct(tasks_dict, ver_map, "", file_name_suffix)

