#!/usr/bin/env python3
"""分摊核心逻辑（与前端 / CLI 无关，纯函数）。

被 allocate_stream.py（流式接口）与 allocate.py（CLI）共用，避免逻辑漂移。

核心规则：
  - 按「主体」join：台账的每一行只分摊到「同一主体」下的成本中心。
  - 人数按主体分组后，在组内按人数占比分摊，保证组内金额守恒。
  - 所有异常 / 特殊情形都结构化输出为 special_cases，供前端展示与同事核对。
"""
import csv
from pathlib import Path

from ledger_extract import inspect_xlsx

# 必需列（表头校验用）。部门 / 交易币种 / 供应商 / 费用类型 为选填，缺失时留空。
LEDGER_REQUIRED = ["主体", "含税总额", "Payment Description"]
HEADCOUNT_REQUIRED = ["主体", "成本中心", "人数"]

# 分摊结果输出列顺序
FIELDNAMES = [
    "成本中心", "部门", "人数", "人数占比",
    "供应商", "Payment Description",
    "国家", "主体", "费用类型", "交易币种",
    "含税总额(原)", "分摊金额(USD)", "账单月", "状态",
]


def _clean(s) -> str:
    return (str(s) if s is not None else "").strip()


def _entity_key(name: str) -> str:
    """把主体名称规范化为统一的 join key（大小写/别名容错）。

    规则：
      CANADA                         -> CA
      BRL / BRAZIL / BRASIL          -> BR
      AMERICA / US / UNITED STATES   -> US
      其它（如 SG、MX）              -> 原样大写
    空值                            -> ""（无法 join）
    """
    n = _clean(name).upper()
    if not n:
        return ""
    if "CANADA" in n:
        return "CA"
    if "BRL" in n or "BRAZIL" in n or "BRASIL" in n:
        return "BR"
    if "AMERICA" in n or n == "US" or "UNITED STATES" in n:
        return "US"
    return n


def validate_headers(path: Path, kind: str):
    """校验必需列是否存在。返回 (ok: bool, error_message: str)。"""
    required = LEDGER_REQUIRED if kind == "ledger" else HEADCOUNT_REQUIRED
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        try:
            info = inspect_xlsx(path, kind)
        except Exception as e:  # noqa: BLE001
            return False, f"无法解析 xlsx：{e}"
        missing = [m for m in info.get("missing", []) if m in required]
        if missing:
            return False, (
                f"表头缺失必需列：{', '.join(missing)}"
                f"（xlsx 工作表「{info.get('target_sheet') or '?'}」未找到这些列）"
            )
        return True, ""
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            header = next(csv.reader(f))
    except StopIteration:
        return False, "文件为空，没有任何表头行"
    header = [h.strip() for h in header]
    missing = [c for c in required if c not in header]
    if missing:
        return False, (
            f"表头缺失必需列：{', '.join(missing)}"
            f"（实际表头：{', '.join(header) or '（空）'}）"
        )
    return True, ""


def _build_special(code, title, count, level, detail, action):
    return {
        "code": code, "title": title, "count": count,
        "level": level, "detail": detail, "action": action,
    }


def allocate(ledger_rows, headcount_rows, period):
    """按主体 join 分摊。返回 (results, report)。

    report 结构：
      {
        "error": None | "no_headcount" | "no_match",
        "special_cases": [ {code,title,count,level,detail,action}, ... ],
        "warnings": [ 人类可读字符串 ],
        "unallocatable": {"missing_entity": int, "no_cost_center": int},
        "skipped_headcount": {"missing_entity": int, "missing_cc": int, "bad_n": int},
        "entities": { entity_key: {"raw":[...], "cost_centers":int, "headcount":int} },
        "rounding_drift": float,
      }
    """
    # —— 构建按主体分组的成本中心映射 ——
    hc_by_entity = {}   # entity_key -> {cc -> {"dept":..., "n":...}}
    hc_total = {}       # entity_key -> 该主体总人数
    entity_raw = {}     # entity_key -> set(原始主体名)
    dup_cc = 0
    dept_conflict = 0
    skip_no_entity_hc = 0
    skip_no_cc = 0
    skip_bad_n = 0

    for row in headcount_rows:
        er = _clean(row.get("主体"))
        ek = _entity_key(er)
        cc = _clean(row.get("成本中心"))
        dept = _clean(row.get("部门"))
        if not ek:
            skip_no_entity_hc += 1
            continue
        if not cc:
            skip_no_cc += 1
            continue
        try:
            n = int(float(_clean(row.get("人数"))))
        except (ValueError, TypeError):
            skip_bad_n += 1
            continue
        entity_raw.setdefault(ek, set()).add(er)
        grp = hc_by_entity.setdefault(ek, {})
        if cc in grp:
            # 多对一：同一主体下同一成本中心多次出现 -> 人数累加合并
            dup_cc += 1
            grp[cc]["n"] += n
            if dept and grp[cc]["dept"] and grp[cc]["dept"] != dept:
                dept_conflict += 1
            elif dept and not grp[cc]["dept"]:
                grp[cc]["dept"] = dept
        else:
            grp[cc] = {"dept": dept, "n": n}
        hc_total[ek] = hc_total.get(ek, 0) + n

    special = []
    if dup_cc:
        special.append(_build_special(
            "H4", "人数表重复成本中心", dup_cc, "warn",
            "同一主体下同一成本中心出现多次，已按人数累加合并",
            "请在人数源表去重，或确认合并无误"))
    if dept_conflict:
        special.append(_build_special(
            "H5", "成本中心部门名称冲突", dept_conflict, "warn",
            "同一成本中心在不同行部门名不一致，已采用首个",
            "请统一成本中心 -> 部门映射"))
    if skip_no_cc:
        special.append(_build_special(
            "H2", "人数行缺成本中心", skip_no_cc, "warn",
            "人数行未填成本中心，已跳过", "补全成本中心后重传"))
    if skip_bad_n:
        special.append(_build_special(
            "H3", "人数行人数无效", skip_bad_n, "warn",
            "人数非数值或缺失，已跳过", "补全合法人数"))
    if skip_no_entity_hc:
        special.append(_build_special(
            "H1", "人数行缺主体", skip_no_entity_hc, "warn",
            "人数行未填主体，无法参与按主体 join，已跳过", "补全主体后重传"))

    if not hc_by_entity:
        report = {
            "error": "no_headcount",
            "special_cases": special,
            "warnings": _warnings_from_special(special),
            "unallocatable": {"missing_entity": 0, "no_cost_center": 0},
            "skipped_headcount": {
                "missing_entity": skip_no_entity_hc, "missing_cc": skip_no_cc, "bad_n": skip_bad_n},
            "entities": {},
            "rounding_drift": 0.0,
        }
        return [], report

    # —— 台账重复预统计（自然键：主体+供应商+事由+金额+币种）——
    seen_keys = {}
    dup_ledger = 0
    for lrow in ledger_rows:
        ek = _entity_key(_clean(lrow.get("主体")))
        supplier = _clean(lrow.get("供应商", ""))
        desc = _clean(lrow.get("Payment Description", ""))
        currency = _clean(lrow.get("交易币种", ""))
        try:
            amt = float(_clean(lrow.get("含税总额")) or 0)
        except ValueError:
            amt = 0.0
        key = (ek, supplier, desc, round(amt, 2), currency)
        if key in seen_keys:
            dup_ledger += 1
        else:
            seen_keys[key] = 1

    # —— 分摊计算（按主体 join）——
    results = []
    bad_amt = 0
    neg_amt = 0
    miss_desc = 0
    unalloc_missing_entity = 0
    unalloc_no_match = 0
    drift = 0.0
    ledger_entity_keys = set()

    for lrow in ledger_rows:
        supplier = _clean(lrow.get("供应商", ""))
        desc = _clean(lrow.get("Payment Description", ""))
        country = _clean(lrow.get("国家", ""))
        entity_raw_val = _clean(lrow.get("主体"))
        ek = _entity_key(entity_raw_val)
        cost_type = _clean(lrow.get("费用类型", ""))
        currency = _clean(lrow.get("交易币种", ""))
        raw_amt = _clean(lrow.get("含税总额"))
        try:
            total_amt = float(raw_amt or 0)
        except ValueError:
            total_amt = 0.0
            bad_amt += 1
        else:
            if raw_amt == "":
                bad_amt += 1
            elif total_amt < 0:
                neg_amt += 1

        if not desc:
            miss_desc += 1

        if not ek:
            unalloc_missing_entity += 1
            continue
        ledger_entity_keys.add(ek)
        grp = hc_by_entity.get(ek)
        if not grp:
            unalloc_no_match += 1
            continue
        total_n = hc_total[ek]
        row_sum = 0.0
        for cc, info in grp.items():
            ratio = info["n"] / total_n
            allocated = round(total_amt * ratio, 2)
            row_sum += allocated
            results.append({
                "成本中心": cc, "部门": info["dept"], "人数": str(info["n"]),
                "人数占比": f"{ratio:.4f}", "供应商": supplier,
                "Payment Description": desc, "国家": country, "主体": entity_raw_val,
                "费用类型": cost_type, "交易币种": currency,
                "含税总额(原)": str(total_amt), "分摊金额(USD)": str(allocated),
                "账单月": period, "状态": "待提交",
            })
        drift += (row_sum - total_amt)

    drift = round(drift, 2)

    # —— 台账侧特殊情形 ——
    if bad_amt:
        special.append(_build_special(
            "S1", "含税总额缺失/非数值", bad_amt, "warn",
            "按 0 元分摊", "补全金额或删除该行"))
    if neg_amt:
        special.append(_build_special(
            "S2", "含税总额为负（退款/冲正）", neg_amt, "warn",
            "已按负数正常分摊", "确认是否为真实冲正"))
    if miss_desc:
        special.append(_build_special(
            "S5", "台账缺 Payment Description", miss_desc, "info",
            "输出事由为空", "补全付款事由"))
    if dup_ledger:
        special.append(_build_special(
            "S6", "台账疑似重复行", dup_ledger, "warn",
            "主体+供应商+事由+金额+币种 完全相同",
            "请同事核对源数据并去重（工具未自动去重）"))
    if unalloc_missing_entity:
        special.append(_build_special(
            "S3", "台账行缺主体", unalloc_missing_entity, "error",
            "无法按主体 join，该行未分摊（未计入结果）", "补全「我方主体」后重传"))
    if unalloc_no_match:
        special.append(_build_special(
            "S4", "主体在人数表无对应成本中心", unalloc_no_match, "error",
            "该主体在人数表中无成本中心，该行未分摊",
            "在人数表补充该主体成本中心，或核对主体名称"))
    unused = [ek for ek in hc_by_entity if ek not in ledger_entity_keys]
    if unused:
        special.append(_build_special(
            "H6", "人数表部分主体未被使用", len(unused), "info",
            "这些主体在台账中无对应行：" + ", ".join(sorted(unused)),
            "如属正常可忽略；否则核对台账主体"))
    if abs(drift) > 0.02:
        special.append(_build_special(
            "G3", "四舍五入累计误差", 0, "info",
            f"全部分摊行合计与台账原额相差 {drift:.2f} 元（正常取整误差）",
            "无需处理；如需精确可保留更多小数位"))

    report = {
        "error": None,
        "special_cases": special,
        "warnings": _warnings_from_special(special),
        "unallocatable": {
            "missing_entity": unalloc_missing_entity,
            "no_cost_center": unalloc_no_match,
        },
        "skipped_headcount": {
            "missing_entity": skip_no_entity_hc,
            "missing_cc": skip_no_cc,
            "bad_n": skip_bad_n,
        },
        "entities": {
            ek: {
                "raw": sorted(entity_raw.get(ek, [])),
                "cost_centers": len(hc_by_entity[ek]),
                "headcount": hc_total[ek],
            }
            for ek in hc_by_entity
        },
        "rounding_drift": drift,
    }
    return results, report


def _warnings_from_special(special):
    out = []
    for s in special:
        if s["count"] or s["code"] in ("G3",):
            out.append(f"[{s['level']}] {s['code']} {s['title']}：{s['count']} 条")
    return out
