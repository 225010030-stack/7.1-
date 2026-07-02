#!/usr/bin/env python3
"""AMER FPP 批量自动分摊（Headcount-based Allocation，按主体 join）

业务逻辑：
  - 输入：台账 CSV（含主体/含税总额/供应商等）+ P1人数 CSV（主体/成本中心/人数）
  - 输出：分摊结果 CSV（每个成本中心按人数占比分摊金额，仅在同主体内分摊）
  - 核心逻辑在 allocate_core.allocate，本文件只负责 CLI 与落盘。

用法：
  python allocate.py --ledger <台账.csv> --headcount <人数.csv> --period 202605
  python allocate.py ... --dry-run        # 只打印统计，不写文件
"""
from __future__ import annotations

import argparse
import csv
import sys
from datetime import datetime
from pathlib import Path

from allocate_core import allocate, validate_headers, FIELDNAMES


def read_csv(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def _print_report(report: dict) -> None:
    print("\n=== 特殊情形处理清单 ===")
    if not report.get("special_cases"):
        print("  （无特殊情形）")
        return
    lvl_mark = {"error": "❌", "warn": "⚠️", "info": "ℹ️"}
    for s in report["special_cases"]:
        mark = lvl_mark.get(s["level"], "•")
        cnt = s["count"] if s["count"] else "-"
        print(f"  {mark} [{s['code']}] {s['title']}：{cnt} 条")
        print(f"        处理：{s['detail']}")
        print(f"        修复：{s['action']}")
    ua = report.get("unallocatable", {})
    if ua.get("missing_entity") or ua.get("no_cost_center"):
        print(f"  → 未分摊台账行：缺主体 {ua.get('missing_entity', 0)} 条 / 主体无成本中心 {ua.get('no_cost_center', 0)} 条")
    drift = report.get("rounding_drift", 0.0)
    if drift:
        print(f"  → 四舍五入累计误差：{drift:.2f} 元")


def main() -> int:
    parser = argparse.ArgumentParser(description="AMER FPP 批量自动分摊（按主体 join）")
    parser.add_argument("--ledger", required=True, help="台账 CSV path")
    parser.add_argument("--headcount", required=True, help="P1人数 CSV path")
    parser.add_argument("--period", required=True, help="账单月 YYYYMM，如 202605")
    parser.add_argument("--output", default=None, help="输出 CSV path")
    parser.add_argument("--dry-run", action="store_true", help="只打印统计，不写文件")
    args = parser.parse_args()

    ledger_path = Path(args.ledger).resolve()
    headcount_path = Path(args.headcount).resolve()
    period = args.period.strip()

    # 表头校验
    ok_l, err_l = validate_headers(ledger_path, "ledger")
    if not ok_l:
        print(f"[ERROR] 台账{err_l}")
        return 2
    ok_h, err_h = validate_headers(headcount_path, "headcount")
    if not ok_h:
        print(f"[ERROR] 人数{err_h}")
        return 2

    if not ledger_path.exists():
        print(f"[ERROR] 台账文件不存在: {ledger_path}")
        return 2
    if not headcount_path.exists():
        print(f"[ERROR] 人数文件不存在: {headcount_path}")
        return 2

    ledger_rows = read_csv(ledger_path)
    headcount_rows = read_csv(headcount_path)

    print(f"[INFO] 台账行数: {len(ledger_rows)}")
    print(f"[INFO] 成本中心行数: {len(headcount_rows)}")
    print(f"[INFO] 账单月: {period}")

    results, report = allocate(ledger_rows, headcount_rows, period)

    if report.get("error") == "no_headcount":
        print("[ERROR] 人数表中没有可用的成本中心数据（请检查主体/成本中心/人数列）")
        _print_report(report)
        return 1
    if not results:
        print("[WARN] 没有可分摊的结果，所有台账行均无法按主体匹配到成本中心")
        _print_report(report)
        return 1

    print(f"[INFO] 分摊结果行数: {len(results)}")
    _print_report(report)

    out_path = Path(args.output).resolve() if args.output else \
        ledger_path.parent / f"分摊结果_{period}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

    if args.dry_run:
        print(f"[DRY-RUN] 将写入 {out_path}（{len(results)} 行）")
        print("[DRY-RUN] 示例前3行:")
        for r in results[:3]:
            print("  ", r)
        return 0

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDNAMES)
        w.writeheader()
        w.writerows(results)

    print(f"[OK] 分摊结果已写入: {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
