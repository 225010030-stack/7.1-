#!/usr/bin/env python3
"""AMER FPP 台账 upsert（增强版）

改造点：
1. LEDGER_FIELDS 明确固化，写出前对既有台账行做 normalize，只保留 LEDGER_FIELDS
   里的键并补空值，避免脏字段导致 DictWriter ValueError；同时 DictWriter 打开
   extrasaction='ignore'，双保险防止崩溃。
2. 唯一键 UNIQUE_KEY_FIELDS 常量化，除原 6 字段外加入 Center，避免不同成本中心
   相同 Payment Description 被合并。
3. --ledger 改为必传；同时暴露 --allow-empty-ledger 让"新建空台账"这种场景显式开启。
4. --dry-run 只打印不落盘，便于前端"预演一下会不会覆盖历史"。
5. 提供 validate_prefill_headers / validate_ledger_headers，方便被 API 层复用。
"""

from __future__ import annotations

import argparse
import csv
import sys
from datetime import datetime
from pathlib import Path
from typing import Iterable


LEDGER_FIELDS = [
    "期间",
    "国家",
    "主体",
    "供应商",
    "费用类型",
    "Payment Description",
    "含税总额",
    "交易币种",
    "可提交(Y/N)",
    "FPP单号",
    "提单日期",
    "提交人",
    "Center",
    "状态",
    "最后更新时间",
]

# 唯一键：原来 6 个 + Center，避免不同成本中心的同名描述互相覆盖
UNIQUE_KEY_FIELDS = [
    "期间",
    "国家",
    "主体",
    "供应商",
    "费用类型",
    "Payment Description",
    "Center",
]

# 预填表最少要有这些列头（判定"这份文件是预填表而不是台账"）
PREFILL_REQUIRED_HEADERS = [
    "期间",
    "国家",
    "主体",
    "供应商",
    "费用类型",
    "Payment Description",
    "FPP场景",
    "Center",
]

# 台账文件最少要有这些列头（判定"这份文件是台账而不是预填表"）
LEDGER_REQUIRED_HEADERS = [
    "FPP单号",
    "状态",
    "最后更新时间",
]


def _clean(s) -> str:
    return (str(s) if s is not None else "").strip()


def key_of(row: dict) -> str:
    return "|".join(_clean(row.get(f)) for f in UNIQUE_KEY_FIELDS)


def normalize_row(row: dict) -> dict:
    """把任意来源行规范成 LEDGER_FIELDS 顺序、只含合法键的字典。"""
    return {f: _clean(row.get(f)) for f in LEDGER_FIELDS}


def _read_csv_headers(path: Path) -> list[str]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        try:
            return [h.strip() for h in next(reader)]
        except StopIteration:
            return []


def validate_prefill_headers(headers: Iterable[str]) -> list[str]:
    hs = {h.strip() for h in headers}
    return [c for c in PREFILL_REQUIRED_HEADERS if c not in hs]


def validate_ledger_headers(headers: Iterable[str]) -> list[str]:
    hs = {h.strip() for h in headers}
    return [c for c in LEDGER_REQUIRED_HEADERS if c not in hs]


def build_ledger_row(source: dict, now_iso: str) -> dict:
    fpp_no = _clean(source.get("FPP单号"))
    submitted = bool(fpp_no) and fpp_no != "待填写"
    return normalize_row({
        "期间": source.get("期间"),
        "国家": source.get("国家"),
        "主体": source.get("主体"),
        "供应商": source.get("供应商"),
        "费用类型": source.get("费用类型"),
        "Payment Description": source.get("Payment Description"),
        "含税总额": source.get("含税总额"),
        "交易币种": source.get("交易币种"),
        "可提交(Y/N)": source.get("可提交(Y/N)"),
        "FPP单号": fpp_no,
        "提单日期": source.get("提单日期"),
        "提交人": source.get("提交人"),
        "Center": source.get("Center"),
        "状态": "已提交" if submitted else "待提交",
        "最后更新时间": now_iso,
    })


def upsert(
    prefill_rows: list[dict],
    existing_rows: list[dict],
    now_iso: str | None = None,
) -> tuple[list[dict], int, int]:
    now_iso = now_iso or datetime.now().isoformat(timespec="seconds")
    existing_map: dict[str, dict] = {}
    for r in existing_rows:
        n = normalize_row(r)
        k = key_of(n)
        if k.strip("|"):
            existing_map[k] = n

    created = 0
    updated = 0
    for s in prefill_rows:
        row = build_ledger_row(s, now_iso)
        k = key_of(row)
        if not k.strip("|"):
            continue
        if k in existing_map:
            existing_map[k].update(row)
            updated += 1
        else:
            existing_map[k] = row
            created += 1

    rows = list(existing_map.values())
    rows.sort(key=lambda r: (
        r.get("期间", ""),
        r.get("国家", ""),
        r.get("供应商", ""),
        r.get("Center", ""),
        r.get("Payment Description", ""),
    ))
    return rows, created, updated


def main() -> int:
    parser = argparse.ArgumentParser(description="Upsert AMER ledger from prefill CSV")
    parser.add_argument("--prefill", required=True, help="提单预填表 CSV path")
    parser.add_argument("--ledger", required=True,
                        help="台账 CSV path（历史台账；不存在需显式加 --allow-empty-ledger）")
    parser.add_argument("--output", default=None,
                        help="输出台账 CSV path，默认写回 --ledger")
    parser.add_argument("--allow-empty-ledger", action="store_true",
                        help="允许在台账文件不存在或空表头时继续（新建初始台账）")
    parser.add_argument("--dry-run", action="store_true",
                        help="只打印统计，不写文件")
    parser.add_argument("--strict-headers", action="store_true",
                        help="严格校验：prefill 缺 FPP场景/Center 等或 ledger 缺 FPP单号/状态 等即报错")
    args = parser.parse_args()

    prefill = Path(args.prefill).resolve()
    ledger = Path(args.ledger).resolve()
    out = Path(args.output).resolve() if args.output else ledger

    if not prefill.exists():
        print(f"[ERROR] Prefill file not found: {prefill}")
        return 2

    # 1) prefill 表头校验
    prefill_headers = _read_csv_headers(prefill)
    missing_prefill = validate_prefill_headers(prefill_headers)
    if missing_prefill and args.strict_headers:
        print(f"[ERROR] Prefill 表头缺少字段: {missing_prefill}；请检查是否传反文件。")
        return 3
    elif missing_prefill:
        print(f"[WARN] Prefill 表头缺少字段: {missing_prefill}")

    with prefill.open("r", encoding="utf-8-sig", newline="") as f:
        source_rows = list(csv.DictReader(f))

    # 2) ledger 校验
    existing_rows: list[dict] = []
    if ledger.exists() and ledger.stat().st_size > 0:
        ledger_headers = _read_csv_headers(ledger)
        missing_ledger = validate_ledger_headers(ledger_headers)
        # 台账本身如果连 FPP单号/状态 都没有，八成传反了
        if missing_ledger and args.strict_headers:
            print(f"[ERROR] Ledger 表头缺少字段: {missing_ledger}；请检查是否传反了预填表。")
            return 4
        elif missing_ledger:
            print(f"[WARN] Ledger 表头缺少字段: {missing_ledger}")
        with ledger.open("r", encoding="utf-8-sig", newline="") as f:
            existing_rows = list(csv.DictReader(f))
    else:
        if not args.allow_empty_ledger:
            print(f"[ERROR] Ledger file not found or empty: {ledger}. "
                  f"如果是首次建台账请加 --allow-empty-ledger。")
            return 5

    rows, created, updated = upsert(source_rows, existing_rows)

    if args.dry_run:
        print(f"[DRY-RUN] would write to {out}")
        print(f"created={created}, updated={updated}, total={len(rows)}")
        return 0

    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=LEDGER_FIELDS, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)

    print(f"Ledger updated: {out}")
    print(f"created={created}, updated={updated}, total={len(rows)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
