#!/usr/bin/env python3
"""配置化的 xlsx -> CSV 提取层（零第三方依赖，绕开 openpyxl 的 stylesheet 崩溃）。

用法：
  from ledger_extract import inspect_xlsx, extract_xlsx

  inspect_xlsx(xlsx_path, kind="ledger")  ->  预览信息 dict（供前端确认映射）
  extract_xlsx(xlsx_path, kind="ledger")  ->  写出 CSV 并返回其路径

kind 取值：
  "ledger"     台账（映射规则见 ledger_config.json 的 ledger 段）
  "headcount"  P1 人数（映射规则见 ledger_config.json 的 headcount 段）

映射规则外置在 ledger_config.json，改表头/列名/目标 sheet 无需改代码。
"""
import csv
import json
import re
import tempfile
import zipfile
from pathlib import Path

CONFIG_PATH = Path(__file__).parent / "ledger_config.json"


# --------------------------------------------------------------------------- #
# 底层：直接解析 xlsx 的 XML（不使用 openpyxl，避免 stylesheet 解析崩溃）
# --------------------------------------------------------------------------- #
def _col_idx(ref: str) -> int:
    s = re.match(r"([A-Z]+)", ref).group(1)
    v = 0
    for ch in s:
        v = v * 26 + (ord(ch) - 64)
    return v


def _read_shared_strings(zf: zipfile.ZipFile):
    names = [n for n in zf.namelist() if n.endswith("sharedStrings.xml")]
    if not names:
        return []
    data = zf.read(names[0]).decode("utf-8", "replace")
    out = []
    for si in re.findall(r"<si>(.*?)</si>", data, re.S):
        out.append("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)))
    return out


def _list_sheets(zf: zipfile.ZipFile):
    """返回 [(sheet_name, sheet_xml_path), ...]"""
    wb = zf.read("xl/workbook.xml").decode("utf-8", "replace")
    sheets = re.findall(
        r'<sheet name="([^"]*)" sheetId="(\d+)"(?: state="(\w+)")? r:id="([^"]*)"',
        wb,
    )
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
    relmap = dict(re.findall(r'Id="([^"]*)"[^>]*Target="([^"]*)"', rels))
    out = []
    for name, _sid, _state, rid in sheets:
        t = relmap.get(rid, "")
        if t.startswith("/"):
            t = t[1:]
        elif not t.startswith("xl/"):
            t = "xl/" + t
        t = t.replace("./", "")
        out.append((name, t))
    return out


def _read_sheet_rows(zf: zipfile.ZipFile, sheet_path: str, shared: list):
    """读取一张表，返回 list[dict[col_idx -> value]]（col_idx 从 1 开始）。"""
    data = zf.read(sheet_path).decode("utf-8", "replace")
    rows = []
    for rb in re.findall(r"<row[^>]*>(.*?)</row>", data, re.S):
        row = {}
        for c in re.findall(r"<c\b[^>]*>.*?</c>|<c\b[^>]*/>", rb, re.S):
            rm = re.search(r'r="([A-Z]+\d+)"', c)
            if not rm:
                continue
            ref = rm.group(1)
            tm = re.search(r'\bt="([^"]*)"', c)
            ttype = tm.group(1) if tm else None
            val = None
            vm = re.search(r"<v>(.*?)</v>", c, re.S)
            if vm:
                raw = vm.group(1)
                if ttype == "s":
                    try:
                        val = shared[int(raw)]
                    except (ValueError, IndexError):
                        val = raw
                elif ttype in ("str", "inlineStr"):
                    val = raw
                else:
                    val = raw
            else:
                im = re.search(r"<is>(.*?)</is>", c, re.S)
                if im:
                    val = "".join(re.findall(r"<t[^>]*>(.*?)</t>", im.group(1), re.S))
            row[_col_idx(ref)] = val
        rows.append(row)
    return rows


# --------------------------------------------------------------------------- #
# 配置
# --------------------------------------------------------------------------- #
def _load_config() -> dict:
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))


def _match_sheet(sheet_names: list, target) -> str | None:
    """按配置的目标 sheet 名匹配；None 取第一张；找不到则尝试大小写/包含匹配。"""
    if not sheet_names:
        return None
    if target is None:
        return sheet_names[0]
    if target in sheet_names:
        return target
    tl = target.lower()
    for n in sheet_names:
        if n.lower() == tl:
            return n
    for n in sheet_names:
        if tl in n.lower():
            return n
    return None


def _apply_derived(value: str, rule: dict) -> str:
    v = (value or "").strip()
    for r in rule.get("rules", []):
        needle = str(r.get("contains", "")).upper()
        if needle and needle in v.upper():
            return r.get("value", "")
    return rule.get("default", "")


# --------------------------------------------------------------------------- #
# 对外接口
# --------------------------------------------------------------------------- #
def inspect_xlsx(xlsx_path, kind: str = "ledger") -> dict:
    """返回预览信息（供前端展示映射确认）。"""
    cfg = _load_config()[kind]
    with zipfile.ZipFile(xlsx_path) as zf:
        shared = _read_shared_strings(zf)
        sheets = _list_sheets(zf)
        sheet_names = [n for n, _ in sheets]
        target = _match_sheet(sheet_names, cfg.get("sheet"))
        target_path = dict(sheets).get(target) if target else None
        rows = _read_sheet_rows(zf, target_path, shared) if target_path else []

    header_row = int(cfg.get("header_row", 1))
    data_start = int(cfg.get("data_start_row", 2))
    header = rows[header_row - 1] if 0 <= header_row - 1 < len(rows) else {}
    headers = [header.get(i, "") for i in range(1, (max(header) if header else 0) + 1)]
    headers_nonempty = [h for h in headers if h]

    mapping = {}
    for sys_col, src in cfg.get("columns", {}).items():
        mapping[sys_col] = src if src in headers_nonempty else None
    missing = [sc for sc, m in mapping.items() if m is None]

    derived = {sc: rule.get("from") for sc, rule in cfg.get("derived", {}).items()}

    # 取前 2 条数据行作为样例
    sample = []
    for r in rows[data_start - 1: data_start + 1]:
        rowvals = {headers[i - 1]: r.get(i, "") for i in range(1, len(headers) + 1)}
        sample.append({h: rowvals.get(h, "") for h in headers_nonempty})

    return {
        "kind": kind,
        "sheets": sheet_names,
        "target_sheet": target,
        "headers": headers_nonempty,
        "mapping": mapping,
        "missing": missing,
        "derived": derived,
        "data_rows": max(0, len(rows) - data_start + 1),
        "sample": sample,
    }


def extract_xlsx(xlsx_path, kind: str = "ledger", out_csv_path: str = None) -> str:
    """按配置把 xlsx 提取为系统所需 CSV，返回 CSV 路径。"""
    cfg = _load_config()[kind]
    with zipfile.ZipFile(xlsx_path) as zf:
        shared = _read_shared_strings(zf)
        sheets = _list_sheets(zf)
        target = _match_sheet([n for n, _ in sheets], cfg.get("sheet"))
        target_path = dict(sheets).get(target) if target else None
        rows = _read_sheet_rows(zf, target_path, shared) if target_path else []

    header_row = int(cfg.get("header_row", 1))
    data_start = int(cfg.get("data_start_row", 2))
    header = rows[header_row - 1] if 0 <= header_row - 1 < len(rows) else {}
    headers = [header.get(i, "") for i in range(1, (max(header) if header else 0) + 1)]
    src_idx = {h: i for i, h in enumerate(headers, 1)}

    out_cols = list(cfg.get("columns", {}).keys()) + list(cfg.get("derived", {}).keys())

    if out_csv_path is None:
        out_csv_path = tempfile.NamedTemporaryFile(
            delete=False, suffix=".csv", dir=str(Path(xlsx_path).parent)
        ).name

    with open(out_csv_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=out_cols)
        w.writeheader()
        for r in rows[data_start - 1:]:
            rowvals = {headers[i - 1]: r.get(i, "") for i in range(1, len(headers) + 1)}
            out = {}
            for sc, src in cfg.get("columns", {}).items():
                idx = src_idx.get(src)
                out[sc] = (r.get(idx, "") or "").strip() if idx else ""
            for sc, rule in cfg.get("derived", {}).items():
                src_col = rule.get("from")
                idx = src_idx.get(src_col)
                out[sc] = _apply_derived(r.get(idx, "") if idx else "", rule)
            w.writerow(out)
    return out_csv_path


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        print("usage: ledger_extract.py <xlsx> <ledger|headcount>")
        sys.exit(1)
    info = inspect_xlsx(sys.argv[1], sys.argv[2])
    print(json.dumps(info, ensure_ascii=False, indent=2))
