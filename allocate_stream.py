#!/usr/bin/env python3
"""带进度回报的分配接口 — StreamingResponse 版本

用法：
  POST /api/pain/allocate-stream
  返回：SSE 格式的进度更新流
    data: {"progress": 10, "message": "校验台账..."}
    ...
    data: {"progress": 100, "done": true, "download": "/download/xxx.csv",
           "rows": N, "warnings": [...], "report": {...}}

核心分摊逻辑下沉到 allocate_core，本文件只负责：上传 / 表头校验 / xlsx 提取 /
进度回报 / 落盘。
"""

import csv
import json
import os
import uuid
from datetime import datetime
from pathlib import Path

from fastapi import APIRouter, File, Form, UploadFile
from fastapi.responses import StreamingResponse

from ledger_extract import extract_xlsx
from allocate_core import (
    allocate,
    validate_headers,
    FIELDNAMES,
)

router = APIRouter(tags=["pain-allocate"])

# 输出行数上限：超过则拒绝处理，避免服务器被超大文件算死
MAX_OUTPUT_ROWS = 500_000

# 上传/下载目录可通过环境变量覆盖（本地测试用）；生产默认 /opt/ssc-noioa
UPLOAD_DIR = Path(os.environ.get("ALLOC_UPLOAD_DIR", "/opt/ssc-noioa/uploads"))
DOWNLOAD_DIR = Path(os.environ.get("ALLOC_DOWNLOAD_DIR", "/opt/ssc-noioa/downloads"))


def _coerce_to_csv(p: Path, kind: str) -> Path:
    """若上传的是 xlsx，先按配置提取为系统所需 CSV；否则原样返回。"""
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        return Path(extract_xlsx(p, kind))
    return p


def _prog(pct, msg):
    return f"data: {json.dumps({'progress': pct, 'message': msg}, ensure_ascii=False)}\n\n"


def _done(payload: dict):
    p = {"progress": 100, "done": True}
    p.update(payload)
    return f"data: {json.dumps(p, ensure_ascii=False)}\n\n"


def allocate_with_progress(ledger_path: Path, headcount_path: Path,
                           period: str, region: str):
    """生成器函数，逐步 yield 进度信息。"""
    # Step 0: 表头校验（CSV 读表头；xlsx 走 inspect_xlsx）
    ok_l, err_l = validate_headers(ledger_path, "ledger")
    if not ok_l:
        yield _prog(100, f"❌ 台账{err_l}")
        yield _done({"error": "bad_header", "message": "台账" + err_l})
        return
    ok_h, err_h = validate_headers(headcount_path, "headcount")
    if not ok_h:
        yield _prog(100, f"❌ 人数{err_h}")
        yield _done({"error": "bad_header", "message": "人数" + err_h})
        return

    # 若上传的是 xlsx，先按配置提取为系统所需 CSV（内部完成，前端无感）
    ledger_path = _coerce_to_csv(ledger_path, "ledger")
    headcount_path = _coerce_to_csv(headcount_path, "headcount")

    total_steps = 6
    step = 0

    # Step 1: 校验文件
    step += 1
    yield _prog(int(step / total_steps * 100), f"📂 校验输入文件 [{region}]")
    if not ledger_path.exists():
        yield _prog(100, "❌ 台账文件不存在")
        yield _done({"error": "missing_file", "message": "台账文件不存在"})
        return
    if not headcount_path.exists():
        yield _prog(100, "❌ 人数文件不存在")
        yield _done({"error": "missing_file", "message": "人数文件不存在"})
        return

    # Step 2: 读台账
    step += 1
    yield _prog(int(step / total_steps * 100), f"📊 读取台账数据...")
    with ledger_path.open("r", encoding="utf-8-sig", newline="") as f:
        ledger_rows = list(csv.DictReader(f))
    ledger_count = len(ledger_rows)

    # Step 3: 读人数
    step += 1
    yield _prog(int(step / total_steps * 100), f"👥 读取 P1 人数数据...")
    with headcount_path.open("r", encoding="utf-8-sig", newline="") as f:
        headcount_rows = list(csv.DictReader(f))
    headcount_count = len(headcount_rows)

    # Step 3.5: 上限检查（在构建映射前拦截，避免算死服务器）
    est_rows = ledger_count * headcount_count
    if est_rows > MAX_OUTPUT_ROWS:
        yield _prog(100, f"⚠️ 数据量过大，请拆分后分批处理｜预估输出 {est_rows:,} 行（上限 {MAX_OUTPUT_ROWS:,} 行）")
        yield _done({
            "error": "data_too_large", "rows": est_rows, "limit": MAX_OUTPUT_ROWS,
            "message": f"数据量过大，请拆分后分批处理（预估 {est_rows:,} 行，上限 {MAX_OUTPUT_ROWS:,} 行）",
        })
        return

    # Step 4: 分摊计算（移交核心模块，按主体 join）
    step += 1
    yield _prog(int(step / total_steps * 100), f"🔢 按主体分摊 ({ledger_count} 条台账 × {headcount_count} 条人数)")
    results, report = allocate(ledger_rows, headcount_rows, period)
    warnings = report.get("warnings", [])

    if report.get("error") == "no_headcount":
        yield _prog(100, "⚠️ 人数表中没有可用的成本中心数据")
        yield _done({
            "error": "no_headcount", "report": report, "warnings": warnings,
            "message": "人数表中没有可用的成本中心数据（请检查主体/成本中心/人数列）",
        })
        return

    if not results:
        yield _prog(100, "⚠️ 没有可分摊的结果（所有台账行均无法按主体匹配）")
        yield _done({
            "error": "no_match", "report": report, "warnings": warnings,
            "message": "所有台账行均无法按主体匹配到成本中心，请检查主体名称与人数表",
        })
        return

    if warnings:
        yield _prog(99, "⚠️ 特殊情形：" + "；".join(warnings))

    # Step 5: 写入结果文件
    step += 1
    result_count = len(results)
    yield _prog(int(step / total_steps * 100), f"💾 生成分摊结果 ({result_count} 行)...")

    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = f"allocate_{region}_{period}_{ts}.csv"
    out_path = DOWNLOAD_DIR / out_name

    with out_path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDNAMES)
        w.writeheader()
        w.writerows(results)

    # 完成！
    yield _prog(100, f"✅ 分摊完成！共 {result_count} 行 | {len(report['entities'])} 个主体")
    yield _done({
        "download": f"/download/{out_name}",
        "rows": result_count,
        "warnings": warnings,
        "report": report,
        "message": "✅ 完成",
    })


@router.post("/api/pain/allocate-stream")
async def api_allocate_stream(
    ledger_file: UploadFile = File(...),
    headcount_file: UploadFile = File(...),
    period: str = Form(...),
    region: str = Form(default="US"),
):
    """流式进度分配接口。"""
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

    # 保存上传文件（保留原扩展名，xlsx 才能被自动提取）
    ledger_ext = Path(ledger_file.filename).suffix or ".csv"
    headcount_ext = Path(headcount_file.filename).suffix or ".csv"
    ledger_path = UPLOAD_DIR / f"ledger_{uuid.uuid4().hex[:8]}{ledger_ext}"
    headcount_path = UPLOAD_DIR / f"hc_{uuid.uuid4().hex[:8]}{headcount_ext}"

    with ledger_path.open("wb") as f:
        f.write(await ledger_file.read())
    with headcount_path.open("wb") as f:
        f.write(await headcount_file.read())

    return StreamingResponse(
        allocate_with_progress(ledger_path, headcount_path, period, region),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )
