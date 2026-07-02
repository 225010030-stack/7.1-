import csv
import io
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles

from cleaner import clean_ee_listing
from update_ledger import (
    LEDGER_FIELDS,
    PREFILL_REQUIRED_HEADERS,
    LEDGER_REQUIRED_HEADERS,
    validate_prefill_headers,
    validate_ledger_headers,
)
from allocate_stream import router as allocate_stream_router

BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

MONTH_ABBR_PATTERN = re.compile(
    r"^([A-Za-z]{3,4})[_\-\s]*(?:Active|Terminated|EElisting|eelisting|Listing|EE)",
    re.IGNORECASE,
)
DIGIT_PREFIX_PATTERN = re.compile(r"^(\d{6})")


def extract_month_prefix(filename: str) -> str:
    if not filename:
        return ""
    stem = Path(filename).stem
    m = MONTH_ABBR_PATTERN.match(stem)
    if m:
        return m.group(1).upper()
    m = DIGIT_PREFIX_PATTERN.match(stem)
    if m:
        return m.group(1)
    return ""


def make_output_filename(active_filename: str) -> str:
    prefix = extract_month_prefix(active_filename)
    if not prefix:
        prefix = datetime.now().strftime("%b").upper()
    return f"{prefix}_eelisting.xlsx"


app = FastAPI(title="SSC no-ioa + EE Listing API", version="1.3.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)
app.include_router(allocate_stream_router)


@app.get("/health")
async def health():
    return {"status": "ok", "service": "ssc-noioa", "version": "1.3.0"}


@app.get("/api/special-cases")
async def api_special_cases():
    """返回分摊特殊情形处理规范（单一数据源），供前端「规范中心」渲染。

    每个 case 含 code / title / level / applies_to / error / check / fix / upload，
    代码与后端 allocate_core 报告的 special_cases 一一对应。
    """
    sc_path = BASE_DIR / "special_cases.json"
    if not sc_path.exists():
        raise HTTPException(status_code=404, detail="special_cases.json 不存在")
    import json
    try:
        return json.loads(sc_path.read_text(encoding="utf-8"))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"读取规范文件失败: {e}")


@app.post("/process-sg-medical")
async def process_sg_medical(
    active: UploadFile = File(...),
    terminated: UploadFile = File(...),
    last_ee: UploadFile = File(...),
):
    """处理 EE Listing 数据清洗"""
    for f in (active, terminated, last_ee):
        if not f.filename or not f.filename.lower().endswith(".xlsx"):
            raise HTTPException(status_code=400, detail=f"文件 {f.filename or '(未命名)'} 必须是 .xlsx")

    job_id = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    job_dir = UPLOAD_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    output_filename = make_output_filename(active.filename)
    paths = {
        "active": job_dir / f"active_{Path(active.filename).name}",
        "terminated": job_dir / f"terminated_{Path(terminated.filename).name}",
        "last_ee": job_dir / f"last_ee_{Path(last_ee.filename).name}",
    }
    try:
        for src, dest in zip((active, terminated, last_ee), paths.values()):
            with open(dest, "wb") as f:
                shutil.copyfileobj(src.file, f)

        output_path = clean_ee_listing(
            ee_path=str(paths["last_ee"]),
            active_path=str(paths["active"]),
            terminated_path=str(paths["terminated"]),
            output_dir=str(OUTPUT_DIR),
            output_filename=output_filename,
        )
        output_name = os.path.basename(output_path)
        return {
            "status": "ok",
            "output_filename": output_name,
            "download_url": f"/download/{output_name}",
        }
    except FileNotFoundError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"处理失败: {e}")
    finally:
        try:
            shutil.rmtree(job_dir)
        except Exception:
            pass


@app.get("/download/{filename}")
async def download_file(filename: str):
    if "/" in filename or "\\" in filename or ".." in filename:
        raise HTTPException(status_code=400, detail="非法文件名")
    # Try OUTPUT_DIR first, then DOWNLOAD_DIR
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        file_path = BASE_DIR / "downloads" / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail=f"文件不存在: {filename}")
    ext = file_path.suffix.lower()
    if ext == ".csv":
        media = "text/csv; charset=utf-8"
    elif ext in (".xlsx", ".xlsm"):
        media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif ext == ".txt":
        media = "text/plain; charset=utf-8"
    else:
        media = "application/octet-stream"
    return FileResponse(str(file_path), media_type=media, filename=filename)


def _read_csv_first_row(path: Path) -> list[str]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        try:
            return [h.strip() for h in next(reader)]
        except StopIteration:
            return []


@app.post("/api/pain/update-ledger")
async def api_update_ledger(
    prefill_file: UploadFile = File(..., description="提单预填表 CSV（必传，含 FPP场景/Center 等）"),
    ledger_file: UploadFile = File(..., description="历史台账 CSV（必传，含 FPP单号/状态/最后更新时间 等）"),
    region: str = Form(default="US"),
    overwrite_confirmed: str = Form(default="false", description="是否已在前端确认会覆盖历史台账；不为 true 则拒绝"),
    dry_run: str = Form(default="false"),
):
    """AMER FPP 台账 upsert，公网后端版。

    - 服务端强制校验 prefill / ledger 表头，两者传反直接 400 返回。
    - ledger 是必传字段（前端最好也校验），避免误新建覆盖历史。
    - overwrite_confirmed 必须为 true 才会真正落盘；否则强制走 dry-run。
    """
    if not prefill_file.filename or not prefill_file.filename.lower().endswith((".csv", ".txt")):
        raise HTTPException(status_code=400, detail="预填表必须是 .csv 文件")
    if not ledger_file.filename or not ledger_file.filename.lower().endswith((".csv", ".txt")):
        raise HTTPException(status_code=400, detail="台账必须是 .csv 文件（必传）")

    job_id = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    region_tag = (region or "US").strip().upper() or "US"
    job_dir = UPLOAD_DIR / f"ledger_{job_id}"
    job_dir.mkdir(parents=True, exist_ok=True)

    prefill_path = job_dir / f"prefill_{Path(prefill_file.filename).name}"
    ledger_path = job_dir / f"ledger_{Path(ledger_file.filename).name}"
    output_name = f"台账_自动更新_{region_tag}_{job_id}.csv"
    output_path = OUTPUT_DIR / output_name

    try:
        with open(prefill_path, "wb") as f:
            shutil.copyfileobj(prefill_file.file, f)
        with open(ledger_path, "wb") as f:
            shutil.copyfileobj(ledger_file.file, f)

        # —— 服务端文件头校验，避免传反 ——
        prefill_headers = _read_csv_first_row(prefill_path)
        ledger_headers = _read_csv_first_row(ledger_path)
        missing_prefill = validate_prefill_headers(prefill_headers)
        missing_ledger = validate_ledger_headers(ledger_headers)

        errors = []
        if missing_prefill:
            errors.append({
                "file": "prefill",
                "filename": prefill_file.filename,
                "missing_headers": missing_prefill,
                "hint": "预填表必须含 FPP场景/Center 等字段；是不是把台账当预填传了？",
            })
        if missing_ledger:
            errors.append({
                "file": "ledger",
                "filename": ledger_file.filename,
                "missing_headers": missing_ledger,
                "hint": "台账必须含 FPP单号/状态/最后更新时间 等字段；是不是把预填当台账传了？",
            })
        if errors:
            raise HTTPException(status_code=400, detail={
                "message": "文件头校验失败，疑似传反或选错文件",
                "errors": errors,
                "expected_prefill_headers": PREFILL_REQUIRED_HEADERS,
                "expected_ledger_headers": LEDGER_REQUIRED_HEADERS,
            })

        # 把上传的 ledger 拷到 output_path 作为"写入基线"（--ledger 参数指向它）
        shutil.copyfile(ledger_path, output_path)

        is_dry = str(dry_run).lower() in ("1", "true", "yes")
        confirmed = str(overwrite_confirmed).lower() in ("1", "true", "yes")
        if not confirmed and not is_dry:
            # 未确认覆盖：强制走 dry-run 只算差异，不落盘覆盖 output_path
            is_dry = True

        cmd = [
            sys.executable,
            str(Path(__file__).parent / "update_ledger.py"),
            "--prefill", str(prefill_path),
            "--ledger", str(output_path),
            "--output", str(output_path),
            "--strict-headers",
        ]
        if is_dry:
            cmd.append("--dry-run")

        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        stdout = proc.stdout.strip()
        stderr = proc.stderr.strip()
        if proc.returncode != 0:
            raise HTTPException(status_code=400, detail={
                "message": "update_ledger.py 执行失败",
                "returncode": proc.returncode,
                "stdout": stdout,
                "stderr": stderr,
            })

        # 解析 created=x, updated=y, total=z
        m = re.search(r"created=(\d+),\s*updated=(\d+),\s*total=(\d+)", stdout)
        created = int(m.group(1)) if m else None
        updated = int(m.group(2)) if m else None
        total = int(m.group(3)) if m else None

        resp = {
            "status": "ok",
            "region": region_tag,
            "dry_run": is_dry,
            "confirmed": confirmed,
            "created": created,
            "updated": updated,
            "total": total,
            "log": stdout,
            "expected_prefill_headers": PREFILL_REQUIRED_HEADERS,
            "expected_ledger_headers": LEDGER_REQUIRED_HEADERS,
        }
        if not is_dry:
            resp["output_filename"] = output_name
            resp["download_url"] = f"/download/{output_name}"
        return resp

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"处理失败: {e}")
    finally:
        try:
            shutil.rmtree(job_dir)
        except Exception:
            pass


@app.post("/api/pain/allocate")
async def api_allocate(
    ledger_file: UploadFile = File(..., description="台账 CSV（含供应商/含税总额等）"),
    headcount_file: UploadFile = File(..., description="P1人数 CSV（成本中心/部门/人数）"),
    period: str = Form(...),
    region: str = Form(default="US"),
    dry_run: str = Form(default="false"),
):
    """AMER FPP 批量自动分摊：按人数占比将台账金额分摊到各成本中心。

    - ledger_file: 台账 CSV（US_P4_台账_YYYYMM_有数据.csv）
    - headcount_file: P1人数 CSV（US_P1_人数输入_YYYYMM.csv）
    - period: 账单月 YYYYMM，如 202605
    - region: US / CA，用于输出文件命名
    """
    if not ledger_file.filename or not ledger_file.filename.lower().endswith((".csv", ".txt", ".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="台账必须是 .csv 或 .xlsx 文件")
    if not headcount_file.filename or not headcount_file.filename.lower().endswith((".csv", ".txt", ".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="人数文件必须是 .csv 或 .xlsx 文件")
    if not period or len(period) != 6:
        raise HTTPException(status_code=400, detail="账单月必须是 YYYYMM 格式，如 202605")

    job_id = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    region_tag = (region or "US").strip().upper()
    job_dir = UPLOAD_DIR / f"alloc_{job_id}"
    job_dir.mkdir(parents=True, exist_ok=True)

    ledger_path = job_dir / f"ledger_{Path(ledger_file.filename).name}"
    headcount_path = job_dir / f"headcount_{Path(headcount_file.filename).name}"
    output_name = f"分摊结果_{region_tag}_{period}_{job_id}.csv"
    output_path = OUTPUT_DIR / output_name

    try:
        with open(ledger_path, "wb") as f:
            shutil.copyfileobj(ledger_file.file, f)
        with open(headcount_path, "wb") as f:
            shutil.copyfileobj(headcount_file.file, f)

        # 若上传的是 xlsx，先按配置提取为系统所需 CSV（内部完成）
        from ledger_extract import extract_xlsx
        if ledger_path.suffix.lower() in (".xlsx", ".xlsm"):
            ledger_path = Path(extract_xlsx(ledger_path, "ledger"))
        if headcount_path.suffix.lower() in (".xlsx", ".xlsm"):
            headcount_path = Path(extract_xlsx(headcount_path, "headcount"))

        # —— 上限检查：台账行数 × 人数行数 超过阈值则直接拒绝 ——
        MAX_OUTPUT_ROWS = 500_000
        try:
            with ledger_path.open("r", encoding="utf-8-sig", newline="") as f:
                ledger_count = sum(1 for _ in csv.reader(f)) - 1
            with headcount_path.open("r", encoding="utf-8-sig", newline="") as f:
                headcount_count = sum(1 for _ in csv.reader(f)) - 1
        except Exception:
            ledger_count = headcount_count = 0
        est_rows = max(ledger_count, 0) * max(headcount_count, 0)
        if est_rows > MAX_OUTPUT_ROWS:
            raise HTTPException(
                status_code=400,
                detail={
                    "message": "数据量过大，请拆分后分批处理",
                    "estimated_rows": est_rows,
                    "limit": MAX_OUTPUT_ROWS,
                    "hint": f"预估输出 {est_rows:,} 行，超过上限 {MAX_OUTPUT_ROWS:,} 行。请减少台账或成本中心数量后重试。",
                },
            )

        is_dry = str(dry_run).lower() in ("1", "true", "yes")

        cmd = [
            sys.executable,
            str(Path(__file__).parent / "allocate.py"),
            "--ledger", str(ledger_path),
            "--headcount", str(headcount_path),
            "--period", period,
            "--output", str(output_path),
        ]
        if is_dry:
            cmd.append("--dry-run")

        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        stdout = proc.stdout.strip()
        stderr = proc.stderr.strip()
        if proc.returncode != 0:
            raise HTTPException(status_code=400, detail={
                "message": "allocate.py 执行失败",
                "returncode": proc.returncode,
                "stdout": stdout,
                "stderr": stderr,
            })

        # 统计输出行数
        result_count = 0
        if output_path.exists():
            with output_path.open("r", encoding="utf-8-sig", newline="") as f:
                result_count = sum(1 for _ in f) - 1  # 减表头

        resp = {
            "status": "ok",
            "region": region_tag,
            "period": period,
            "dry_run": is_dry,
            "result_count": result_count,
            "log": stdout,
        }
        if not is_dry:
            resp["output_filename"] = output_name
            resp["download_url"] = f"/download/{output_name}"
        return resp

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"分摊处理失败: {e}")
    finally:
        try:
            shutil.rmtree(job_dir)
        except Exception:
            pass


@app.post("/api/pain/extract-preview")
async def api_extract_preview(
    xlsx_file: UploadFile = File(...),
    kind: str = Form("ledger"),
):
    """上传 xlsx，返回目标表 / 检测到的表头 / 建议映射，供前端确认后提交。"""
    if not xlsx_file.filename or not xlsx_file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="必须是 .xlsx 文件")
    if kind not in ("ledger", "headcount"):
        raise HTTPException(status_code=400, detail="kind 必须是 ledger 或 headcount")
    import tempfile
    from ledger_extract import inspect_xlsx
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=str(UPLOAD_DIR))
    try:
        tmp.write(await xlsx_file.read())
        tmp.close()
        info = inspect_xlsx(tmp.name, kind)
    finally:
        try:
            Path(tmp.name).unlink(missing_ok=True)
        except Exception:
            pass
    return info


@app.post("/api/pain/validate-headers")
async def api_validate_headers(
    prefill_file: UploadFile = File(...),
    ledger_file: UploadFile = File(...),
):
    """只做文件头校验，不做任何写入。前端在真正提交前用这个做 preflight。"""
    if not prefill_file.filename or not ledger_file.filename:
        raise HTTPException(status_code=400, detail="缺少 prefill 或 ledger 文件")
    try:
        prefill_bytes = await prefill_file.read()
        ledger_bytes = await ledger_file.read()
        prefill_headers = next(csv.reader(io.StringIO(prefill_bytes.decode("utf-8-sig", errors="ignore"))), [])
        ledger_headers = next(csv.reader(io.StringIO(ledger_bytes.decode("utf-8-sig", errors="ignore"))), [])
        missing_prefill = validate_prefill_headers(prefill_headers)
        missing_ledger = validate_ledger_headers(ledger_headers)
        return {
            "prefill": {
                "filename": prefill_file.filename,
                "headers": prefill_headers,
                "missing": missing_prefill,
                "ok": not missing_prefill,
            },
            "ledger": {
                "filename": ledger_file.filename,
                "headers": ledger_headers,
                "missing": missing_ledger,
                "ok": not missing_ledger,
            },
            "expected_prefill_headers": PREFILL_REQUIRED_HEADERS,
            "expected_ledger_headers": LEDGER_REQUIRED_HEADERS,
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"校验失败: {e}")


@app.get("/")
async def root_redirect():
    return RedirectResponse(url="/no-ioa.html")


# 最后挂 StaticFiles —— FastAPI 路由优先级：显式路由 > mount
# 所以 /health, /process-sg-medical, /download/{filename}, / 都不会被覆盖
app.mount("/", StaticFiles(directory=str(BASE_DIR), html=True), name="static")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app:app", host="0.0.0.0", port=8081, reload=False)
