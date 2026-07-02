# SSC 分摊工作台（no-ioa）· 可迁移代码库

把「台账」按**主体 join × 人数占比**自动分摊到账成本中心：后端 FastAPI + 单页前端。
含 xlsx 自动提取、数据质量校验，以及一个**按国家自助式特殊情形处理规范（Playbook）**，
让同事在页面里点开每条报错，按 **报错 → 检查错误 → 修改 → 上传** 自行修正后重传。

本目录即可作为独立仓库迁出到「总台」工程。

---

## 目录结构

| 文件 | 作用 |
|---|---|
| `app.py` | FastAPI 入口，挂载所有接口与静态页 |
| `allocate_core.py` | 分摊核心（纯函数，CLI 与流式接口共用，避免逻辑漂移） |
| `allocate_stream.py` | SSE 流式分摊接口 `/api/pain/allocate-stream` |
| `allocate.py` | CLI 版分摊（便于本地/批处理） |
| `ledger_extract.py` | xlsx 自动识别并提取为系统所需 CSV |
| `ledger_config.json` | 提取映射配置（台账/人数各自的源列→系统列） |
| `special_cases.json` | **特殊情形规范（单一数据源）**，前端 Playbook 与后端报错共用 |
| `no-ioa.html` | 单页前端：区域工具 + 知识库(规范中心) + 身份门禁 |
| `人数模板_按主体.csv` | 人数表模板（含「主体」列） |
| `分摊工具-特殊情形处理规范.md` | 规范说明（人读版，与 JSON 同源） |
| `ssc-backend.service` | systemd 服务样例（生产自愈托管） |
| `requirements.txt` | Python 依赖 |
| `cleaner.py` / `update_ledger.py` | 同工程内其他 SSC 工具（EE Listing 清洗/台账更新），一并随库迁出 |

---

## 本地运行

```bash
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
python -m uvicorn app:app --host 0.0.0.0 --port 8081
# 浏览器打开 http://localhost:8081/no-ioa.html
```

### 环境变量（均可选）

- `ALLOC_UPLOAD_DIR`：上传目录，默认 `/opt/ssc-noioa/uploads`
- `ALLOC_DOWNLOAD_DIR`：结果输出目录，默认 `/opt/ssc-noioa/downloads`

---

## 部署到总台（迁移 / 给别人迁入）

1. 把本目录整体迁入总台工程（代码 + 静态页 `no-ioa.html` 一起）。
2. 后端用任意 ASGI 托管（FastAPI/uvicorn 即可），前端 `no-ioa.html` 作为静态资源。
3. 自托管进程管理：参考 `ssc-backend.service` 作为 systemd 样例；
   迁到总台后改用总台统一的进程/容器管理即可。
4. **SSO / HR 数仓（当前未接入，仅建议）**：
   - SSO 建议在总台**网关层**统一做，本工具内部不做登录；
   - HR 人数数据建议由总台通过 **API** 提供给本工具，保留「上传文件」作为兜底。

---

## 特殊情形处理规范（Playbook）

- `GET /api/special-cases` 返回 `special_cases.json`。
- 前端「知识库 → 分摊规范中心」按国家（🌐全球 / 🇺🇸美国 / 🇨🇦加拿大 / 🇸🇬新加坡）渲染。
- 后端 `allocate_core` 报出的 `special_cases` **代码**与 JSON 的 `code` 一一对应；
  运行报错时，结果页带「查看处理指引」按钮，一键跳到对应条目。

### 报错 ↔ 文档 配对

结果页出现的每条报错（如 `S3 台账行缺主体`），点击「查看处理指引」→ 跳转规范中心 →
展开即看到该编号的 **①报错 ②检查错误 ③修改 ④上传** 四步，同事照着改完重传即可。
