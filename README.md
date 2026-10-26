# SheetPilot

通过自然语言完成表格处理的 AI 应用。上传 Excel 或 CSV，描述需求，确认执行计划，再预览、核验并下载结果。

本项目用于产品实践与作品展示，重点探索从需求澄清到可验证交付的完整表格处理流程。目前版本为 **2.1.0**，默认面向小规模试用，尚未按商业服务的可用性和并发要求建设。

## 产品流程

上传文件 → 分析表格结构 → 描述需求 → 澄清规则 → 确认计划 → 执行与核验 → 下载或恢复历史版本。

例如：

- “按客户编号匹配客户表，给订单表补充客户等级。”
- “把编号列两端的空格去掉，保留前导零。”
- “按部门汇总金额，并按金额从高到低排序。”
- “在总额列填入数量 × 单价。”

请求需要落在当前能力范围内；遇到来源不明确、重复键、覆盖已有内容等情况，应用会要求澄清或确认。

## 主要能力

| 能力 | 说明 |
| --- | --- |
| 匹配与关联 | 同一工作簿内，两个工作表之间的单列键精确匹配 |
| 标准化与清洗 | 文本规范化、类型转换、常量填充、去重和筛选 |
| 聚合与排序 | 分组统计、排序、排名和 Top N |
| 数据准备 | 数值填补、分类映射、尺寸解析和受控派生指标 |
| 公式处理 | 公式诊断、修复、明确列名算式写入与重算 |
| 文件转换 | CSV 转 XLSX、指定工作表导出 CSV |
| 版本与核验 | 确认计划后执行，核验通过才提交新版本，支持恢复历史版本 |

CSV 需先通过确认的转换计划转为 XLSX，才能执行其他表格处理；XLSX 转 CSV 可以作为复合计划的最后一步。

## 技术实现

- **网页**：React 19、TypeScript、Vite、Tailwind CSS。
- **后端**：Python 3.12+、FastAPI、Uvicorn。
- **处理与编排**：openpyxl、pandas、LangGraph、声明式能力契约。
- **模型**：通过服务端配置调用 OpenAI 兼容接口，默认配置面向 DeepSeek；密钥不发送给浏览器。
- **存储**：SQLite 与本地工作簿文件。
- **公式重算**：Windows 优先使用 Microsoft Excel，其他环境使用 LibreOffice；Docker 镜像安装 LibreOffice Calc。

前端构建后由同一个 FastAPI 服务提供，因此首版只需部署一个 Web 服务。

## 本地运行（Windows / PowerShell）

### 1. 准备环境

安装 Python 3.12+ 和 Node.js 24 LTS。以下命令均从仓库根目录开始执行；首次克隆后，目录通常名为 `SheetPilot`。

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e ".[test]"
Copy-Item .env.example .env
```

`Copy-Item` 仅用于首次创建配置；已有 `.env` 时保留现有文件。

### 2. 配置模型

编辑 `.env`，填入实际可用的接口地址、模型 ID 和 API Key：

```dotenv
SHEETPILOT_LLM_PROVIDER=deepseek
SHEETPILOT_API_KEY=your-api-key
SHEETPILOT_BASE_URL=https://api.deepseek.com
SHEETPILOT_MODEL=your-model-id
SHEETPILOT_SESSION_SECRET=replace-with-a-long-random-secret
```

`your-api-key`、`your-model-id` 和会话密钥都是占位值，不能直接使用。模型名称应以你的供应商账户实际可用的模型为准。可用以下命令生成会话密钥，再复制到 `.env`：

```powershell
.\.venv\Scripts\python.exe -c "import secrets; print(secrets.token_hex(32))"
```

其余文件大小、行数、留存时间和执行预算等参数见 [.env.example](.env.example)。没有模型配置时，自然语言规划功能不可用。

### 3. 构建网页并启动

```powershell
Set-Location web
npm ci
npm run build
Set-Location ..
.\.venv\Scripts\python.exe -m uvicorn sheetpilot.web_api:app --host 127.0.0.1 --port 8000 --workers 1
```

打开 <http://127.0.0.1:8000>，健康检查地址为 <http://127.0.0.1:8000/api/health>。

修改网页时，可在第二个终端进入 `web` 运行 `npm run dev`，通过 <http://localhost:5173> 访问；开发服务器会将 `/api` 转发到本地 8000 端口。后端仍须保持运行。

需要公式重算时，还需安装 Microsoft Excel 或 LibreOffice。缺少计算引擎时，相关任务会暂停，不能仅凭公式文本就提交计算结果。

## Docker 运行与云端部署

安装并启动 Docker，在仓库根目录准备好 `.env` 后运行：

```powershell
docker compose up --build -d
```

本地访问 <http://127.0.0.1:8000>。查看日志和停止服务：

```powershell
docker compose logs -f sheetpilot
docker compose down
```

Compose 使用命名数据卷保存 `/app/data_web`。不要在希望保留任务数据时使用 `docker compose down -v`。

部署到支持 Docker 的云平台时：

1. 使用仓库根目录的 `Dockerfile` 构建，服务监听端口为 **8000**。
2. 在平台配置模型变量和固定的随机会话密钥，设置 `SHEETPILOT_WEB_DATA_DIR=/app/data_web`。
3. 将持久化磁盘挂载到 **`/app/data_web`**，用于保存数据库、上传文件和结果版本。
4. 健康检查路径设置为 **`/api/health`**。
5. 保持 **单实例、单 worker**；当前本地文件、SQLite 和进程内任务执行器不支持多实例共享运行。
6. 通过平台提供的 HTTPS 地址检查完整的上传、规划、执行和下载流程，并验证重启后的数据保留。

当前仓库包含部署配置，云端部署仍需在实际平台上验收。公开试用前，应补充访问控制、请求频率及资源使用限制，并检查反向代理下的 HTTPS 与会话行为。匿名会话隔离不等于登录或使用额度控制。

## 数据与使用边界

- 默认上传上限为 20 MB，单工作表最多 100,000 行；支持 `.xlsx` 和 `.csv`。
- 默认向模型发送任务文字、表格结构和统计，不默认发送完整工作簿；特定异常值的发送需要额外授权。
- 网页数据保存在 `data_web/`，旧 Streamlit 入口使用 `data/`；任务默认闲置 7 天后清理，也可主动删除。
- 匿名会话逐请求校验任务所有权；更换浏览器或丢失会话信息后，不保证能重新访问原任务。
- 不支持宏、Power Query、密码保护、外部链接、复杂合并表头、多工作簿关联或任意代码执行。
- 单个复合任务最多五个业务步骤；公式结构与计算结果通过核验，不代表业务规则本身正确。
- 展示和测试请使用合成或脱敏数据。

## 项目结构

```text
SheetPilot/
├── src/sheetpilot/       # 后端、规划、处理、存储与核验
├── web/                 # 前端源码、字体与 npm 依赖清单
├── tests/               # 自动化测试
├── evals/               # 评估用例与运行器
├── scripts/             # 验证、统计导出等辅助脚本
├── docs/                # 历史说明与仓库发布指南
├── app.py               # 旧 Streamlit 入口
├── pyproject.toml       # Python 依赖与项目配置
├── Dockerfile           # 前后端统一构建镜像
├── docker-compose.yml   # 本地容器运行与数据卷配置
└── .env.example         # 环境变量示例，不含真实密钥
```

`app_chat.py` 和 `app_v2.py` 保留为旧入口；当前主要展示入口为网页。`data/`、`data_web/`、虚拟环境、依赖缓存、网页副本和评估运行产物保留在本地，不进入仓库。

## 验证与测试

在仓库根目录执行：

```powershell
.\.venv\Scripts\python.exe scripts/smoke_test.py
.\.venv\Scripts\python.exe -m pytest -q --basetemp=.pytest-tmp
```

真实公式引擎集成测试需先安装 Excel 或 LibreOffice：

```powershell
$env:SHEETPILOT_EXCEL_INTEGRATION = "1"
.\.venv\Scripts\python.exe -m pytest tests/test_excel_recalculation.py -q --basetemp=.pytest-tmp-excel
Remove-Item Env:SHEETPILOT_EXCEL_INTEGRATION
```

部分历史用例依赖仓库外的本地样本，缺少时会跳过；完整测试结果需结合运行环境判断。评估入口见 [evals/README.md](evals/README.md)。部分历史辅助脚本和报告引用原工作区的 `Test/`、`Guidance/`，这些研究资料不随应用仓库发布。

## 更多说明

- [创建 GitHub 仓库与首次推送](docs/GITHUB_PUBLISH.md)
- [历史版本与实现说明](docs/VERSION_NOTES.md)
