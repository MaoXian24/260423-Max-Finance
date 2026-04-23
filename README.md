# Max Finance

A Python-based WRDS-powered U.S. equity data retrieval platform.

Max Finance is a WRDS-based financial analysis toolkit delivered in two interfaces:

1. Web App (Streamlit)
2. Desktop App (Tkinter)

Both versions provide an end-to-end workflow for company lookup, stock daily analysis, reported fundamentals visualization, derived metrics analysis, SIC industry benchmarking, and Excel export.

## 1. Project Scope

### 1.1 Data Sources
- WRDS / CRSP daily stock data: `crsp.dsf`, `crsp.stocknames`
- WRDS / Compustat fundamentals: `comp.funda`, `comp.company`

### 1.2 Analysis Coverage
- Stock daily metrics for one selected year (2015-2024)
- Reported fundamentals trends (2015-2024)
- Derived metrics trends (2015-2024)
- SIC industry benchmark trends (2015-2024)

### 1.3 Version Positioning
- Web version: presentation-friendly dashboard with browser-native download.
- Desktop version: standalone local GUI with responsive scroll layout and local file export.

## 2. Installation (Two Versions)

### 2.1 Common Requirements
- Python 3.10+
- Windows (batch launch scripts included)
- Valid WRDS account credentials
- Dependencies listed in [requirements.txt](requirements.txt)

Install dependencies once:

```bash
python -m pip install -r requirements.txt
```

You can also run [install_streamlit.bat](install_streamlit.bat) for a quick setup.

### 2.2 Web Version Installation and Launch
- Entry file: [streamlit_app.py](streamlit_app.py)
- Recommended launcher: [start_web_app.bat](start_web_app.bat)

Steps:
1. Ensure dependencies are installed.
2. Double-click [start_web_app.bat](start_web_app.bat), or run:

```bash
streamlit run streamlit_app.py
```

3. In the sidebar, enter WRDS user/password, ticker, and year.
4. Click `Run Analysis`.

### 2.3 Desktop Version Installation and Launch
- Entry file: [max_finance_desktop.py](max_finance_desktop.py)

Steps:
1. Ensure dependencies are installed.
2. Launch from terminal in project folder:

```bash
python max_finance_desktop.py
```

3. In the GUI, enter WRDS credentials, ticker, and year.
4. Click `Query`.

### 2.4 Public Deployment (Streamlit Community Cloud)

Use this path when you want a public URL that others can open directly.

Steps:
1. Push this project to a GitHub repository.
2. Open Streamlit Community Cloud and choose **Deploy a public app from GitHub**.
3. Select repository/branch and set the main file to [streamlit_app.py](streamlit_app.py).
4. In app settings, click **Advanced settings** and select Python 3.12.
5. In app settings, add WRDS credentials in **Secrets**:

```toml
WRDS_USER = "your_wrds_username"
WRDS_PASSWORD = "your_wrds_password"
```

6. Deploy and open the generated app URL.

Streamlit Community Cloud does not use `runtime.txt` to select the Python version here. The supported way is to choose Python 3.12 from **Advanced settings** during deployment.

Notes:
1. This app supports both secret keys above and nested format (`[wrds] user/password`).
2. Never commit real credentials to GitHub.
3. If secrets are configured, enable **Use Streamlit Secrets credentials** in sidebar for one-click login.

## 3. Operation Guide (Features + Export)

This section combines functional usage and export behavior in one workflow-oriented guide.

### 3.1 Standard Workflow
1. Provide WRDS credentials in the UI.
2. Input ticker and year.
3. Run query (`Run Analysis` in web, `Query` in desktop).
4. Review visual sections in order:
	- Stock Data Visualization
	- Financial Data Visualization
	- Derived Metrics Visualization
	- SIC Industry Visualization
	- All Data Tables
5. Export integrated Excel workbook by clicking `Download`.

### 3.2 Functional Modules

| Module | What It Shows | Time Window | Notes |
|---|---|---|---|
| Company Profile | Ticker, company name, SIC | Latest valid record | Uses `comp.funda`, fallback to `comp.company` |
| Stock Data | Close, Return, Volume, Market Cap | Selected year only | Built from `crsp.dsf`; market cap uses `shrout * 1000` |
| Reported Fundamentals | Revenue, Net Income, EBIT, Pretax Income, Assets, Liabilities, Equity | 2015-2024 | Annual Compustat fundamentals directly reported by firms |
| Derived Metrics | ROE (DuPont), Profit Margin (DuPont), Asset Turnover (DuPont), Equity Multiplier (DuPont), ROA, ROC, Debt Ratio, Debt-to-Equity, Equity Ratio, Capital Intensity, Liability-to-Revenue, EBIT Margin, Pretax Margin, Tax Burden, Interest Burden, EBIT to Assets | 2015-2024 | Computed from fundamentals with safe-division guards and visualized by metric selector |
| SIC Industry Benchmark | Avg sale/assets/equity + observation count | 2015-2024 | Aggregated by SIC and fiscal year |

### 3.3 Web vs Desktop Behavior Differences

| Topic | Web Version (Streamlit) | Desktop Version (Tkinter) |
|---|---|---|
| Result table preview | Stock table shows top 30 rows | Stock table shows top 10 rows |
| Download trigger | `st.download_button` in browser page | Local `Download` button in app window |
| Download save location | Browser-managed path (typically your browser default download folder unless overridden by browser settings) | Python process working directory (the folder from which `python max_finance_desktop.py` was started) |
| Download feedback | Browser download event and file save indicator | Non-blocking popup that auto-closes after about 3 seconds |
| Runtime responsiveness | Streamlit rerun/session-state model | Background thread + incremental stock chart updates by quarter |

### 3.4 Export Specification
- File name: `{TICKER}_{YEAR}_Full_Data.xlsx`
- Workbook sheets (when data exists):
	- `Company_SIC`
	- `Stock_Data`
	- `Finance_Data`
	- `Industry_Avg`

### 3.5 Operational Notes
1. Year selector affects stock module only.
2. Reported fundamentals, derived metrics, and industry modules always use 2015-2024.
3. If a chart is empty, the current ticker/metric may have no valid rows in WRDS for that scope.
4. WRDS authentication failures are surfaced in UI with concise error text.

## 4. Code Deep Dive (Implementation Details)

This section explains the internal implementation logic and highlights the design choices that make the software simple, stable, and easy to use.

### 4.1 Architecture Overview
- [streamlit_app.py](streamlit_app.py): web UI, WRDS querying, visualization, in-memory Excel generation, browser download.
- [max_finance_desktop.py](max_finance_desktop.py): desktop GUI, threaded query pipeline, chart refresh callbacks, local Excel export.
- Shared concept across both versions:
  - Query company profile first (to obtain SIC).
  - Query stock daily data for selected year.
	- Query reported fundamentals data and compute derived metrics for 2015-2024.
  - Query SIC benchmark using derived SIC.
  - Render charts/tables and export unified workbook.

### 4.2 Data Access and Query Design
1. WRDS connection handling:
	- Web version uses explicit PostgreSQL connection via `psycopg2` and WRDS connection constants.
	- Desktop version uses `wrds.Connection`.
2. SQL safety and consistency:
	- Core queries are parameterized for ticker/SIC values.
	- Date ranges and Compustat filters (`datafmt='STD'`, `consol='C'`, `indfmt='INDL'`) are applied to keep data consistent.
3. Company lookup fallback logic:
	- First attempt: latest qualified row from `comp.funda`.
	- Fallback: SIC from `comp.company` if no qualified fundamentals row exists.
4. Industry benchmark generation:
	- `GROUP BY fyear, SIC` with annual averages and observation counts.

### 4.3 Metric Construction Logic
1. Stock metrics:
	- `close = abs(prc)`
	- `daily_return = ret`
	- `volume = vol`
	- `market_cap = close * shrout * 1000`
2. Financial metrics:
	- Revenue, net income, EBIT, pretax income, assets, liabilities, equity come from Compustat annual fundamentals.
3. Derived metrics (computed at query time):
	- Profit Margin = `ni / sale`
	- Asset Turnover = `sale / at`
	- Equity Multiplier = `at / ceq`
	- ROE (DuPont) = `(ni/sale) * (sale/at) * (at/ceq)`
	- ROA = `ni / at`
	- ROC = `ni / (ceq + lt)`
	- Debt Ratio = `lt / at`
	- Debt-to-Equity = `lt / ceq`
	- Equity Ratio = `ceq / at`
	- Capital Intensity = `at / sale`
	- Liability-to-Revenue = `lt / sale`
	- EBIT Margin = `ebit / sale`
	- Pretax Margin = `pi / sale`
	- Tax Burden = `ni / pi`
	- Interest Burden = `pi / ebit`
	- EBIT to Assets = `ebit / at`
	- Division-by-zero is guarded by converting zero denominators to null before division.

### 4.4 Visualization and Readability Enhancements
1. Clear sectioned dashboard flow:
	- Stock -> Reported Fundamentals -> Derived Metrics -> Industry -> Tables.
2. Metric-level selectors:
	- Each module supports selector-driven chart focus to reduce clutter and improve interpretation.
3. Color semantics for faster cognition:
	- Stock: blue family
	- Reported fundamentals: amber/orange family
	- Derived metrics: green family
	- Industry: violet family
4. Table readability optimization (web):
	- Distinct header/body color themes by table category.
	- Alternating row background for scan efficiency.
5. Theme-aware plotting (web):
	- Text/grid color and palette adapt to Streamlit light/dark base theme.

### 4.5 UX and Reliability Highlights
1. Web credential safety and cache control:
	- Session-state fingerprint is used to detect credential changes and clear stale analysis results.
2. Fast failure diagnostics:
	- Authentication errors are normalized into short user-facing messages.
3. Desktop non-blocking query execution:
	- Query pipeline runs in a background thread.
	- Stock chart supports incremental refresh as each quarter is returned.
4. Download completion feedback (desktop):
	- Lightweight success popup closes automatically in approximately 3000 ms, reducing manual UI cleanup.
5. Graceful shutdown:
	- Desktop app uses a cooperative stop event and thread join during window close.

### 4.6 Export Implementation Details
1. Web export:
	- Data is written to an in-memory `BytesIO` workbook.
	- Download is served through Streamlit `st.download_button`.
	- Final storage path is decided by browser download settings.
2. Desktop export:
	- `pd.ExcelWriter` writes directly to a local filename.
	- Relative filename resolves under the process working directory.
	- For predictable location, launch from project folder before downloading.

## 5. Project Structure

- [README.md](README.md): project documentation
- [streamlit_app.py](streamlit_app.py): web app entry and logic
- [max_finance_desktop.py](max_finance_desktop.py): desktop app entry and logic
- [start_web_app.bat](start_web_app.bat): one-click web launcher
- [install_streamlit.bat](install_streamlit.bat): dependency install helper
- [requirements.txt](requirements.txt): Python package list
- [.streamlit/config.toml](.streamlit/config.toml): Streamlit runtime configuration
- [.streamlit/secrets.toml.example](.streamlit/secrets.toml.example): secrets template for cloud deployment
- [.gitignore](.gitignore): ignores local secrets and cache files

## 6. Suggested Demo Flow
1. Launch web app and complete full analysis path.
2. Launch desktop app and show equivalent modules.
3. Compare download behavior between browser-managed and local working-directory export.

---

<details>
<summary>中文说明</summary>

### 1. 项目范围

Max Finance 是一个基于 WRDS 的金融分析工具，提供两种交付界面：

1. 网页版（Streamlit）
2. 桌面版（Tkinter）

两个版本均支持从数据查询到图表展示再到 Excel 导出的完整流程，覆盖公司信息、股票日频数据、基础披露数据展示、派生指标分析与 SIC 行业基准分析。

#### 1.1 数据来源
- WRDS / CRSP 日频股票数据：`crsp.dsf`、`crsp.stocknames`
- WRDS / Compustat 基本面数据：`comp.funda`、`comp.company`

#### 1.2 分析范围
- 股票日频指标：针对单一选择年份（2015-2024）
- 基础披露数据趋势：2015-2024
- 派生分析指标趋势：2015-2024
- SIC 行业基准趋势：2015-2024

#### 1.3 双版本定位
- 网页版：适合展示与演示，图表看板化布局，下载由浏览器接管。
- 桌面版：适合本地独立运行，滚动式 GUI 布局，下载为本地文件直写。

### 2. 安装说明（双版本）

#### 2.1 通用环境要求
- Python 3.10+
- Windows（包含批处理脚本）
- 有效 WRDS 账号
- 依赖列表见 [requirements.txt](requirements.txt)

建议先统一安装依赖：

```bash
python -m pip install -r requirements.txt
```

也可使用 [install_streamlit.bat](install_streamlit.bat) 快速安装。

#### 2.2 网页版安装与启动
- 入口文件：[streamlit_app.py](streamlit_app.py)
- 推荐启动脚本：[start_web_app.bat](start_web_app.bat)

步骤：
1. 确认依赖安装完成。
2. 双击 [start_web_app.bat](start_web_app.bat)，或命令行执行：

```bash
streamlit run streamlit_app.py
```

3. 在侧边栏输入 WRDS 用户名、密码、Ticker、年份。
4. 点击 `Run Analysis` 执行查询。

#### 2.3 桌面版安装与启动
- 入口文件：[max_finance_desktop.py](max_finance_desktop.py)

步骤：
1. 确认依赖安装完成。
2. 在项目目录终端执行：

```bash
python max_finance_desktop.py
```

3. 在 GUI 输入 WRDS 凭证、Ticker、年份。
4. 点击 `Query` 执行查询。

#### 2.4 公网部署（Streamlit Community Cloud）

如果你希望他人通过链接直接访问，请使用该方式。

步骤：
1. 将本项目推送到 GitHub 仓库。
2. 在 Streamlit Community Cloud 选择 **Deploy a public app from GitHub**。
3. 选择仓库与分支，主入口文件设置为 [streamlit_app.py](streamlit_app.py)。
4. 在应用设置的 **Secrets** 中配置 WRDS 凭证：

```toml
WRDS_USER = "your_wrds_username"
WRDS_PASSWORD = "your_wrds_password"
```

5. 在 **Advanced settings** 中选择 Python 3.12。
6. 点击部署并访问生成的公网链接。

Streamlit Community Cloud 这里不读取 `runtime.txt` 来控制 Python 版本；部署时请在 **Advanced settings** 里手动选择 Python 3.12。

说明：
1. 本项目同时支持上述扁平写法和嵌套写法（`[wrds] user/password`）。
2. 请勿将真实账号密码提交到 GitHub。
3. 若已配置 Secrets，可在侧边栏启用 **Use Streamlit Secrets credentials** 实现免手动输入。

### 3. 操作指南（功能与导出整合）

本节将功能介绍与导出行为合并为一条可执行流程。

#### 3.1 标准操作流程
1. 在界面中填写 WRDS 账号信息。
2. 输入 Ticker 与年份。
3. 执行查询（网页版 `Run Analysis`，桌面版 `Query`）。
4. 按顺序查看五个模块：
	- Stock Data Visualization（股票）
	- Financial Data Visualization（财务）
	- Derived Metrics Visualization（派生指标）
	- SIC Industry Visualization（行业）
	- All Data Tables（数据表）
5. 点击 `Download` 导出整合工作簿。

#### 3.2 功能模块说明

| 模块 | 展示内容 | 时间范围 | 说明 |
|---|---|---|---|
| 公司信息 | Ticker、公司名称、SIC | 最新有效记录 | 优先使用 `comp.funda`，不足时回退 `comp.company` |
| 股票数据 | 收盘价、收益率、成交量、市值 | 仅所选年份 | 来源 `crsp.dsf`，市值计算含 `shrout * 1000` 转换 |
| 基础披露数据 | 营收、净利润、EBIT、税前利润、资产、负债、权益 | 2015-2024 | 基于 Compustat 年度披露原始字段 |
| 派生分析指标 | ROE (DuPont)、利润率 (DuPont)、资产周转率 (DuPont)、权益乘数 (DuPont)、ROA、ROC、资产负债率、负债权益比、权益比率、资本密集度、负债营收比、EBIT 利润率、税前利润率、税负系数、利息负担系数、EBIT to Assets | 2015-2024 | 由基础字段计算并进行安全除法处理后可视化 |
| SIC 行业基准 | 行业均值与样本量 | 2015-2024 | 按 SIC 与财年聚合 |

#### 3.3 网页版与桌面版差异

| 维度 | 网页版（Streamlit） | 桌面版（Tkinter） |
|---|---|---|
| 表格预览 | 股票表默认显示前 30 行 | 股票表默认显示前 10 行 |
| 下载触发 | 页面中的 `st.download_button` | 本地窗口中的 `Download` 按钮 |
| 下载保存位置 | 浏览器管理（通常为默认下载目录，最终以浏览器设置为准） | Python 进程当前工作目录（即启动命令所在目录） |
| 下载反馈 | 浏览器下载状态提示 | 非阻塞成功弹窗，约 3 秒自动关闭 |
| 运行响应机制 | Streamlit 的重跑与会话状态模型 | 后台线程 + 股票分季度增量刷新 |

#### 3.4 导出规范
- 文件名：`{TICKER}_{YEAR}_Full_Data.xlsx`
- 工作表（存在数据时写入）：
  - `Company_SIC`
  - `Stock_Data`
  - `Financial_DuPont`
  - `Industry_Avg`

#### 3.5 使用注意事项
1. 年份选择仅影响股票模块。
2. 基础披露数据、派生分析指标、行业模块固定使用 2015-2024。
3. 图表为空通常意味着当前 Ticker 与指标在对应范围内无有效数据。
4. WRDS 认证失败会在界面显示简洁错误信息。

### 4. 代码详解（实现逻辑与亮点）

本节对核心实现方式进行完整说明，覆盖架构、查询逻辑、指标构建、可视化设计、可用性与稳定性细节。

#### 4.1 架构总览
- [streamlit_app.py](streamlit_app.py)：网页端 UI、WRDS 查询、图表渲染、内存工作簿导出、浏览器下载。
- [max_finance_desktop.py](max_finance_desktop.py)：桌面端 GUI、线程化查询流水线、图表刷新回调、本地 Excel 导出。
- 两版本共享主流程：
  - 先查公司信息并提取 SIC。
  - 再查所选年份股票日频数据。
	- 再查 2015-2024 基础披露数据并计算派生指标。
  - 再查 SIC 行业基准。
  - 最后统一展示与导出。

#### 4.2 数据访问与查询设计
1. WRDS 连接策略：
	- 网页版通过 `psycopg2` 显式连接 WRDS PostgreSQL。
	- 桌面版通过 `wrds.Connection` 建立会话。
2. 查询一致性：
	- 关键查询使用参数化输入（如 ticker、SIC）。
	- 统一使用 Compustat 过滤条件 `datafmt='STD'`、`consol='C'`、`indfmt='INDL'`。
3. 公司信息回退机制：
	- 先从 `comp.funda` 获取最新有效记录。
	- 若无记录，则回退 `comp.company` 的 SIC。
4. 行业基准构造：
	- 以 `fyear + SIC` 分组计算均值和样本量，保证可横向比较。

#### 4.3 指标构建逻辑
1. 股票指标：
	- `close = abs(prc)`
	- `daily_return = ret`
	- `volume = vol`
	- `market_cap = close * shrout * 1000`
2. 财务指标：
	- 营收、净利润、EBIT、税前利润、资产、负债、权益来自 Compustat 年度基础字段。
3. 派生分析指标（查询后计算）：
	- Profit Margin = `ni / sale`
	- Asset Turnover = `sale / at`
	- Equity Multiplier = `at / ceq`
	- ROE (DuPont) = `(ni/sale) * (sale/at) * (at/ceq)`
	- ROA = `ni / at`
	- ROC = `ni / (ceq + lt)`
	- Debt Ratio = `lt / at`
	- Debt-to-Equity = `lt / ceq`
	- Equity Ratio = `ceq / at`
	- Capital Intensity = `at / sale`
	- Liability-to-Revenue = `lt / sale`
	- EBIT Margin = `ebit / sale`
	- Pretax Margin = `pi / sale`
	- Tax Burden = `ni / pi`
	- Interest Burden = `pi / ebit`
	- EBIT to Assets = `ebit / at`
	- 分母为 0 时统一转为空值后再计算，避免除零错误。

#### 4.4 可视化与可读性实现
1. 统一分区结构：股票 -> 基础披露数据 -> 派生分析指标 -> 行业 -> 表格，降低认知切换成本。
2. 指标下拉选择：每个模块支持单指标聚焦展示，减少信息拥挤。
3. 颜色语义化区分：
	- 股票：蓝色系
	- 基础披露数据：琥珀/橙色系
	- 派生分析指标：绿色系
	- 行业：紫色系
4. 网页版表格增强：
	- 按类别使用不同表头/表体配色。
	- 交替行底色提升扫描效率。
5. 网页版主题适配：图表文字、网格、配色会随浅色/深色主题自动调整。

#### 4.5 易用性与稳定性亮点
1. 网页版状态安全：
	- 使用凭证指纹检测登录信息变更，自动清除旧结果，避免状态串用。
2. 错误反馈友好：
	- 认证异常会格式化为短文本，便于快速定位问题。
3. 桌面版流畅交互：
	- 查询在线程中执行，不阻塞主界面。
	- 股票模块支持按季度增量刷新图表，用户无需等待全量完成。
4. 桌面版下载提示：
	- 下载成功后显示轻量弹窗，约 3000 ms 自动关闭，减少额外交互。
5. 安全退出机制：
	- 关闭窗口时通过 stop event 协作停止线程，并进行短时 join。

#### 4.6 导出实现细节
1. 网页版导出：
	- 使用 `BytesIO` 在内存生成 Excel。
	- 通过 `st.download_button` 触发下载。
	- 最终保存路径由浏览器下载设置决定。
2. 桌面版导出：
	- 使用 `pd.ExcelWriter` 直接写本地文件。
	- 相对路径会解析到当前工作目录。
	- 建议在项目目录启动程序以便统一管理导出文件。

### 5. 项目结构

- [README.md](README.md)：项目说明文档
- [streamlit_app.py](streamlit_app.py)：网页版主入口与核心逻辑
- [max_finance_desktop.py](max_finance_desktop.py)：桌面版主入口与核心逻辑
- [start_web_app.bat](start_web_app.bat)：网页版一键启动脚本
- [install_streamlit.bat](install_streamlit.bat)：依赖安装辅助脚本
- [requirements.txt](requirements.txt)：Python 依赖列表
- [.streamlit/config.toml](.streamlit/config.toml)：Streamlit 运行配置
- [.streamlit/secrets.toml.example](.streamlit/secrets.toml.example)：云部署 Secrets 示例模板
- [.gitignore](.gitignore)：忽略本地密钥与缓存文件

### 6. 演示建议流程
1. 先运行网页版，完整展示查询到导出的全流程。
2. 再运行桌面版，展示同等功能模块与交互方式。
3. 最后对比两版下载路径与下载反馈差异。

</details>