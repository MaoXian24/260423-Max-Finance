import io
import hashlib
import random
import warnings

import matplotlib.pyplot as plt
import pandas as pd
import psycopg2
import streamlit as st
from wrds.sql import WRDS_CONNECT_ARGS, WRDS_POSTGRES_DB, WRDS_POSTGRES_HOST, WRDS_POSTGRES_PORT

warnings.filterwarnings("ignore")

# Global constants used across queries, chart formatting, and defaults.
SHROUT_MULTIPLIER = 1000
YEAR_OPTIONS = [str(y) for y in range(2015, 2025)]
DEFAULT_TICKER_POOL = ["AAPL", "MSFT", "NVDA", "AMZN", "GOOGL", "META", "TSLA"]

plt.rcParams["font.sans-serif"] = ["DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False


def open_wrds_connection(wrds_user, wrds_password):
    """Open a direct WRDS PostgreSQL connection without interactive prompts."""
    connect_kwargs = {
        "host": WRDS_POSTGRES_HOST,
        "port": WRDS_POSTGRES_PORT,
        "dbname": WRDS_POSTGRES_DB,
        "user": wrds_user,
        "password": wrds_password,
        "connect_timeout": 10,
    }
    # WRDS recommends SSL; keep compatibility with wrds defaults when available.
    if isinstance(WRDS_CONNECT_ARGS, dict):
        connect_kwargs.update(WRDS_CONNECT_ARGS)
    return psycopg2.connect(**connect_kwargs)


def run_raw_sql(conn, sql, params=None, date_cols=None):
    """Execute SQL through pandas and normalize optional date columns."""
    df = pd.read_sql_query(sql, conn, params=params)
    if date_cols:
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def format_auth_error(exc):
    """Convert DB auth exceptions into short user-facing diagnostics."""
    msg = str(exc).strip().replace("\n", " ")
    if not msg:
        msg = exc.__class__.__name__
    if len(msg) > 220:
        msg = msg[:220] + "..."
    return msg


def clear_runtime_state():
    """Clear result buffers so old data is not reused across credential switches."""
    st.session_state.result = None
    st.session_state.last_ticker = ""
    st.session_state.last_year = ""


def build_credential_fingerprint(wrds_user, wrds_password):
    """Build a stable non-plaintext fingerprint for credential change detection."""
    raw = f"{wrds_user}\0{wrds_password}".encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def validate_credentials(wrds_user, wrds_password):
    """Validate WRDS credentials with a lightweight query before full data requests."""
    conn = None
    try:
        conn = open_wrds_connection(wrds_user, wrds_password)
        run_raw_sql(conn, "SELECT 1 AS ok")
        return True, ""
    except Exception as exc:
        return False, format_auth_error(exc)
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass


def get_secret_string(container, key):
    """Safely read one optional string value from Streamlit secrets."""
    if container is None:
        return ""
    try:
        value = container.get(key, "")
    except Exception:
        value = ""
    return str(value).strip() if value is not None else ""


def load_wrds_secrets():
    """Load WRDS credentials from Streamlit Cloud secrets if configured."""
    secret_user = get_secret_string(st.secrets, "WRDS_USER")
    secret_password = get_secret_string(st.secrets, "WRDS_PASSWORD")

    if not secret_user or not secret_password:
        try:
            wrds_block = st.secrets.get("wrds", {})
        except Exception:
            wrds_block = {}
        secret_user = secret_user or get_secret_string(wrds_block, "user")
        secret_password = secret_password or get_secret_string(wrds_block, "password")

    return secret_user, secret_password


# ============================
# SIC Industry Benchmark Query
# ============================
# Pulls yearly SIC peer averages from Compustat funda for 2015-2024.
def get_industry_avg(sic_code, wrds_user, wrds_password):
    if sic_code is None or pd.isna(sic_code):
        return pd.DataFrame()

    conn = None
    try:
        sic_code = int(float(sic_code))
        conn = open_wrds_connection(wrds_user, wrds_password)

        ind_df = run_raw_sql(
            conn,
            """
            SELECT fyear,
                   sich AS sic_code,
                   COUNT(*) AS num_obs,
                   AVG(sale) AS avg_sale,
                   AVG(at) AS avg_total_assets,
                   AVG(ceq) AS avg_common_equity
            FROM comp.funda
            WHERE sich = %s
            AND fyear BETWEEN 2015 AND 2024
            AND datafmt = 'STD'
            AND consol = 'C'
            AND indfmt = 'INDL'
            GROUP BY fyear, sich
            ORDER BY fyear
            """,
            params=(sic_code,),
        )
        return ind_df.round(2)
    except Exception:
        return pd.DataFrame()
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass


def get_company_info(ticker, wrds_user, wrds_password):
    # Tries Compustat funda first (latest valid record), then falls back to
    # comp.company SIC when no funda row is available.
    conn = None
    try:
        conn = open_wrds_connection(wrds_user, wrds_password)
        sql = """
            SELECT tic, conm, sich
            FROM comp.funda
            WHERE UPPER(tic) = UPPER(%s)
            AND datafmt = 'STD' AND consol = 'C' AND indfmt = 'INDL'
            AND sich IS NOT NULL
            ORDER BY datadate DESC NULLS LAST, fyear DESC NULLS LAST
            LIMIT 1
        """
        df = run_raw_sql(conn, sql, params=(ticker,))

        if df.empty:
            sql_fallback = """
                SELECT tic, conm, sic AS sich
                FROM comp.company
                WHERE UPPER(tic) = UPPER(%s)
                AND sic IS NOT NULL
                LIMIT 1
            """
            df = run_raw_sql(conn, sql_fallback, params=(ticker,))

        return df
    except Exception:
        return pd.DataFrame()
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass


def get_year_quarters(year):
    return [
        (f"{year}-01-01", f"{year}-03-31"),
        (f"{year}-04-01", f"{year}-06-30"),
        (f"{year}-07-01", f"{year}-09-30"),
        (f"{year}-10-01", f"{year}-12-31"),
    ]


# ============================
# Daily Stock Data (CRSP)
# ============================
# Pulls selected-year daily pricing/return/volume/share-outstanding data.
def get_single_year_daily(ticker, year, wrds_user, wrds_password):
    quarters = get_year_quarters(year)
    all_data = []

    conn = None
    try:
        conn = open_wrds_connection(wrds_user, wrds_password)
        permno_sql = "SELECT DISTINCT permno FROM crsp.stocknames WHERE UPPER(ticker)=UPPER(%s) LIMIT 1"
        permno_df = run_raw_sql(conn, permno_sql, params=(ticker,))
        if permno_df.empty:
            return pd.DataFrame(), "Stock not found in CRSP"
        permno = int(permno_df.iloc[0]["permno"])
    except Exception:
        return pd.DataFrame(), "WRDS connect failed"
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass

    for start_date, end_date in quarters:
        # Query by quarter windows to keep request sizes manageable.
        conn = None
        try:
            conn = open_wrds_connection(wrds_user, wrds_password)
            sql = (
                "SELECT date, prc, ret, vol, shrout "
                "FROM crsp.dsf "
                f"WHERE permno={permno} AND date>='{start_date}' AND date<='{end_date}' "
                "ORDER BY date"
            )
            df_q = run_raw_sql(conn, sql, date_cols=["date"])
            if not df_q.empty:
                all_data.append(df_q)
        except Exception:
            continue
        finally:
            if conn is not None:
                try:
                    conn.close()
                except Exception:
                    pass

    if not all_data:
        return pd.DataFrame(), "No daily price data for selected year"

    df = pd.concat(all_data, ignore_index=True)
    df["close"] = df["prc"].abs()
    df["daily_return"] = df["ret"]
    df["volume"] = df["vol"]
    df["market_cap"] = (df["close"] * df["shrout"].fillna(0) * SHROUT_MULTIPLIER).round(2)

    stock_df = (
        df[["date", "close", "daily_return", "volume", "market_cap"]]
        .round(4)
        .sort_values("date")
        .drop_duplicates()
    )
    return stock_df, ""


# ============================
# Financial Statement + DuPont Metrics
# ============================
# Pulls annual fundamentals and computes DuPont decomposition components.
def get_financial_data(ticker, wrds_user, wrds_password):
    conn = None
    try:
        conn = open_wrds_connection(wrds_user, wrds_password)
        sql = """
        SELECT gvkey, tic, conm, datadate, ni, sale, at, ceq, lt, ebit, pi,
        ROUND(ni/sale,4) as profit_margin, ROUND(sale/at,4) as asset_turnover,
        ROUND(at/ceq,4) as equity_multiplier, ROUND((ni/sale)*(sale/at)*(at/ceq),4) as roe_dupont
        FROM comp.funda
        WHERE UPPER(tic)=UPPER(%s) AND datadate>='2015-01-01' AND datadate<='2024-12-31'
        AND sale>0 AND at>0 AND ceq>0 ORDER BY datadate;
        """
        df = run_raw_sql(conn, sql, params=(ticker,), date_cols=["datadate"])
        rename = {
            "tic": "Ticker",
            "conm": "Company",
            "datadate": "Date",
            "ni": "NetIncome",
            "sale": "Revenue",
            "at": "TotalAssets",
            "lt": "TotalLiabs",
            "ceq": "TotalEquity",
            "ebit": "EBIT",
            "pi": "PretaxIncome",
            "profit_margin": "ProfitMargin",
            "asset_turnover": "AssetTurnover",
            "equity_multiplier": "EqMultiplier",
            "roe_dupont": "ROE_DuPont",
        }
        out_df = df.rename(columns=rename)
        for required_col in [
            "NetIncome",
            "Revenue",
            "TotalAssets",
            "TotalLiabs",
            "TotalEquity",
            "EBIT",
            "PretaxIncome",
            "ProfitMargin",
            "AssetTurnover",
            "EqMultiplier",
            "ROE_DuPont",
        ]:
            if required_col not in out_df.columns:
                out_df[required_col] = pd.NA

        # Add robust extended metrics using existing fields only.
        def safe_div(numerator_col, denominator_col):
            numerator = pd.to_numeric(out_df.get(numerator_col), errors="coerce")
            denominator = pd.to_numeric(out_df.get(denominator_col), errors="coerce")
            denominator = denominator.replace(0, pd.NA)
            return numerator / denominator

        out_df["ROA"] = safe_div("NetIncome", "TotalAssets")
        out_df["DebtRatio"] = safe_div("TotalLiabs", "TotalAssets")
        out_df["DebtToEquity"] = safe_div("TotalLiabs", "TotalEquity")
        out_df["EquityRatio"] = safe_div("TotalEquity", "TotalAssets")
        out_df["CapitalIntensity"] = safe_div("TotalAssets", "Revenue")
        out_df["LiabilityToRevenue"] = safe_div("TotalLiabs", "Revenue")
        out_df["EBITMargin"] = safe_div("EBIT", "Revenue")
        out_df["PretaxMargin"] = safe_div("PretaxIncome", "Revenue")
        out_df["TaxBurden"] = safe_div("NetIncome", "PretaxIncome")
        out_df["InterestBurden"] = safe_div("PretaxIncome", "EBIT")
        out_df["EBITToAssets"] = safe_div("EBIT", "TotalAssets")

        # ROC uses TotalEquity + TotalLiabs as a simple capital base.
        capital_base = (
            pd.to_numeric(out_df.get("TotalEquity"), errors="coerce")
            + pd.to_numeric(out_df.get("TotalLiabs"), errors="coerce")
        )
        capital_base = capital_base.replace(0, pd.NA)
        out_df["ROC"] = pd.to_numeric(out_df.get("NetIncome"), errors="coerce") / capital_base

        return out_df
    except Exception:
        return pd.DataFrame()
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass


def get_chart_theme():
    """Return contrast-safe chart styling based on Streamlit light/dark theme."""
    base = st.get_option("theme.base") or "light"
    if base == "dark":
        return {
            "text": "#E8EEF7",
            "grid": "#4A5464",
            "palette": ["#66B2FF", "#FF9D6E", "#6FE59D", "#FF7FA8", "#D9A8FF", "#4DD5FF"],
            "bg": "none",
        }
    return {
        "text": "#1C2430",
        "grid": "#C7CFDA",
        "palette": ["#1F77B4", "#D97706", "#15803D", "#BE185D", "#7C3AED", "#0E7490"],
        "bg": "none",
    }


def make_multi_line_chart(
    df,
    x_col,
    selected_series,
    title,
    y_label,
    is_currency=False,
    marker=False,
    fixed_color=None,
    currency_divisor=1e9,
):
    """Render a theme-aware line chart with dynamic selected metrics."""
    theme = get_chart_theme()
    fig, ax = plt.subplots(figsize=(11.8, 4.8), dpi=120)

    for idx, (series_col, series_label) in enumerate(selected_series):
        color = fixed_color or theme["palette"][idx % len(theme["palette"])]
        clean_df = df[[x_col, series_col]].dropna()
        if clean_df.empty:
            continue
        ax.plot(
            clean_df[x_col],
            clean_df[series_col],
            label=series_label,
            linewidth=2.3,
            marker="o" if marker else None,
            color=color,
        )

    ax.set_title(title, fontsize=15, pad=12, fontweight="bold", color=theme["text"])
    ax.set_ylabel(y_label, fontsize=12, color=theme["text"])
    ax.tick_params(axis="x", labelsize=10, colors=theme["text"])
    ax.tick_params(axis="y", labelsize=10, colors=theme["text"])
    ax.grid(alpha=0.35, linestyle="--", linewidth=0.8, color=theme["grid"])
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["bottom"].set_color(theme["grid"])
    ax.spines["left"].set_color(theme["grid"])
    ax.set_facecolor(theme["bg"])
    fig.patch.set_alpha(0)
    if is_currency:
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"{x/currency_divisor:.1f}B"))
    ax.legend(frameon=False, ncol=2, loc="upper left", labelcolor=theme["text"])
    fig.tight_layout()
    return fig


def inject_custom_css():
    # Injects all visual styles for section headers and table themes.
    base = st.get_option("theme.base") or "light"
    section_title_color = "#D1D5DB" if base == "dark" else "#6B7280"
    table_title_color = "#D1D5DB" if base == "dark" else "#6B7280"
    table_header_text_color = "#FFFFFF"
    table_body_text_color = "#4B5563"
    css = """
        <style>
        .mf-section-title {
            font-size: 2rem;
            font-weight: 700;
            line-height: 1.2;
            margin-top: 1.1rem;
            margin-bottom: 0.7rem;
            color: __SECTION_TITLE_COLOR__ !important;
            border-left: 6px solid var(--accent);
            padding: 0.18rem 0 0.18rem 0.7rem;
            border-radius: 2px;
        }
        .mf-table-title {
            font-size: 1.08rem;
            font-weight: 700;
            margin: 0.65rem 0 0.25rem 0;
            color: __TABLE_TITLE_COLOR__ !important;
        }
        .mf-table-wrap {
            width: 100%;
            overflow-x: auto;
            border: 1px solid #3B4253;
            border-radius: 10px;
            margin-bottom: 0.95rem;
        }
        .mf-table {
            border-collapse: collapse;
            width: 100%;
            min-width: 860px;
        }
        .mf-table th, .mf-table td {
            padding: 0.42rem 0.55rem;
            border-bottom: 1px solid rgba(17, 24, 39, 0.08);
            font-size: 0.89rem;
            white-space: nowrap;
        }
        .mf-table-wrap.profile th { background: #16324F; color: __TABLE_HEADER_TEXT_COLOR__; }
        .mf-table-wrap.profile tbody tr:nth-child(odd) td { background: #E9F2FA; color: __TABLE_BODY_TEXT_COLOR__; }
        .mf-table-wrap.profile tbody tr:nth-child(even) td { background: #F4F8FC; color: __TABLE_BODY_TEXT_COLOR__; }

        .mf-table-wrap.stock th { background: #0F766E; color: __TABLE_HEADER_TEXT_COLOR__; }
        .mf-table-wrap.stock tbody tr:nth-child(odd) td { background: #D1FAE5; color: __TABLE_BODY_TEXT_COLOR__; }
        .mf-table-wrap.stock tbody tr:nth-child(even) td { background: #ECFDF5; color: __TABLE_BODY_TEXT_COLOR__; }

        .mf-table-wrap.financial th { background: #92400E; color: __TABLE_HEADER_TEXT_COLOR__; }
        .mf-table-wrap.financial tbody tr:nth-child(odd) td { background: #FFEDD5; color: __TABLE_BODY_TEXT_COLOR__; }
        .mf-table-wrap.financial tbody tr:nth-child(even) td { background: #FFF7ED; color: __TABLE_BODY_TEXT_COLOR__; }

        .mf-table-wrap.industry th { background: #4C1D95; color: __TABLE_HEADER_TEXT_COLOR__; }
        .mf-table-wrap.industry tbody tr:nth-child(odd) td { background: #EDE9FE; color: __TABLE_BODY_TEXT_COLOR__; }
        .mf-table-wrap.industry tbody tr:nth-child(even) td { background: #F5F3FF; color: __TABLE_BODY_TEXT_COLOR__; }
        </style>
        """
    css = (
        css.replace("__SECTION_TITLE_COLOR__", section_title_color)
        .replace("__TABLE_TITLE_COLOR__", table_title_color)
        .replace("__TABLE_HEADER_TEXT_COLOR__", table_header_text_color)
        .replace("__TABLE_BODY_TEXT_COLOR__", table_body_text_color)
    )
    st.markdown(
        css,
        unsafe_allow_html=True,
    )


def render_section_title(number, title, accent_color, text_color=None):
    extra_style = f" color: {text_color};" if text_color else ""
    st.markdown(
        f'<div class="mf-section-title" style="--accent: {accent_color};{extra_style}">{number}. {title}</div>',
        unsafe_allow_html=True,
    )


def render_table_block(title, df, theme, max_rows=None):
    shown_df = df.head(max_rows) if max_rows is not None else df
    st.markdown(f"<div class='mf-table-title'>{title}</div>", unsafe_allow_html=True)
    table_html = shown_df.to_html(index=False, classes="mf-table", border=0)
    st.markdown(f"<div class='mf-table-wrap {theme}'>{table_html}</div>", unsafe_allow_html=True)


def render_metric_button_group(label, options, state_key, columns_per_row=4):
    # Renders flat metric buttons and persists the selected option in session state.
    st.markdown(f"<div class='mf-table-title'>{label}</div>", unsafe_allow_html=True)
    current = st.session_state.get(state_key, options[0])
    if current not in options:
        current = options[0]
        st.session_state[state_key] = current

    def select_metric(option):
        st.session_state[state_key] = option

    for row_start in range(0, len(options), columns_per_row):
        row_options = options[row_start:row_start + columns_per_row]
        columns = st.columns(len(row_options), gap="small")
        for column, option in zip(columns, row_options):
            column.button(
                option,
                key=f"{state_key}_{option}",
                use_container_width=True,
                type="primary" if option == current else "secondary",
                on_click=select_metric,
                args=(option,),
            )

    return st.session_state.get(state_key, current)


def build_excel(info_df, stock_df, financial_df, industry_df, ticker, year):
    # Exports all retrieved datasets into one multi-sheet Excel workbook.
    file_name = f"{ticker}_{year}_Full_Data.xlsx"
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if info_df is not None and not info_df.empty:
            info_df.to_excel(writer, sheet_name="Company_SIC", index=False)
        if stock_df is not None and not stock_df.empty:
            stock_df.to_excel(writer, sheet_name="Stock_Data", index=False)
        if financial_df is not None and not financial_df.empty:
            financial_df.to_excel(writer, sheet_name="Financial_Data", index=False)
        if industry_df is not None and not industry_df.empty:
            industry_df.to_excel(writer, sheet_name="Industry_Avg", index=False)
    output.seek(0)
    return file_name, output


def run_query(ticker, year, wrds_user, wrds_password):
    # End-to-end query orchestration used by the Search action.
    result = {
        "info_df": pd.DataFrame(),
        "stock_df": pd.DataFrame(),
        "financial_df": pd.DataFrame(),
        "industry_df": pd.DataFrame(),
        "industry_reason": "",
        "stock_reason": "",
    }

    info_df = get_company_info(ticker, wrds_user, wrds_password)
    result["info_df"] = info_df

    if info_df is not None and not info_df.empty:
        try:
            current_sic = int(info_df["sich"].iloc[0])
            industry_df = get_industry_avg(current_sic, wrds_user, wrds_password)
            result["industry_df"] = industry_df
            if industry_df is None or industry_df.empty:
                result["industry_reason"] = f"No industry rows for SIC {current_sic} in 2015-2024"
        except Exception as e:
            result["industry_reason"] = f"Industry query failed: {e}"
    else:
        result["industry_reason"] = "No valid SIC found for this ticker"

    stock_df, stock_reason = get_single_year_daily(ticker, year, wrds_user, wrds_password)
    result["stock_df"] = stock_df
    result["stock_reason"] = stock_reason

    result["financial_df"] = get_financial_data(ticker, wrds_user, wrds_password)
    return result


def render_app():
    # Main Streamlit entry point: input panel, query handling, charts, tables, download.
    st.set_page_config(page_title="Max Finance", layout="wide")
    inject_custom_css()
    st.title("Max Finance")
    st.caption("A Python-based WRDS-powered U.S. equity data retrieval platform.")
    section_text_color = "#D1D5DB" if (st.get_option("theme.base") or "light") == "dark" else "#6B7280"

    if "default_ticker" not in st.session_state:
        # Easter egg: assign a random default ticker at first load.
        st.session_state.default_ticker = random.choice(DEFAULT_TICKER_POOL)

    if "result" not in st.session_state:
        st.session_state.result = None
        st.session_state.last_ticker = ""
        st.session_state.last_year = ""

    if "cached_wrds_user" not in st.session_state:
        st.session_state.cached_wrds_user = ""
    if "credential_fingerprint" not in st.session_state:
        st.session_state.credential_fingerprint = ""
    if "auth_error" not in st.session_state:
        st.session_state.auth_error = ""

    secret_user, secret_password = load_wrds_secrets()
    has_secret_credentials = bool(secret_user and secret_password)
    if "use_secret_credentials" not in st.session_state:
        st.session_state.use_secret_credentials = has_secret_credentials

    with st.sidebar:
        # Sidebar handles credentials, ticker input, and one-year stock selector.
        st.header("Input Panel")
        if has_secret_credentials:
            use_secret_credentials = st.toggle(
                "Use Streamlit Secrets credentials",
                value=st.session_state.use_secret_credentials,
                help="Recommended for public deployment. Configure WRDS_USER/WRDS_PASSWORD in app secrets.",
            )
            st.session_state.use_secret_credentials = use_secret_credentials
        else:
            use_secret_credentials = False

        if use_secret_credentials:
            wrds_user = secret_user
            wrds_password = secret_password
            st.caption("Credentials are loaded from Streamlit Secrets.")
        else:
            wrds_user = st.text_input("WRDS User", value=st.session_state.cached_wrds_user).strip()
            wrds_password = st.text_input("WRDS Password", type="password").strip()

        ticker = st.text_input("Ticker", value=st.session_state.default_ticker).strip().upper()
        year = st.selectbox("Year", YEAR_OPTIONS, index=len(YEAR_OPTIONS) - 1)
        st.caption("Year selector controls stock data only. Financial, DuPont, SIC use full 2015-2024.")
        run_btn = st.button("Search", width="stretch")

    # Cache credentials in session state; reset previous result when credentials change.
    current_fingerprint = build_credential_fingerprint(wrds_user, wrds_password) if wrds_user and wrds_password else ""
    if current_fingerprint and current_fingerprint != st.session_state.credential_fingerprint:
        st.session_state.cached_wrds_user = wrds_user
        st.session_state.credential_fingerprint = current_fingerprint
        st.session_state.auth_error = ""
        clear_runtime_state()

    if run_btn:
        # Validate inputs and credentials before executing full WRDS queries.
        st.session_state.auth_error = ""
        if not ticker:
            st.session_state.auth_error = "Please enter a ticker symbol."
            clear_runtime_state()
        elif not wrds_user or not wrds_password:
            if has_secret_credentials:
                st.session_state.auth_error = "WRDS credentials are missing. Disable secrets mode or check app secrets."
            else:
                st.session_state.auth_error = "Please enter WRDS credentials."
            clear_runtime_state()
        else:
            ok, reason = validate_credentials(wrds_user, wrds_password)
            if not ok:
                clear_runtime_state()
                st.session_state.auth_error = f"WRDS login failed: {reason}"
            else:
                with st.spinner("Loading data from WRDS..."):
                    st.session_state.result = run_query(ticker, year, wrds_user, wrds_password)
                    st.session_state.last_ticker = ticker
                    st.session_state.last_year = year

    if st.session_state.auth_error:
        st.error(st.session_state.auth_error)

    result = st.session_state.result
    if result is None:
        if not st.session_state.auth_error:
            st.info("Ready")
        return

    info_df = result["info_df"]
    stock_df = result["stock_df"]
    financial_df = result["financial_df"]
    industry_df = result["industry_df"]
    current_sic = "-"
    if info_df is not None and not info_df.empty and "sich" in info_df.columns:
        try:
            current_sic = int(info_df["sich"].iloc[0])
        except Exception:
            current_sic = "-"

    render_section_title(1, "Stock Data Visualization", "#38BDF8", text_color=section_text_color)
    if stock_df is None or stock_df.empty:
        st.write("No pricing data available.")
        if result["stock_reason"]:
            st.caption(f"Reason: {result['stock_reason']}")
    else:
        stock_metric_map = {
            "Market Cap": "market_cap",
            "Close": "close",
            "Daily Return": "daily_return",
            "Volume": "volume",
        }
        st.session_state.setdefault("stock_metric_pick", "Market Cap")
        stock_pick = render_metric_button_group(
            "Select one stock metric",
            list(stock_metric_map.keys()),
            "stock_metric_pick",
        )
        selected_series = [(stock_metric_map[stock_pick], stock_pick)]
        stock_fig = make_multi_line_chart(
            stock_df,
            "date",
            selected_series,
            f"{st.session_state.last_ticker} Stock Trend ({st.session_state.last_year})",
            "Value",
            is_currency=stock_pick in ["Close", "Market Cap"],
            fixed_color="#38BDF8",
        )
        st.pyplot(stock_fig, clear_figure=True)

    render_section_title(2, "Financial Data Visualization", "#F59E0B", text_color=section_text_color)
    if financial_df is None or financial_df.empty:
        st.write("No data.")
    else:
        financial_metric_map = {
            "Revenue": "Revenue",
            "Net Income": "NetIncome",
            "EBIT": "EBIT",
            "Pretax Income": "PretaxIncome",
            "Total Assets": "TotalAssets",
            "Total Liabilities": "TotalLiabs",
            "Total Equity": "TotalEquity",
        }
        st.session_state.setdefault("financial_metric_pick", "Revenue")
        fin_pick = render_metric_button_group(
            "Select one reported fundamentals metric",
            list(financial_metric_map.keys()),
            "financial_metric_pick",
        )
        fin_series = [(financial_metric_map[fin_pick], fin_pick)]
        fin_fig = make_multi_line_chart(
            financial_df,
            "Date",
            fin_series,
            f"{st.session_state.last_ticker} Reported Fundamentals Trend (2015-2024)",
            "USD",
            is_currency=True,
            fixed_color="#F59E0B",
            currency_divisor=1e3,
        )
        st.pyplot(fin_fig, clear_figure=True)

    render_section_title(3, "Derived Metrics Visualization", "#34D399", text_color=section_text_color)
    if financial_df is None or financial_df.empty:
        st.write("No data.")
    else:
        dupont_metric_map = {
            "ROE (DuPont)": "ROE_DuPont",
            "Profit Margin (DuPont)": "ProfitMargin",
            "Asset Turnover (DuPont)": "AssetTurnover",
            "Equity Multiplier (DuPont)": "EqMultiplier",
            "ROA": "ROA",
            "Debt Ratio": "DebtRatio",
            "Debt to Equity": "DebtToEquity",
            "Equity Ratio": "EquityRatio",
            "ROC": "ROC",
            "Capital Intensity": "CapitalIntensity",
            "Liability to Revenue": "LiabilityToRevenue",
            "EBIT Margin": "EBITMargin",
            "Pretax Margin": "PretaxMargin",
            "Tax Burden": "TaxBurden",
            "Interest Burden": "InterestBurden",
            "EBIT to Assets": "EBITToAssets",
        }
        st.session_state.setdefault("dupont_metric_pick", "ROE (DuPont)")
        dupont_pick = render_metric_button_group(
            "Select one derived metric",
            list(dupont_metric_map.keys()),
            "dupont_metric_pick",
        )
        dupont_series = [(dupont_metric_map[dupont_pick], dupont_pick)]
        dupont_fig = make_multi_line_chart(
            financial_df,
            "Date",
            dupont_series,
            f"{st.session_state.last_ticker} Derived Metrics Trend (2015-2024)",
            "Ratio",
            is_currency=False,
            fixed_color="#34D399",
        )
        st.pyplot(dupont_fig, clear_figure=True)

    render_section_title(4, "SIC Industry Visualization", "#A78BFA", text_color=section_text_color)
    if industry_df is not None and not industry_df.empty:
        industry_metric_map = {
            "Avg Sale": "avg_sale",
            "Avg Total Assets": "avg_total_assets",
            "Avg Common Equity": "avg_common_equity",
            "Observations": "num_obs",
        }
        st.session_state.setdefault("industry_metric_pick", "Avg Sale")
        industry_pick = render_metric_button_group(
            "Select one industry metric",
            list(industry_metric_map.keys()),
            "industry_metric_pick",
        )
        industry_series = [(industry_metric_map[industry_pick], industry_pick)]
        ind_fig = make_multi_line_chart(
            industry_df,
            "fyear",
            industry_series,
            f"SIC {current_sic} Industry Benchmark (2015-2024)",
            "Value",
            is_currency=industry_pick != "Observations",
            marker=True,
            fixed_color="#A78BFA",
            currency_divisor=1e3,
        )
        st.pyplot(ind_fig, clear_figure=True)
    else:
        st.write("No industry benchmark data available.")
        if result["industry_reason"]:
            st.caption(f"Reason: {result['industry_reason']}")

    render_section_title(5, "All Data Tables", "#22D3EE", text_color=section_text_color)
    # Tables are grouped after charts for quick scan first, detail inspection second.
    st.caption("Tables are grouped below after the chart sections.")

    if info_df is None or info_df.empty:
        st.write("No company profile available.")
    else:
        render_table_block("Company Profile", info_df, "profile")

    if stock_df is None or stock_df.empty:
        st.write("No stock data.")
    else:
        render_table_block("Stock Data (Top 30 rows)", stock_df, "stock", max_rows=30)

    if financial_df is None or financial_df.empty:
        st.write("No financial data.")
    else:
        reported_cols = [
            "Ticker",
            "Date",
            "NetIncome",
            "EBIT",
            "PretaxIncome",
            "Revenue",
            "TotalAssets",
            "TotalLiabs",
            "TotalEquity",
        ]
        reported_df = financial_df[[c for c in reported_cols if c in financial_df.columns]]
        render_table_block("Reported Fundamentals Data (2015-2024)", reported_df, "financial")

        derived_cols = [
            "Ticker",
            "Date",
            "ProfitMargin",
            "AssetTurnover",
            "EqMultiplier",
            "ROE_DuPont",
            "ROA",
            "ROC",
            "DebtRatio",
            "DebtToEquity",
            "EquityRatio",
            "CapitalIntensity",
            "LiabilityToRevenue",
            "EBITMargin",
            "PretaxMargin",
            "TaxBurden",
            "InterestBurden",
            "EBITToAssets",
        ]
        derived_df = financial_df[[c for c in derived_cols if c in financial_df.columns]]
        render_table_block("Derived Metrics Data (2015-2024)", derived_df, "financial")

    if industry_df is None or industry_df.empty:
        st.write("No industry data.")
    else:
        render_table_block("SIC Industry Benchmark (2015-2024)", industry_df, "industry")

    dl_name, dl_bytes = build_excel(
        info_df,
        stock_df,
        financial_df,
        industry_df,
        st.session_state.last_ticker,
        st.session_state.last_year,
    )
    st.download_button(
        label="Download",
        data=dl_bytes,
        file_name=dl_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )


if __name__ == "__main__":
    render_app()
