"""
=====================================================================================
AI ANALYSIS AND MONITORING AGENT
=====================================================================================
A single-file, production-style business monitoring system that:

  1. Reads & validates Excel business data
  2. Cleans / imputes / de-duplicates the data
  3. Calculates business KPIs (growth %, margins, ratios, trends)
  4. Detects anomalies using 4 independent methods
        - Moving Average deviation
        - Percentage Change threshold
        - Z-Score
        - Isolation Forest (scikit-learn)
  5. Runs a 25+ rule Business Rule Engine to explain *why* something happened
  6. Generates an AI executive summary via a local Ollama model (llama3.1:8b)
     with a safe, deterministic fallback if Ollama is not running
  7. Sends SMTP email alerts for High / Critical anomalies
  8. Persists metrics / anomalies / alerts to a database via SQLAlchemy
     (SQLite by default, MySQL by changing DATABASE_URL in .env)
  9. Exposes a FastAPI backend (Swagger docs included)
 10. Exposes a Streamlit dashboard (same file, different entry point)

-------------------------------------------------------------------------------------
HOW TO RUN
-------------------------------------------------------------------------------------
Backend API:      python app.py                 (or) uvicorn app:api --reload
Dashboard UI:     streamlit run app.py
Generate sample data:  python app.py --make-sample-data

The script auto-detects whether it is being executed by Streamlit or by Python/uvicorn
and boots the correct entry point, so this ONE file serves both roles.
-------------------------------------------------------------------------------------
"""

import os
import io
import sys
import json
import time
import smtplib
import logging
import sqlite3
import datetime as dt
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any

import numpy as np
import pandas as pd

# ------------------------------------------------------------------------------------
# 0. CONFIGURATION  (loaded from .env if python-dotenv is available)
# ------------------------------------------------------------------------------------
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


class Settings:
    DATABASE_URL: str = os.getenv("DATABASE_URL", "sqlite:///./business_monitor.db")
    # Example MySQL:  mysql+pymysql://user:password@localhost:3306/business_monitor

    SMTP_HOST: str = os.getenv("SMTP_HOST", "smtp.gmail.com")
    SMTP_PORT: int = int(os.getenv("SMTP_PORT", "587"))
    SMTP_USER: str = os.getenv("SMTP_USER", "")
    SMTP_PASSWORD: str = os.getenv("SMTP_PASSWORD", "")
    ALERT_EMAIL_FROM: str = os.getenv("ALERT_EMAIL_FROM", "")
    ALERT_EMAIL_TO: str = os.getenv("ALERT_EMAIL_TO", "")
    EMAIL_ENABLED: bool = os.getenv("EMAIL_ENABLED", "false").lower() == "true"

    OLLAMA_HOST: str = os.getenv("OLLAMA_HOST", "http://localhost:11434")
    OLLAMA_MODEL: str = os.getenv("OLLAMA_MODEL", "llama3.1:8b")

    DATA_DIR: str = os.getenv("DATA_DIR", "data")
    LOG_DIR: str = os.getenv("LOG_DIR", "logs")

    REQUIRED_COLUMNS = ["Date", "Revenue", "Orders", "Traffic",
                        "Conversion_Rate", "Cost", "Profit", "Refunds"]


settings = Settings()
os.makedirs(settings.DATA_DIR, exist_ok=True)
os.makedirs(settings.LOG_DIR, exist_ok=True)

# ------------------------------------------------------------------------------------
# 1. LOGGING
# ------------------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(settings.LOG_DIR, "app.log")),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("business_monitor")


# ------------------------------------------------------------------------------------
# 2. DATABASE LAYER (SQLAlchemy models -- works with SQLite or MySQL)
# ------------------------------------------------------------------------------------
from sqlalchemy import (create_engine, Column, Integer, String, Float, DateTime, Text)
from sqlalchemy.orm import declarative_base, sessionmaker

Base = declarative_base()
engine = create_engine(settings.DATABASE_URL, connect_args={"check_same_thread": False}
                        if settings.DATABASE_URL.startswith("sqlite") else {})
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)


class MetricRecord(Base):
    __tablename__ = "metrics"
    id = Column(Integer, primary_key=True, index=True)
    date = Column(DateTime, index=True)
    revenue = Column(Float)
    orders = Column(Float)
    traffic = Column(Float)
    conversion_rate = Column(Float)
    cost = Column(Float)
    profit = Column(Float)
    refunds = Column(Float)


class AnomalyRecord(Base):
    __tablename__ = "anomalies"
    id = Column(Integer, primary_key=True, index=True)
    date = Column(DateTime, index=True)
    metric = Column(String(50))
    method = Column(String(50))
    value = Column(Float)
    expected = Column(Float)
    deviation_pct = Column(Float)
    severity = Column(String(20))
    direction = Column(String(10))


class AlertRecord(Base):
    __tablename__ = "alerts"
    id = Column(Integer, primary_key=True, index=True)
    created_at = Column(DateTime, default=dt.datetime.utcnow)
    severity = Column(String(20))
    subject = Column(String(255))
    summary = Column(Text)
    emailed = Column(Integer, default=0)


class AuditLog(Base):
    __tablename__ = "audit_logs"
    id = Column(Integer, primary_key=True, index=True)
    created_at = Column(DateTime, default=dt.datetime.utcnow)
    action = Column(String(100))
    details = Column(Text)


class UserRecord(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String(100), unique=True)
    email = Column(String(150))
    role = Column(String(50), default="viewer")


def init_db():
    Base.metadata.create_all(bind=engine)
    logger.info("Database initialized at %s", settings.DATABASE_URL)


# ------------------------------------------------------------------------------------
# 3. EXCEL INGESTION + CLEANING MODULE
# ------------------------------------------------------------------------------------
class ExcelIngestionError(Exception):
    pass


def read_excel_file(file_path_or_buffer) -> pd.DataFrame:
    """Read an xlsx file/buffer and validate required columns exist."""
    try:
        df = pd.read_excel(file_path_or_buffer)
    except Exception as e:
        logger.error("Failed to read Excel file: %s", e)
        raise ExcelIngestionError(f"Could not read Excel file: {e}")

    missing = [c for c in settings.REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ExcelIngestionError(f"Missing required columns: {missing}")

    logger.info("Excel file read successfully with %d rows", len(df))
    return df


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Full cleaning pipeline: dates, duplicates, missing values, dtypes, outliers."""
    df = df.copy()

    # --- Date conversion ---
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    n_bad_dates = df["Date"].isna().sum()
    if n_bad_dates:
        logger.warning("Dropping %d rows with invalid dates", n_bad_dates)
    df = df.dropna(subset=["Date"])

    # --- Duplicate removal ---
    before = len(df)
    df = df.drop_duplicates(subset=["Date"], keep="last")
    if before != len(df):
        logger.warning("Removed %d duplicate rows", before - len(df))

    # --- Numeric dtype validation ---
    numeric_cols = ["Revenue", "Orders", "Traffic", "Conversion_Rate",
                     "Cost", "Profit", "Refunds"]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # --- Missing value treatment (median imputation, then forward/back fill) ---
    for col in numeric_cols:
        median_val = df[col].median()
        df[col] = df[col].fillna(median_val)
    df[numeric_cols] = df[numeric_cols].ffill().bfill()

    # --- Outlier handling (IQR clipping, not deletion, to preserve signal) ---
    for col in numeric_cols:
        q1, q3 = df[col].quantile(0.25), df[col].quantile(0.75)
        iqr = q3 - q1
        lower, upper = q1 - 3 * iqr, q3 + 3 * iqr
        df[col] = df[col].clip(lower=lower, upper=upper)

    df = df.sort_values("Date").reset_index(drop=True)
    logger.info("Cleaning complete. Final shape: %s", df.shape)
    return df


# ------------------------------------------------------------------------------------
# 4. KPI CALCULATION ENGINE
# ------------------------------------------------------------------------------------
def pct_change_safe(current, previous):
    if previous in (0, None) or pd.isna(previous):
        return 0.0
    return ((current - previous) / abs(previous)) * 100


def calculate_kpis(df: pd.DataFrame) -> pd.DataFrame:
    """Adds growth %, margins and ratios as new columns."""
    df = df.copy()
    df["Revenue_Growth_%"] = df["Revenue"].pct_change().fillna(0) * 100
    df["Order_Growth_%"] = df["Orders"].pct_change().fillna(0) * 100
    df["Traffic_Growth_%"] = df["Traffic"].pct_change().fillna(0) * 100
    df["Refund_Rate_%"] = np.where(df["Orders"] > 0,
                                    (df["Refunds"] / df["Orders"]) * 100, 0)
    df["Profit_Margin_%"] = np.where(df["Revenue"] > 0,
                                      (df["Profit"] / df["Revenue"]) * 100, 0)
    df["Cost_Ratio_%"] = np.where(df["Revenue"] > 0,
                                   (df["Cost"] / df["Revenue"]) * 100, 0)
    df["Conversion_Rate_Trend"] = df["Conversion_Rate"].rolling(7, min_periods=1).mean()
    return df


def kpi_summary(df: pd.DataFrame) -> Dict[str, float]:
    """Latest-vs-previous snapshot used by the dashboard KPI cards."""
    if len(df) < 2:
        latest, prev = df.iloc[-1], df.iloc[-1]
    else:
        latest, prev = df.iloc[-1], df.iloc[-2]
    return {
        "revenue": latest["Revenue"],
        "revenue_growth_pct": pct_change_safe(latest["Revenue"], prev["Revenue"]),
        "orders": latest["Orders"],
        "order_growth_pct": pct_change_safe(latest["Orders"], prev["Orders"]),
        "traffic": latest["Traffic"],
        "traffic_growth_pct": pct_change_safe(latest["Traffic"], prev["Traffic"]),
        "profit_margin_pct": latest.get("Profit_Margin_%", 0.0),
        "refund_rate_pct": latest.get("Refund_Rate_%", 0.0),
        "cost_ratio_pct": latest.get("Cost_Ratio_%", 0.0),
    }


# ------------------------------------------------------------------------------------
# 5. ANOMALY DETECTION ENGINE (4 methods)
# ------------------------------------------------------------------------------------
@dataclass
class Anomaly:
    date: dt.datetime
    metric: str
    method: str
    value: float
    expected: float
    deviation_pct: float
    severity: str
    direction: str  # "spike" or "drop"

    def to_dict(self):
        d = self.__dict__.copy()
        d["date"] = self.date.isoformat()
        return d


def classify_severity(deviation_pct: float) -> str:
    dev = abs(deviation_pct)
    if dev >= 50:
        return "Critical"
    elif dev >= 30:
        return "High"
    elif dev >= 15:
        return "Medium"
    else:
        return "Low"


def detect_moving_average(df: pd.DataFrame, metric: str, window: int = 7,
                           threshold_pct: float = 15.0) -> List[Anomaly]:
    anomalies = []
    rolling_mean = df[metric].rolling(window, min_periods=3).mean()
    for i in range(len(df)):
        expected = rolling_mean.iloc[i]
        actual = df[metric].iloc[i]
        if pd.isna(expected) or expected == 0:
            continue
        dev = pct_change_safe(actual, expected)
        if abs(dev) >= threshold_pct:
            anomalies.append(Anomaly(
                date=df["Date"].iloc[i], metric=metric, method="MovingAverage",
                value=actual, expected=expected, deviation_pct=dev,
                severity=classify_severity(dev),
                direction="spike" if dev > 0 else "drop"))
    return anomalies


def detect_pct_change(df: pd.DataFrame, metric: str,
                       threshold_pct: float = 20.0) -> List[Anomaly]:
    anomalies = []
    changes = df[metric].pct_change().fillna(0) * 100
    for i in range(1, len(df)):
        dev = changes.iloc[i]
        if abs(dev) >= threshold_pct:
            anomalies.append(Anomaly(
                date=df["Date"].iloc[i], metric=metric, method="PctChangeThreshold",
                value=df[metric].iloc[i], expected=df[metric].iloc[i - 1],
                deviation_pct=dev, severity=classify_severity(dev),
                direction="spike" if dev > 0 else "drop"))
    return anomalies


def detect_zscore(df: pd.DataFrame, metric: str, z_threshold: float = 2.5) -> List[Anomaly]:
    anomalies = []
    series = df[metric]
    mean, std = series.mean(), series.std()
    if std == 0 or pd.isna(std):
        return anomalies
    z_scores = (series - mean) / std
    for i in range(len(df)):
        z = z_scores.iloc[i]
        if abs(z) >= z_threshold:
            dev = pct_change_safe(series.iloc[i], mean)
            anomalies.append(Anomaly(
                date=df["Date"].iloc[i], metric=metric, method="ZScore",
                value=series.iloc[i], expected=mean, deviation_pct=dev,
                severity=classify_severity(z * 20),  # scale z into a severity-like pct
                direction="spike" if z > 0 else "drop"))
    return anomalies


def detect_isolation_forest(df: pd.DataFrame, metric: str,
                             contamination: float = 0.05) -> List[Anomaly]:
    anomalies = []
    try:
        from sklearn.ensemble import IsolationForest
    except ImportError:
        logger.warning("scikit-learn not installed; skipping Isolation Forest")
        return anomalies

    if len(df) < 10:
        return anomalies

    X = df[[metric]].values
    model = IsolationForest(contamination=contamination, random_state=42)
    preds = model.fit_predict(X)  # -1 = anomaly, 1 = normal
    mean = df[metric].mean()

    for i, p in enumerate(preds):
        if p == -1:
            dev = pct_change_safe(df[metric].iloc[i], mean)
            anomalies.append(Anomaly(
                date=df["Date"].iloc[i], metric=metric, method="IsolationForest",
                value=df[metric].iloc[i], expected=mean, deviation_pct=dev,
                severity=classify_severity(dev),
                direction="spike" if dev > 0 else "drop"))
    return anomalies


def run_all_anomaly_detectors(df: pd.DataFrame,
                               metrics: Optional[List[str]] = None) -> List[Anomaly]:
    if metrics is None:
        metrics = ["Revenue", "Orders", "Traffic", "Conversion_Rate",
                   "Cost", "Profit", "Refunds"]
    all_anomalies: List[Anomaly] = []
    for metric in metrics:
        all_anomalies += detect_moving_average(df, metric)
        all_anomalies += detect_pct_change(df, metric)
        all_anomalies += detect_zscore(df, metric)
        all_anomalies += detect_isolation_forest(df, metric)
    logger.info("Detected %d raw anomaly signals", len(all_anomalies))
    return all_anomalies


# ------------------------------------------------------------------------------------
# 6. BUSINESS RULE ENGINE (25+ rules)
# ------------------------------------------------------------------------------------
def business_rule_engine(latest: pd.Series, prev: pd.Series) -> List[str]:
    """Compares the latest row to the previous row and returns plain-English
    explanations for the patterns detected. 25+ rules covering revenue, traffic,
    orders, cost, profit and refunds interactions."""

    insights = []

    def g(col):
        return pct_change_safe(latest[col], prev[col])

    rev_g, ord_g, traf_g = g("Revenue"), g("Orders"), g("Traffic")
    cost_g, profit_g, refund_g = g("Cost"), g("Profit"), g("Refunds")
    conv_g = g("Conversion_Rate")

    rules = [
        (rev_g < -10 and traf_g > 10, "Revenue dropped while traffic increased — visitors are arriving but not converting into paying customers."),
        (rev_g < -10 and ord_g < -10, "Revenue and orders both dropped sharply — demand may be weakening across the board."),
        (refund_g > 20, "Refunds increased significantly — this may indicate a product quality or fulfillment issue."),
        (cost_g > 10 and abs(rev_g) < 5, "Cost increased while revenue stayed flat — profitability may decline if this continues."),
        (ord_g < -15, "Orders dropped sharply — possible demand reduction, stockouts, or an operational/checkout issue."),
        (traf_g > 20 and conv_g < -10, "Traffic surged but conversion rate fell — the new visitors may be lower quality or the funnel has friction."),
        (rev_g > 20 and profit_g < 0, "Revenue grew but profit declined — costs may be outpacing sales growth."),
        (profit_g < -20, "Profit fell sharply — review both cost structure and pricing strategy."),
        (cost_g > 20, "Cost increased sharply — check supplier pricing, ad spend, or operational overhead."),
        (refund_g > 30 and ord_g > 0, "Refunds are rising even as orders grow — quality control should be reviewed urgently."),
        (traf_g < -20, "Traffic dropped significantly — check marketing channels, SEO rankings, or site outages."),
        (conv_g < -20, "Conversion rate collapsed — investigate checkout flow, pricing changes, or site performance."),
        (rev_g > 15 and traf_g < 0, "Revenue grew despite falling traffic — existing customers may be spending more (positive sign)."),
        (ord_g > 15 and rev_g < 5, "Orders grew faster than revenue — average order value may be shrinking (check discounting)."),
        (cost_g < -10 and profit_g > 10, "Costs decreased while profit improved — efficiency gains are paying off."),
        (refund_g < -20, "Refunds decreased notably — product quality or customer satisfaction may be improving."),
        (rev_g > 30, "Revenue spiked sharply — validate this is a real trend (promotion, seasonality) and not a data error."),
        (rev_g < -30, "Revenue dropped sharply — treat as a critical business event requiring immediate review."),
        (traf_g > 30 and rev_g > 30, "Both traffic and revenue surged together — marketing campaigns may be performing very well."),
        (ord_g > 0 and refund_g > 0 and refund_g > ord_g, "Refunds are growing faster than orders — a rising share of purchases are being returned."),
        (cost_g > 0 and cost_g > rev_g, "Cost is growing faster than revenue — margin compression risk."),
        (profit_g > 30, "Profit grew significantly — identify and reinforce what is driving this improvement."),
        (conv_g > 20, "Conversion rate improved notably — recent funnel or UX changes may be working well."),
        (abs(rev_g) < 2 and abs(traf_g) < 2 and abs(ord_g) < 2, "Metrics are stable with no significant movement — business is in a steady state."),
        (traf_g < -10 and ord_g > 10, "Traffic fell but orders rose — returning customers may be driving repeat purchases."),
        (refund_g > 15 and profit_g < -10, "Rising refunds are coinciding with falling profit — refund-driven margin erosion likely."),
        (rev_g > 10 and cost_g > 10 and profit_g > 10, "Revenue, cost, and profit all grew together — scaling appears healthy."),
    ]

    for condition, message in rules:
        if condition:
            insights.append(message)

    if not insights:
        insights.append("No significant business pattern detected in the latest period.")

    return insights


# ------------------------------------------------------------------------------------
# 7. AI INSIGHT GENERATOR (Ollama: llama3.1:8b, safe offline fallback)
# ------------------------------------------------------------------------------------
def build_ai_prompt(kpis: Dict[str, float], anomalies: List[Anomaly],
                     rule_insights: List[str]) -> str:
    anomaly_lines = "\n".join(
        f"- {a.metric}: {a.direction} of {a.deviation_pct:.1f}% ({a.severity}) via {a.method}"
        for a in anomalies[:10]
    ) or "None detected"

    return f"""You are a senior business analyst. Given the data below, write a concise report with
four sections: Executive Summary, Root Cause Analysis, Business Impact, Recommended Actions.

KPI Snapshot:
- Revenue: {kpis['revenue']:.2f} ({kpis['revenue_growth_pct']:.1f}% change)
- Orders: {kpis['orders']:.2f} ({kpis['order_growth_pct']:.1f}% change)
- Traffic: {kpis['traffic']:.2f} ({kpis['traffic_growth_pct']:.1f}% change)
- Profit Margin: {kpis['profit_margin_pct']:.1f}%
- Refund Rate: {kpis['refund_rate_pct']:.1f}%
- Cost Ratio: {kpis['cost_ratio_pct']:.1f}%

Detected Anomalies:
{anomaly_lines}

Rule-Based Observations:
{chr(10).join('- ' + r for r in rule_insights)}

Keep the report under 200 words, business-friendly, no jargon.
"""


def generate_ai_summary(kpis: Dict[str, float], anomalies: List[Anomaly],
                         rule_insights: List[str]) -> str:
    """Calls the local Ollama model (llama3.1:8b). Falls back to a deterministic
    template if Ollama is not reachable, so the system never breaks in an
    offline setting or before the model has been pulled locally."""
    prompt = build_ai_prompt(kpis, anomalies, rule_insights)

    try:
        import requests
        resp = requests.post(
            f"{settings.OLLAMA_HOST}/api/generate",
            json={"model": settings.OLLAMA_MODEL, "prompt": prompt, "stream": False},
            timeout=60,
        )
        if resp.status_code == 200:
            text = resp.json().get("response", "").strip()
            if text:
                logger.info("AI summary generated using Ollama model '%s'", settings.OLLAMA_MODEL)
                return text
    except Exception as e:
        logger.warning("Ollama call failed for model '%s': %s", settings.OLLAMA_MODEL, e)

    # ---- Deterministic offline fallback (Ollama not running / model not pulled) ----
    logger.info("Falling back to template-based summary (Ollama unavailable)")
    top_rule = rule_insights[0] if rule_insights else "No major pattern detected."
    return (
        f"Executive Summary: Revenue moved {kpis['revenue_growth_pct']:.1f}% and traffic "
        f"moved {kpis['traffic_growth_pct']:.1f}% in the latest period.\n\n"
        f"Root Cause Analysis: {top_rule}\n\n"
        f"Business Impact: Profit margin is currently {kpis['profit_margin_pct']:.1f}% "
        f"with a refund rate of {kpis['refund_rate_pct']:.1f}%.\n\n"
        f"Recommended Actions: Review the metrics and rules above, prioritizing any "
        f"Critical or High severity anomalies first."
    )


# ------------------------------------------------------------------------------------
# 8. EMAIL ALERT SYSTEM
# ------------------------------------------------------------------------------------
def send_email_alert(subject: str, anomalies: List[Anomaly], ai_summary: str) -> bool:
    if not settings.EMAIL_ENABLED:
        logger.info("Email disabled via EMAIL_ENABLED=false; skipping send.")
        return False
    if not (settings.SMTP_USER and settings.SMTP_PASSWORD and settings.ALERT_EMAIL_TO):
        logger.warning("Email settings incomplete; skipping send.")
        return False

    body_lines = ["<h2>Business Monitoring Alert</h2>"]
    body_lines.append(f"<p><b>Subject:</b> {subject}</p>")
    body_lines.append("<h3>Affected Metrics</h3><ul>")
    for a in anomalies:
        body_lines.append(
            f"<li>{a.metric} — {a.direction} of {a.deviation_pct:.1f}% "
            f"(<b>{a.severity}</b>, method: {a.method})</li>")
    body_lines.append("</ul>")
    body_lines.append(f"<h3>AI Business Summary</h3><p>{ai_summary.replace(chr(10), '<br>')}</p>")
    html_body = "\n".join(body_lines)

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = settings.ALERT_EMAIL_FROM or settings.SMTP_USER
    msg["To"] = settings.ALERT_EMAIL_TO
    msg.attach(MIMEText(html_body, "html"))

    try:
        with smtplib.SMTP(settings.SMTP_HOST, settings.SMTP_PORT) as server:
            server.starttls()
            server.login(settings.SMTP_USER, settings.SMTP_PASSWORD)
            server.send_message(msg)
        logger.info("Alert email sent to %s", settings.ALERT_EMAIL_TO)
        return True
    except Exception as e:
        logger.error("Failed to send alert email: %s", e)
        return False


def maybe_send_alert(anomalies: List[Anomaly], ai_summary: str, db) -> Optional[Dict[str, Any]]:
    """Returns a plain dict (not an ORM object) so the caller can safely use the
    result after the DB session that created it has been closed."""
    critical_or_high = [a for a in anomalies if a.severity in ("Critical", "High")]
    if not critical_or_high:
        return None

    subject = f"[{critical_or_high[0].severity}] Business Monitoring Alert — " \
              f"{len(critical_or_high)} anomaly(ies) detected"
    emailed = send_email_alert(subject, critical_or_high, ai_summary)

    record = AlertRecord(
        severity=critical_or_high[0].severity,
        subject=subject,
        summary=ai_summary,
        emailed=1 if emailed else 0,
    )
    db.add(record)
    db.commit()
    return {"severity": record.severity, "subject": subject, "emailed": emailed}


# ------------------------------------------------------------------------------------
# 9. CORE PIPELINE (used by both API and Dashboard)
# ------------------------------------------------------------------------------------
def run_full_pipeline(file_path_or_buffer, persist: bool = True) -> Dict[str, Any]:
    """End-to-end: ingest -> clean -> KPIs -> anomalies -> rules -> AI summary -> alert."""
    raw_df = read_excel_file(file_path_or_buffer)
    clean_df = clean_dataframe(raw_df)
    kpi_df = calculate_kpis(clean_df)

    anomalies = run_all_anomaly_detectors(kpi_df)
    kpis = kpi_summary(kpi_df)

    latest, prev = kpi_df.iloc[-1], kpi_df.iloc[-2] if len(kpi_df) > 1 else kpi_df.iloc[-1]
    rule_insights = business_rule_engine(latest, prev)

    ai_summary = generate_ai_summary(kpis, anomalies, rule_insights)

    alert_record = None
    if persist:
        init_db()  # idempotent: safe to call even if tables already exist
        db = SessionLocal()
        try:
            for _, row in kpi_df.iterrows():
                db.add(MetricRecord(
                    date=row["Date"], revenue=row["Revenue"], orders=row["Orders"],
                    traffic=row["Traffic"], conversion_rate=row["Conversion_Rate"],
                    cost=row["Cost"], profit=row["Profit"], refunds=row["Refunds"]))
            for a in anomalies:
                db.add(AnomalyRecord(
                    date=a.date, metric=a.metric, method=a.method, value=a.value,
                    expected=a.expected, deviation_pct=a.deviation_pct,
                    severity=a.severity, direction=a.direction))
            db.commit()
            alert_record = maybe_send_alert(anomalies, ai_summary, db)
            db.add(AuditLog(action="pipeline_run",
                             details=f"{len(kpi_df)} rows, {len(anomalies)} anomalies"))
            db.commit()
        finally:
            db.close()

    return {
        "kpis": kpis,
        "rows_processed": len(kpi_df),
        "anomalies": [a.to_dict() for a in anomalies],
        "rule_insights": rule_insights,
        "ai_summary": ai_summary,
        "alert_sent": bool(alert_record and alert_record.get("emailed")),
        "dataframe": kpi_df,
    }


# ------------------------------------------------------------------------------------
# 10. SAMPLE DATA GENERATOR
# ------------------------------------------------------------------------------------
def generate_sample_data(path: str = None, days: int = 365, seed: int = 42) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    dates = pd.date_range(end=dt.date.today(), periods=days, freq="D")

    base_traffic = 5000 + np.linspace(0, 1500, days) + rng.normal(0, 200, days)
    conv_rate = np.clip(rng.normal(0.03, 0.004, days), 0.005, 0.08)
    orders = base_traffic * conv_rate
    revenue = orders * rng.normal(55, 5, days)
    cost = revenue * rng.normal(0.55, 0.05, days)
    profit = revenue - cost
    refunds = orders * np.clip(rng.normal(0.03, 0.01, days), 0, 0.2)

    # inject realistic anomalies
    anomaly_days = rng.choice(days, size=max(5, days // 40), replace=False)
    for idx in anomaly_days:
        kind = rng.choice(["revenue_drop", "traffic_spike", "refund_spike", "cost_spike"])
        if kind == "revenue_drop":
            revenue[idx] *= 0.5
        elif kind == "traffic_spike":
            base_traffic[idx] *= 1.8
            orders[idx] = base_traffic[idx] * conv_rate[idx]
        elif kind == "refund_spike":
            refunds[idx] *= 4
        elif kind == "cost_spike":
            cost[idx] *= 1.6
            profit[idx] = revenue[idx] - cost[idx]

    df = pd.DataFrame({
        "Date": dates,
        "Revenue": revenue.round(2),
        "Orders": orders.round(0),
        "Traffic": base_traffic.round(0),
        "Conversion_Rate": conv_rate.round(4),
        "Cost": cost.round(2),
        "Profit": profit.round(2),
        "Refunds": refunds.round(0),
    })

    if path:
        df.to_excel(path, index=False)
        logger.info("Sample data written to %s (%d rows)", path, len(df))
    return df


# ------------------------------------------------------------------------------------
# 11. FASTAPI BACKEND
# ------------------------------------------------------------------------------------
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

api = FastAPI(
    title="AI Analysis and Monitoring Agent",
    description="Business metrics monitoring, anomaly detection, and AI insights API.",
    version="1.0.0",
)
api.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])


class PipelineResponse(BaseModel):
    kpis: Dict[str, float]
    rows_processed: int
    anomalies: List[Dict[str, Any]]
    rule_insights: List[str]
    ai_summary: str
    alert_sent: bool


_last_run_cache: Dict[str, Any] = {}


@api.on_event("startup")
def on_startup():
    init_db()
    logger.info("FastAPI backend started.")


@api.post("/upload", response_model=PipelineResponse)
async def upload_file(file: UploadFile = File(...)):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Only .xlsx/.xls files are supported.")
    try:
        contents = await file.read()
        result = run_full_pipeline(io.BytesIO(contents))
        _last_run_cache.update(result)
        result.pop("dataframe", None)
        return result
    except ExcelIngestionError as e:
        raise HTTPException(422, str(e))
    except Exception as e:
        logger.exception("Pipeline failed")
        raise HTTPException(500, f"Internal error: {e}")


@api.get("/metrics")
def get_metrics(limit: int = 100):
    db = SessionLocal()
    try:
        rows = db.query(MetricRecord).order_by(MetricRecord.date.desc()).limit(limit).all()
        return [{"date": r.date.isoformat(), "revenue": r.revenue, "orders": r.orders,
                  "traffic": r.traffic, "conversion_rate": r.conversion_rate,
                  "cost": r.cost, "profit": r.profit, "refunds": r.refunds} for r in rows]
    finally:
        db.close()


@api.get("/anomalies")
def get_anomalies(severity: Optional[str] = None, limit: int = 100):
    db = SessionLocal()
    try:
        q = db.query(AnomalyRecord)
        if severity:
            q = q.filter(AnomalyRecord.severity == severity)
        rows = q.order_by(AnomalyRecord.date.desc()).limit(limit).all()
        return [{"date": r.date.isoformat(), "metric": r.metric, "method": r.method,
                  "value": r.value, "expected": r.expected,
                  "deviation_pct": r.deviation_pct, "severity": r.severity,
                  "direction": r.direction} for r in rows]
    finally:
        db.close()


@api.get("/alerts")
def get_alerts(limit: int = 50):
    db = SessionLocal()
    try:
        rows = db.query(AlertRecord).order_by(AlertRecord.created_at.desc()).limit(limit).all()
        return [{"created_at": r.created_at.isoformat(), "severity": r.severity,
                  "subject": r.subject, "summary": r.summary,
                  "emailed": bool(r.emailed)} for r in rows]
    finally:
        db.close()


@api.get("/summary")
def get_summary():
    if not _last_run_cache:
        raise HTTPException(404, "No pipeline run yet. Call /upload first.")
    return {"ai_summary": _last_run_cache.get("ai_summary"),
            "rule_insights": _last_run_cache.get("rule_insights")}


@api.get("/dashboard-data")
def get_dashboard_data():
    if not _last_run_cache:
        raise HTTPException(404, "No pipeline run yet. Call /upload first.")
    data = {k: v for k, v in _last_run_cache.items() if k != "dataframe"}
    return data


# ------------------------------------------------------------------------------------
# 12. STREAMLIT DASHBOARD (runs only when launched via `streamlit run app.py`)
# ------------------------------------------------------------------------------------
def render_streamlit_dashboard():
    import streamlit as st
    import plotly.express as px

    st.set_page_config(page_title="AI Analysis and Monitoring Agent", layout="wide")
    st.title("📊 AI Analysis and Monitoring Agent")
    st.caption("Automated business metrics monitoring, anomaly detection & AI insights")

    init_db()

    with st.sidebar:
        st.header("Upload Data")
        uploaded = st.file_uploader("Upload business Excel file (.xlsx)", type=["xlsx"])
        st.markdown("---")
        st.header("Filters")
        severity_filter = st.multiselect(
            "Severity", ["Low", "Medium", "High", "Critical"],
            default=["Medium", "High", "Critical"])
        metric_filter = st.selectbox(
            "Metric", ["Revenue", "Orders", "Traffic", "Conversion_Rate",
                       "Cost", "Profit", "Refunds"])
        if st.button("Use Sample Data"):
            sample_path = os.path.join(settings.DATA_DIR, "sample_data.xlsx")
            if not os.path.exists(sample_path):
                generate_sample_data(sample_path)
            uploaded = open(sample_path, "rb")

    if not uploaded:
        st.info("Upload an Excel file or click 'Use Sample Data' in the sidebar to begin.")
        return

    with st.spinner("Running pipeline: cleaning, KPIs, anomaly detection, AI summary..."):
        result = run_full_pipeline(uploaded)

    df = result["dataframe"]
    kpis = result["kpis"]

    # ---- KPI Cards ----
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Revenue", f"{kpis['revenue']:.0f}", f"{kpis['revenue_growth_pct']:.1f}%")
    c2.metric("Orders", f"{kpis['orders']:.0f}", f"{kpis['order_growth_pct']:.1f}%")
    c3.metric("Traffic", f"{kpis['traffic']:.0f}", f"{kpis['traffic_growth_pct']:.1f}%")
    c4.metric("Profit Margin", f"{kpis['profit_margin_pct']:.1f}%")

    # ---- Trend Charts ----
    col1, col2 = st.columns(2)
    with col1:
        st.plotly_chart(px.line(df, x="Date", y="Revenue", title="Revenue Trend"),
                         use_container_width=True)
    with col2:
        st.plotly_chart(px.line(df, x="Date", y="Traffic", title="Traffic Trend"),
                         use_container_width=True)

    # ---- Anomaly Table + Severity Distribution ----
    anomalies_df = pd.DataFrame(result["anomalies"])
    st.subheader("Detected Anomalies")
    if not anomalies_df.empty:
        filtered = anomalies_df[anomalies_df["severity"].isin(severity_filter)]
        filtered = filtered[filtered["metric"] == metric_filter] if metric_filter else filtered
        st.dataframe(filtered, use_container_width=True)

        colA, colB = st.columns(2)
        with colA:
            st.plotly_chart(px.histogram(anomalies_df, x="severity",
                                          title="Severity Distribution"),
                             use_container_width=True)
        with colB:
            st.plotly_chart(px.histogram(anomalies_df, x="metric", color="severity",
                                          title="Anomalies by Metric"),
                             use_container_width=True)
    else:
        st.success("No anomalies detected in this dataset.")

    # ---- AI Insights Panel ----
    st.subheader("🤖 AI Business Insights")
    st.info(result["ai_summary"])
    with st.expander("Rule-Based Observations"):
        for r in result["rule_insights"]:
            st.write("• " + r)

    # ---- Alert History ----
    st.subheader("Alert History")
    db = SessionLocal()
    try:
        alerts = db.query(AlertRecord).order_by(AlertRecord.created_at.desc()).limit(20).all()
        if alerts:
            st.table(pd.DataFrame([{
                "Time": a.created_at, "Severity": a.severity,
                "Subject": a.subject, "Emailed": bool(a.emailed)} for a in alerts]))
        else:
            st.write("No alerts triggered yet.")
    finally:
        db.close()


# ------------------------------------------------------------------------------------
# 13. ENTRY POINT (auto-detect Streamlit vs. plain Python execution)
# ------------------------------------------------------------------------------------
def _running_under_streamlit() -> bool:
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx() is not None
    except Exception:
        return False


if _running_under_streamlit():
    render_streamlit_dashboard()
elif __name__ == "__main__":
    if "--make-sample-data" in sys.argv:
        out_path = os.path.join(settings.DATA_DIR, "sample_data.xlsx")
        generate_sample_data(out_path)
        print(f"Sample data written to {out_path}")
    else:
        import uvicorn
        init_db()
        uvicorn.run(api, host="0.0.0.0", port=8000)
