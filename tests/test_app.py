"""
Basic test suite for AI Analysis and Monitoring Agent.
Run with:  pytest tests/ -v
"""
import os
import sys
import pandas as pd

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import app as m  # noqa: E402


def _sample_df(rows=30):
    return m.generate_sample_data(days=rows)


def test_generate_sample_data_shape():
    df = _sample_df(30)
    assert len(df) == 30
    for col in m.settings.REQUIRED_COLUMNS:
        assert col in df.columns


def test_clean_dataframe_no_nulls():
    df = _sample_df(30)
    df.loc[2, "Revenue"] = None
    df.loc[5, "Date"] = None
    cleaned = m.clean_dataframe(df)
    assert cleaned.isna().sum().sum() == 0
    assert len(cleaned) <= len(df)


def test_calculate_kpis_adds_columns():
    df = m.clean_dataframe(_sample_df(30))
    kpi_df = m.calculate_kpis(df)
    for col in ["Revenue_Growth_%", "Profit_Margin_%", "Refund_Rate_%"]:
        assert col in kpi_df.columns


def test_pct_change_safe_handles_zero():
    assert m.pct_change_safe(100, 0) == 0.0


def test_classify_severity_bounds():
    assert m.classify_severity(5) == "Low"
    assert m.classify_severity(20) == "Medium"
    assert m.classify_severity(35) == "High"
    assert m.classify_severity(60) == "Critical"


def test_anomaly_detectors_return_lists():
    df = m.calculate_kpis(m.clean_dataframe(_sample_df(60)))
    for fn in (m.detect_moving_average, m.detect_pct_change, m.detect_zscore):
        result = fn(df, "Revenue")
        assert isinstance(result, list)


def test_business_rule_engine_returns_messages():
    df = m.calculate_kpis(m.clean_dataframe(_sample_df(10)))
    insights = m.business_rule_engine(df.iloc[-1], df.iloc[-2])
    assert isinstance(insights, list)
    assert len(insights) >= 1


def test_full_pipeline_smoke(tmp_path):
    file_path = tmp_path / "sample.xlsx"
    m.generate_sample_data(str(file_path), days=45)
    result = m.run_full_pipeline(str(file_path), persist=False)
    assert result["rows_processed"] == 45
    assert "ai_summary" in result
    assert isinstance(result["anomalies"], list)
