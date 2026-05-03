import requests
import json
import time
import re
import base64
import html
from urllib.parse import quote_plus
import xml.etree.ElementTree as ET

SEC_USER_AGENT = "StockAnalyzerPro/1.0 contact@example.com"
import os
from datetime import datetime, timedelta
from copy import copy

import yfinance as yf
import pandas as pd
import streamlit as st
import io
from openai import OpenAI

try:
    from openpyxl import load_workbook
except Exception:
    load_workbook = None

WATCHLIST_FILE = "watchlist.txt"
TEMPLATE_FILE = "market_template.xlsx"
CHART_CONFIRMATION_FILE = "chart_confirmation.json"
OUTPUT_FILE = "market_template.xlsx"

DEFAULT_WATCHLIST = [
    "ENVX", "SOUN", "WULF", "FRMI", "PATH", "RCAT", "QXO",
    "NVDA", "VGT", "VOO", "SLV", "GDX"
]

# ETF / Fund symbols: these do not have company-style earnings reports
# and usually should not be screened with SEC company filing risk.
ETF_SYMBOLS = {
    "VOO", "VTI", "VT", "SPY", "QQQ", "DIA", "IWM",
    "VGT", "VUG", "VTV", "SCHD", "XLK", "XLF", "XLE", "XLV", "XLY", "XLI", "XLP", "XLU", "XLB", "XLRE",
    "SMH", "SOXX", "ARKK", "ARKG", "ARKW",
    "SLV", "GLD", "IAU", "GDX", "GDXJ", "SIL", "SILJ",
    "URA", "URNM", "USO", "UNG", "TLT", "IEF", "SHY", "HYG", "LQD"
}

def is_etf_symbol(ticker):
    return str(ticker).strip().upper() in ETF_SYMBOLS

def load_chart_confirmations():
    if not os.path.exists(CHART_CONFIRMATION_FILE):
        return {}

    try:
        with open(CHART_CONFIRMATION_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_chart_confirmations(data):
    try:
        with open(CHART_CONFIRMATION_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass
def load_watchlist():
    if not os.path.exists(WATCHLIST_FILE):
        with open(WATCHLIST_FILE, "w", encoding="utf-8") as f:
            for ticker in DEFAULT_WATCHLIST:
                f.write(ticker + "\n")

    with open(WATCHLIST_FILE, "r", encoding="utf-8") as f:
        tickers = [line.strip().upper() for line in f if line.strip()]

    return list(dict.fromkeys(tickers))


def save_watchlist(tickers):
    tickers = [t.strip().upper() for t in tickers if t.strip()]
    tickers = list(dict.fromkeys(tickers))

    with open(WATCHLIST_FILE, "w", encoding="utf-8") as f:
        for ticker in tickers:
            f.write(ticker + "\n")


def download_price_data(ticker, period="8mo"):
    data = yf.download(
        ticker,
        period=period,
        interval="1d",
        progress=False,
        auto_adjust=False,
        threads=False
    )

    if isinstance(data.columns, pd.MultiIndex):
        data.columns = data.columns.get_level_values(0)

    return data.dropna()


def calculate_rsi(data, period=14):
    delta = data["Close"].diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)

    avg_gain = gain.ewm(alpha=1 / period, adjust=False).mean()
    avg_loss = loss.ewm(alpha=1 / period, adjust=False).mean()

    rs = avg_gain / avg_loss
    rsi = 100 - (100 / (1 + rs))
    return rsi


def calculate_macd(data):
    ema12 = data["Close"].ewm(span=12, adjust=False).mean()
    ema26 = data["Close"].ewm(span=26, adjust=False).mean()
    macd = ema12 - ema26
    signal = macd.ewm(span=9, adjust=False).mean()
    histogram = macd - signal
    return macd, signal, histogram


def calculate_atr(data, period=14):
    high_low = data["High"] - data["Low"]
    high_close = (data["High"] - data["Close"].shift()).abs()
    low_close = (data["Low"] - data["Close"].shift()).abs()
    true_range = pd.concat([high_low, high_close, low_close], axis=1).max(axis=1)
    return true_range.ewm(alpha=1 / period, adjust=False).mean()


def _cluster_price_levels(levels, tolerance_pct=1.2):
    """Cluster nearby pivot prices into support/resistance zones.
    Returns levels sorted by strength. Each level has price, touches, last_idx.
    """
    if not levels:
        return []
    levels = sorted(levels, key=lambda x: x["price"])
    clusters = []
    for lv in levels:
        placed = False
        for c in clusters:
            mid = c["price"]
            if mid > 0 and abs(lv["price"] - mid) / mid * 100 <= tolerance_pct:
                total = c["touches"] + 1
                c["price"] = (c["price"] * c["touches"] + lv["price"]) / total
                c["touches"] = total
                c["last_idx"] = max(c["last_idx"], lv["idx"])
                c["volume_score"] += lv.get("volume_score", 0)
                placed = True
                break
        if not placed:
            clusters.append({"price": lv["price"], "touches": 1, "last_idx": lv["idx"], "volume_score": lv.get("volume_score", 0)})
    return clusters


def find_chart_trade_levels(data, lookback=180):
    """Find levels like a trader drawing horizontal lines on TradingView.
    Logic:
    - detect swing highs/lows over recent daily candles
    - cluster close levels into zones
    - choose nearest strong resistance above current price
    - choose nearest support below current price and one deeper support
    - if price broke a small resistance intraday but closed back below it,
      treat that level as a minor failed pivot, not the main resistance.
    This mimics manual chart lines such as AMPX 22.79/19.76/18.81 and ONDS 11.02/9.16/8.72.
    """
    recent = data.tail(min(lookback, len(data))).copy().reset_index(drop=True)
    if recent.empty or len(recent) < 20:
        close = float(data["Close"].iloc[-1])
        return {"support": close * 0.95, "deep_support": close * 0.90, "resistance": close * 1.08, "method": "fallback"}

    close = float(recent["Close"].iloc[-1])
    avg_vol = float(recent["Volume"].tail(20).mean()) if "Volume" in recent else 0
    swing_lows, swing_highs = [], []
    window = 2
    for i in range(window, len(recent) - window):
        low = float(recent.loc[i, "Low"])
        high = float(recent.loc[i, "High"])
        vol = float(recent.loc[i, "Volume"]) if "Volume" in recent else 0
        vol_score = 1 if avg_vol > 0 and vol > avg_vol else 0
        if low <= float(recent.loc[i-window:i+window, "Low"].min()):
            swing_lows.append({"price": low, "idx": i, "volume_score": vol_score})
        if high >= float(recent.loc[i-window:i+window, "High"].max()):
            swing_highs.append({"price": high, "idx": i, "volume_score": vol_score})

    # Include important recent highs/lows so fresh breakout/pullback zones are not missed.
    # IMPORTANT: do NOT use the latest unfinished daily candle high as a main resistance.
    # Example ONDS: price traded above a minor intraday pivot then closed back below it;
    # that high should be treated as intraday noise/failed pivot, not Target 1.
    recent_start = max(0, len(recent) - 25)
    last_completed_idx = max(recent_start, len(recent) - 2)
    for i in range(recent_start, last_completed_idx + 1):
        swing_lows.append({"price": float(recent.loc[i, "Low"]), "idx": i, "volume_score": 0})
        swing_highs.append({"price": float(recent.loc[i, "High"]), "idx": i, "volume_score": 0})
    # Latest low can be useful for support risk, but latest high is not a confirmed resistance.
    if len(recent) >= 2:
        i = len(recent) - 1
        swing_lows.append({"price": float(recent.loc[i, "Low"]), "idx": i, "volume_score": 0})

    support_clusters = _cluster_price_levels(swing_lows, tolerance_pct=1.35)
    resistance_clusters = _cluster_price_levels(swing_highs, tolerance_pct=1.35)

    def strength(c):
        recency = c["last_idx"] / max(1, len(recent)-1)
        return c["touches"] * 2.0 + c.get("volume_score", 0) * 0.5 + recency

    supports = [c for c in support_clusters if c["price"] < close * 0.995]
    raw_resistances = [c for c in resistance_clusters if c["price"] > close * 1.005]

    # Prefer nearby strong zones, not the absolute min/max.
    supports.sort(key=lambda c: (abs(close - c["price"]) / close * 100 - strength(c) * 0.15))
    raw_resistances.sort(key=lambda c: (abs(c["price"] - close) / close * 100 - strength(c) * 0.15))

    # Manual chart rule from AMPX / ONDS examples:
    # A small level just above current close should NOT be the main resistance if today's high
    # already broke above it but the candle closed back below. That is a failed intraday pivot.
    latest_high = float(recent["High"].iloc[-1])
    failed_pivots = []
    resistances = []
    for c in raw_resistances:
        price = float(c["price"])
        broke_intraday = latest_high >= price * 1.002
        failed_to_hold = close < price * 0.998
        too_close_minor = (price - close) / max(close, 0.01) * 100 <= 2.5
        if broke_intraday and failed_to_hold and too_close_minor:
            failed_pivots.append(c)
            continue
        resistances.append(c)

    # If every nearby resistance was filtered, still keep higher raw resistance candidates.
    if not resistances:
        resistances = [c for c in raw_resistances if c not in failed_pivots] or raw_resistances

    # Main resistance rule: Target 1 should be a meaningful resistance ABOVE current price,
    # not a tiny pivot only 1-3% above price. If the nearest level is too close and there is
    # a higher resistance candidate, promote the higher level to Resistance 1.
    # This matches the manual ONDS logic: 10.42/10.60 = minor pivot, 11.02 = main resistance.
    if len(resistances) >= 2:
        first_price = float(resistances[0]["price"])
        first_dist = (first_price - close) / max(close, 0.01) * 100
        if first_dist < 4.0:
            failed_pivots.append(resistances[0])
            resistances = resistances[1:]

    support = supports[0]["price"] if supports else float(recent["Low"].tail(30).min())
    deep_support_candidates = [c for c in supports if c["price"] < support * 0.985]
    deep_support_candidates.sort(key=lambda c: (abs(support - c["price"]) / support * 100 - strength(c) * 0.10))
    deep_support = deep_support_candidates[0]["price"] if deep_support_candidates else min(support * 0.955, float(recent["Low"].tail(60).min()))
    resistance = resistances[0]["price"] if resistances else float(recent["High"].tail(60).max())
    failed_pivot = failed_pivots[0]["price"] if failed_pivots else None

    # Resistance 2 = stronger / higher resistance above Resistance 1.
    resistance_2_candidates = [c for c in resistances if c["price"] > resistance * 1.015]
    resistance_2_candidates.sort(key=lambda c: (abs(c["price"] - resistance) / max(resistance, 0.01) * 100 - strength(c) * 0.10))
    if resistance_2_candidates:
        resistance_2 = resistance_2_candidates[0]["price"]
    else:
        # Fallback: measured move above Resistance 1, similar to manual target zone when no clear line exists.
        resistance_2 = max(resistance * 1.06, close * 1.12)

    # Round later in display; keep raw for calculations.
    return {
        "support": float(support),
        "deep_support": float(deep_support),
        "resistance": float(resistance),
        "resistance_2": float(resistance_2),
        "failed_intraday_pivot": float(failed_pivot) if failed_pivot else None,
        "method": "standard_zone_formula_v26_support_resistance_entry",
    }


def find_support_resistance(data, lookback=180):
    levels = find_chart_trade_levels(data, lookback=lookback)
    return levels["support"], levels["resistance"]


def find_recent_swing_support(data, lookback=60):
    return find_chart_trade_levels(data, lookback=max(60, lookback))["deep_support"]


def analyze_single_market_index(ticker):
    try:
        data = download_price_data(ticker, period="6mo")
        if data.empty or len(data) < 60:
            return {"Ticker": ticker, "Condition": "Unknown"}

        data["EMA21"] = data["Close"].ewm(span=21, adjust=False).mean()
        data["EMA50"] = data["Close"].ewm(span=50, adjust=False).mean()
        data["RSI"] = calculate_rsi(data)

        latest = data.iloc[-1]
        close = float(latest["Close"])
        ema21 = float(latest["EMA21"])
        ema50 = float(latest["EMA50"])
        rsi = float(latest["RSI"])

        if close > ema21 > ema50 and rsi < 75:
            condition = "Bullish"
        elif close > ema50:
            condition = "Neutral"
        else:
            condition = "Bearish"

        return {"Ticker": ticker, "Condition": condition, "Close": close, "RSI": rsi}
    except Exception:
        return {"Ticker": ticker, "Condition": "Unknown"}


def get_market_condition():
    spy = analyze_single_market_index("SPY")
    qqq = analyze_single_market_index("QQQ")
    vix = analyze_single_market_index("^VIX")

    score = 0
    notes = []

    for item in [spy, qqq]:
        condition = item.get("Condition", "Unknown")
        ticker = item.get("Ticker", "")
        if condition == "Bullish":
            score += 2
            notes.append(f"{ticker} bullish")
        elif condition == "Neutral":
            score += 1
            notes.append(f"{ticker} neutral")
        elif condition == "Bearish":
            score -= 2
            notes.append(f"{ticker} bearish")
        else:
            notes.append(f"{ticker} unknown")

    try:
        vix_close = float(vix.get("Close", 0))
        if vix_close < 18:
            score += 1
            notes.append("VIX low")
        elif vix_close <= 25:
            notes.append("VIX normal")
        else:
            score -= 2
            notes.append("VIX high risk")
    except Exception:
        notes.append("VIX unknown")

    if score >= 3:
        market = "Bullish"
    elif score >= 0:
        market = "Neutral"
    else:
        market = "Bearish"

    return market, "; ".join(notes)


def normalize_earnings_date(value):
    """Convert many possible earnings-date formats to a pandas Timestamp or None."""
    if value is None:
        return None
    try:
        if isinstance(value, (list, tuple)) and len(value) > 0:
            value = value[0]
        if isinstance(value, pd.Series):
            value = value.dropna().iloc[0] if len(value.dropna()) else None
        if value is None:
            return None
        ts = pd.to_datetime(value, errors="coerce")
        if pd.isna(ts):
            return None
        return ts.tz_localize(None) if getattr(ts, "tzinfo", None) is not None else ts
    except Exception:
        return None


def get_nasdaq_earnings_date(ticker):
    """Backup earnings lookup from Nasdaq's public calendar endpoint. No API key required, but may fail."""
    try:
        url = f"https://api.nasdaq.com/api/calendar/earnings?symbol={ticker.upper()}"
        headers = {
            "User-Agent": "Mozilla/5.0",
            "Accept": "application/json, text/plain, */*",
            "Origin": "https://www.nasdaq.com",
            "Referer": "https://www.nasdaq.com/market-activity/earnings",
        }
        r = requests.get(url, headers=headers, timeout=12)
        if r.status_code != 200:
            return None
        data = r.json()
        rows = data.get("data", {}).get("rows", []) or []
        candidates = []
        for row in rows:
            raw_date = row.get("date") or row.get("reportDate") or row.get("time")
            ts = normalize_earnings_date(raw_date)
            if ts is not None:
                candidates.append(ts)
        if not candidates:
            return None
        today = pd.Timestamp.today().normalize()
        future = sorted([d for d in candidates if d.normalize() >= today])
        return future[0] if future else sorted(candidates)[-1]
    except Exception:
        return None


def get_earnings_warning(ticker):
    info = get_earnings_status(ticker)
    status = info.get("status", "Unknown")
    days = info.get("days_until")
    if status == "ETF / No Earnings":
        return "ETF / No Earnings"
    if status == "High Risk":
        return f"High Risk: Earnings in {days} days"
    if status == "Caution":
        return f"Watch: Earnings in {days} days"
    if status == "Clear":
        return f"OK: Earnings in {days} days"
    if status == "Reported":
        return "Past/Unknown"
    return "Unknown"


def get_earnings_status(ticker):
    """
    Multi-source earnings lookup:
    1) yfinance get_earnings_dates
    2) yfinance calendar
    3) Nasdaq public earnings calendar backup
    """
    if is_etf_symbol(ticker):
        return {
            "status": "ETF / No Earnings",
            "date": "N/A",
            "days_until": None,
            "note": "ETF/Fund - no company earnings report",
            "source": "ETF filter"
        }

    earnings_date = None
    source = "Unknown"

    try:
        stock = yf.Ticker(ticker)

        try:
            ed = stock.get_earnings_dates(limit=12)
            if ed is not None and not ed.empty:
                today = pd.Timestamp.today(tz=ed.index.tz) if ed.index.tz is not None else pd.Timestamp.today()
                future_dates = ed[ed.index >= today]
                earnings_date = future_dates.index[0] if not future_dates.empty else ed.index[0]
                source = "yfinance earnings_dates"
        except Exception:
            pass

        if earnings_date is None:
            try:
                cal = stock.calendar
                if cal is not None and len(cal) > 0:
                    possible = None
                    if isinstance(cal, dict):
                        possible = cal.get("Earnings Date") or cal.get("EarningsDate")
                    elif hasattr(cal, "index"):
                        for idx in cal.index:
                            if "Earnings" in str(idx):
                                value = cal.loc[idx]
                                possible = value.dropna().iloc[0] if hasattr(value, "dropna") else value
                                break
                    ts = normalize_earnings_date(possible)
                    if ts is not None:
                        earnings_date = ts
                        source = "yfinance calendar"
            except Exception:
                pass

        if earnings_date is None:
            ts = get_nasdaq_earnings_date(ticker)
            if ts is not None:
                earnings_date = ts
                source = "Nasdaq backup"

        if earnings_date is None:
            return {"status": "Unknown", "date": "Unknown", "days_until": None, "note": "No confirmed earnings date", "source": source}

        earnings_date = normalize_earnings_date(earnings_date)
        if earnings_date is None:
            return {"status": "Unknown", "date": "Unknown", "days_until": None, "note": "Earnings date unreadable", "source": source}

        today = pd.Timestamp.today().normalize()
        days_until = (earnings_date.normalize() - today).days
        date_text = earnings_date.strftime("%Y-%m-%d")

        if 0 <= days_until <= 7:
            status = "High Risk"
            note = f"Earnings in {days_until} day(s)"
        elif 8 <= days_until <= 14:
            status = "Caution"
            note = f"Earnings soon: {days_until} day(s)"
        elif days_until < 0:
            status = "Reported"
            note = f"Last earnings was {abs(days_until)} day(s) ago"
        else:
            status = "Clear"
            note = f"Earnings is {days_until} day(s) away"

        return {"status": status, "date": date_text, "days_until": days_until, "note": f"{note} | Source: {source}", "source": source}

    except Exception as e:
        return {"status": "Unknown", "date": "Unknown", "days_until": None, "note": f"Earnings data unavailable: {str(e)[:80]}", "source": source}


def safe_text(value):
    if value is None:
        return ""
    return str(value)


def simple_news_sentiment(text):
    text_l = text.lower()

    positive_words = [
        "contract", "partnership", "deal", "award", "beats", "beat", "raises guidance",
        "upgrade", "upgraded", "buy rating", "outperform", "launch", "approval",
        "record revenue", "strong demand", "expands", "growth", "collaboration"
    ]

    negative_words = [
        "offering", "dilution", "downgrade", "downgraded", "lawsuit", "investigation",
        "probe", "misses", "missed", "cuts guidance", "bankruptcy", "resignation",
        "sec investigation", "short seller", "fraud", "restatement", "going concern",
        "recall", "layoffs", "termination"
    ]

    pos = sum(1 for w in positive_words if w in text_l)
    neg = sum(1 for w in negative_words if w in text_l)

    if neg >= 2:
        return "Negative"
    if pos > neg:
        return "Positive"
    if neg > pos:
        return "Negative"

    return "Neutral"


def fetch_rss_titles(url, source_name, max_items=5):
    """Read simple RSS/Atom feeds and return recent news titles."""
    titles = []
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        r = requests.get(url, headers=headers, timeout=10)
        if r.status_code != 200 or not r.text:
            return []
        root = ET.fromstring(r.content)
        for item in root.findall(".//item")[:max_items]:
            title = item.findtext("title") or ""
            if title:
                titles.append({"title": safe_text(title), "source": source_name})
        if not titles:
            ns = {"atom": "http://www.w3.org/2005/Atom"}
            for entry in root.findall(".//atom:entry", ns)[:max_items]:
                title = entry.findtext("atom:title", default="", namespaces=ns)
                if title:
                    titles.append({"title": safe_text(title), "source": source_name})
    except Exception:
        return []
    return titles


def classify_news_risk_from_titles(titles):
    text = " | ".join([t.get("title", "") for t in titles]).lower()
    if not text:
        return "Neutral", "Unknown"

    high_risk_phrases = [
        "offering", "shelf offering", "registered direct", "public offering", "dilution",
        "bankruptcy", "going concern", "fraud", "sec investigation", "investigation",
        "lawsuit", "class action", "downgrade", "cuts guidance", "misses estimates",
        "resignation", "restatement", "short seller"
    ]
    positive_phrases = [
        "contract", "partnership", "deal", "award", "beats", "beat estimates",
        "raises guidance", "upgrade", "upgraded", "buy rating", "outperform",
        "launch", "approval", "record revenue", "strong demand", "expands", "growth"
    ]

    neg = sum(1 for p in high_risk_phrases if p in text)
    pos = sum(1 for p in positive_phrases if p in text)

    if neg >= 2:
        return "Negative", "High"
    if neg == 1 and pos == 0:
        return "Negative", "Medium"
    if pos > neg:
        return "Positive", "Low"
    if neg > pos:
        return "Negative", "Medium"
    return "Neutral", "Medium"


def get_yahoo_news_signal(ticker):
    """Multi-source news check: yfinance/Yahoo + Yahoo RSS + Google News RSS."""
    all_titles = []
    try:
        stock = yf.Ticker(ticker)
        news = stock.news or []
        now_ts = datetime.now().timestamp()
        for item in news[:10]:
            title = ""
            days_old = 999
            if isinstance(item, dict):
                title = safe_text(item.get("title", ""))
                content = item.get("content", {}) if isinstance(item.get("content"), dict) else {}
                if title == "":
                    title = safe_text(content.get("title", "") or content.get("headline", ""))
                publish_time = item.get("providerPublishTime") or content.get("pubDate")
                try:
                    if isinstance(publish_time, str):
                        publish_dt = pd.to_datetime(publish_time)
                        days_old = (datetime.now() - publish_dt.to_pydatetime().replace(tzinfo=None)).days
                    elif publish_time:
                        days_old = (now_ts - float(publish_time)) / 86400
                except Exception:
                    days_old = 999
            if title and days_old <= 30:
                all_titles.append({"title": title, "source": "Yahoo/yfinance"})
    except Exception:
        pass

    yahoo_rss = f"https://feeds.finance.yahoo.com/rss/2.0/headline?s={ticker.upper()}&region=US&lang=en-US"
    all_titles.extend(fetch_rss_titles(yahoo_rss, "Yahoo RSS", max_items=5))

    google_query = quote_plus(f"{ticker.upper()} stock OR shares")
    google_rss = f"https://news.google.com/rss/search?q={google_query}&hl=en-US&gl=US&ceid=US:en"
    all_titles.extend(fetch_rss_titles(google_rss, "Google News RSS", max_items=5))

    seen = set()
    unique = []
    for item in all_titles:
        title = re.sub(r"\s+", " ", item.get("title", "")).strip()
        key = title.lower()
        if title and key not in seen:
            seen.add(key)
            unique.append({"title": title, "source": item.get("source", "News")})

    if not unique:
        return {"News Sentiment": "Neutral", "News Risk": "Unknown", "Recent Catalyst": "No readable news found from Yahoo/Google RSS"}

    sentiment, news_risk = classify_news_risk_from_titles(unique[:8])
    catalyst = " | ".join([f"{x['source']}: {x['title']}" for x in unique[:3]])[:700]
    return {"News Sentiment": sentiment, "News Risk": news_risk, "Recent Catalyst": catalyst}


def get_sec_cik_map():
    cache_file = "sec_company_tickers_cache.json"

    if os.path.exists(cache_file):
        try:
            file_age = datetime.now() - datetime.fromtimestamp(os.path.getmtime(cache_file))
            if file_age.days < 7:
                with open(cache_file, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass

    headers = {
        "User-Agent": SEC_USER_AGENT,
        "Accept-Encoding": "gzip, deflate",
        "Host": "www.sec.gov"
    }

    url = "https://www.sec.gov/files/company_tickers.json"
    r = requests.get(url, headers=headers, timeout=15)
    r.raise_for_status()
    data = r.json()

    ticker_map = {}
    for _, item in data.items():
        ticker = safe_text(item.get("ticker", "")).upper()
        cik = int(item.get("cik_str"))
        title = safe_text(item.get("title", ""))
        ticker_map[ticker] = {"cik": cik, "title": title}

    with open(cache_file, "w", encoding="utf-8") as f:
        json.dump(ticker_map, f)

    return ticker_map


def get_sec_filing_signal(ticker):
    """
    Smarter SEC filter. High risk only for recent financing/dilution, resale, warrants,
    going-concern, restatement, investigation, or bankruptcy-type filings.
    """
    if is_etf_symbol(ticker):
        return {
            "SEC Filing Risk": "ETF / Not Applicable",
            "Recent SEC Filing": "ETF/Fund - SEC company filing risk not applicable"
        }

    try:
        ticker_map = get_sec_cik_map()
        info = ticker_map.get(ticker.upper())
        if not info:
            return {"SEC Filing Risk": "Unknown", "Recent SEC Filing": "Ticker not found in SEC ticker list"}

        cik = str(info["cik"]).zfill(10)
        headers = {"User-Agent": SEC_USER_AGENT, "Accept-Encoding": "gzip, deflate", "Host": "data.sec.gov"}
        url = f"https://data.sec.gov/submissions/CIK{cik}.json"
        r = requests.get(url, headers=headers, timeout=15)
        r.raise_for_status()
        data = r.json()
        recent = data.get("filings", {}).get("recent", {})

        forms = recent.get("form", [])
        filing_dates = recent.get("filingDate", [])
        descriptions = recent.get("primaryDocDescription", [])
        today = datetime.now().date()
        recent_items = []

        for i in range(min(len(forms), 35)):
            form = safe_text(forms[i]).upper()
            filing_date_raw = safe_text(filing_dates[i]) if i < len(filing_dates) else ""
            desc = safe_text(descriptions[i]) if i < len(descriptions) else ""
            try:
                filing_date = datetime.strptime(filing_date_raw, "%Y-%m-%d").date()
                days_old = (today - filing_date).days
            except Exception:
                days_old = 999
            if days_old <= 60:
                recent_items.append({"form": form, "date": filing_date_raw, "desc": desc, "days_old": days_old})

        if not recent_items:
            return {"SEC Filing Risk": "Low", "Recent SEC Filing": "No major recent SEC filing in 60 days"}

        high_forms = {"S-1", "S-3", "F-1", "F-3", "424B3", "424B4", "424B5", "424B7", "424B8", "FWP"}
        medium_forms = {"8-K", "10-Q", "10-K", "6-K", "20-F"}
        low_forms = {"3", "4", "5", "DEF 14A", "DEFA14A", "SC 13G", "SC 13D"}
        high_keywords = ["offering", "shelf", "prospectus", "atm", "at-the-market", "warrant", "resale", "convertible", "senior notes", "going concern", "restatement", "investigation", "subpoena", "bankruptcy", "chapter 11", "delisting", "material weakness", "default", "termination"]
        medium_keywords = ["acquisition", "merger", "credit agreement", "debt", "financing", "strategic review", "restructuring", "layoff", "cost reduction"]

        risk_score = 0
        notes = []
        high_notes = []
        for item in recent_items[:10]:
            form = item["form"]
            desc = item["desc"]
            days_old = item["days_old"]
            text_l = f"{form} {desc}".lower()
            label = f"{item['date']} {form} {desc}".strip()
            notes.append(label)
            if form in high_forms and days_old <= 45:
                risk_score = max(risk_score, 3)
                high_notes.append(label)
            if any(k in text_l for k in high_keywords) and days_old <= 45:
                risk_score = max(risk_score, 3)
                high_notes.append(label)
            elif any(k in text_l for k in medium_keywords) and days_old <= 45:
                risk_score = max(risk_score, 2)
            elif form in medium_forms and days_old <= 21:
                risk_score = max(risk_score, 1)
            elif form in low_forms:
                risk_score = max(risk_score, 0)

        if risk_score >= 3:
            risk = "High"
            prefix = "High-risk SEC item: "
            body = " | ".join(high_notes[:3]) if high_notes else " | ".join(notes[:3])
        elif risk_score == 2:
            risk = "Medium"
            prefix = "Medium-risk SEC item: "
            body = " | ".join(notes[:4])
        elif risk_score == 1:
            risk = "Medium"
            prefix = "Normal recent SEC filing; review: "
            body = " | ".join(notes[:4])
        else:
            risk = "Low"
            prefix = "Recent SEC filings appear normal: "
            body = " | ".join(notes[:4])

        return {"SEC Filing Risk": risk, "Recent SEC Filing": (prefix + body)[:700]}
    except Exception as e:
        return {"SEC Filing Risk": "Unknown", "Recent SEC Filing": f"SEC check error: {str(e)[:150]}"}


def get_news_sec_summary(ticker):
    news = get_yahoo_news_signal(ticker)
    time.sleep(0.15)
    sec = get_sec_filing_signal(ticker)

    return {
        "News Sentiment": news.get("News Sentiment", "Unknown"),
        "News Risk": news.get("News Risk", "Unknown"),
        "Recent Catalyst": news.get("Recent Catalyst", ""),
        "SEC Filing Risk": sec.get("SEC Filing Risk", "Unknown"),
        "Recent SEC Filing": sec.get("Recent SEC Filing", "")
    }
def get_chart_confirmation_needed(final_decision, signal, setup, trend):
    d = str(final_decision).upper()
    s = str(signal).upper()

    if "NO TRADE" in d:
        return "Không vào lệnh; chỉ theo dõi hoặc chờ rủi ro giảm"
    if "BREAKOUT" in d:
        return "Xác nhận breakout: đóng trên resistance + volume mạnh"
    if "REVERSAL" in d:
        return "Xác nhận reversal: nến xanh giữ support / higher low"
    if "PULLBACK" in d:
        return "Chờ pullback về Entry Zone và có nến giữ support"
    if "WATCH TO ENTER" in d or "BUY WATCH" in s:
        return "Có thể canh vào nhưng phải xác nhận chart trước khi mua"
    if "WAIT" in d:
        return "Chờ setup rõ hơn; chưa mua ngay"
    return "Luôn kiểm tra chart trước khi vào lệnh"


def get_app_role(final_decision):
    return "Lọc mã + cảnh báo rủi ro + gợi ý vùng giá; không phải lệnh mua tự động"


def analyze_stock(ticker, market_condition="Neutral", chart_confirmations=None):
    if chart_confirmations is None:
        chart_confirmations = {}
    try:
        data = download_price_data(ticker, period="8mo")
        if data.empty or len(data) < 80:
            return {"Ticker": ticker, "Status": "Không đủ dữ liệu"}

        data["EMA9"] = data["Close"].ewm(span=9, adjust=False).mean()
        data["EMA21"] = data["Close"].ewm(span=21, adjust=False).mean()
        data["EMA50"] = data["Close"].ewm(span=50, adjust=False).mean()
        data["EMA200"] = data["Close"].ewm(span=200, adjust=False).mean()
        data["RSI"] = calculate_rsi(data)
        data["MACD"], data["MACD_SIGNAL"], data["MACD_HIST"] = calculate_macd(data)
        data["ATR"] = calculate_atr(data)
        data["AVG_VOLUME_20"] = data["Volume"].rolling(20).mean()

        latest = data.iloc[-1]
        previous = data.iloc[-2]

        close = float(latest["Close"])
        prev_close = float(previous["Close"])
        day_change_pct = ((close - prev_close) / prev_close) * 100 if prev_close > 0 else 0
        ema9 = float(latest["EMA9"])
        ema21 = float(latest["EMA21"])
        ema50 = float(latest["EMA50"])
        ema200 = float(latest["EMA200"])
        rsi = float(latest["RSI"])
        macd_hist = float(latest["MACD_HIST"])
        prev_macd_hist = float(previous["MACD_HIST"])
        volume = float(latest["Volume"])
        avg_volume = float(latest["AVG_VOLUME_20"])
        atr = float(latest["ATR"])

        chart_levels = find_chart_trade_levels(data, lookback=180)
        support = chart_levels["support"]
        support_deep = chart_levels["deep_support"]
        resistance = chart_levels["resistance"]
        resistance_2 = chart_levels.get("resistance_2", max(resistance * 1.06, close * 1.12))
        failed_intraday_pivot = chart_levels.get("failed_intraday_pivot")
        swing_support = support_deep

        score = 0
        reasons = []
        if failed_intraday_pivot:
            reasons.append(f"Pivot nhỏ {failed_intraday_pivot:.2f} đã bị phá trong phiên nhưng chưa giữ được khi đóng cửa; dùng resistance chính phía trên")

        chart_data = chart_confirmations.get(ticker.upper(), {})
        chart_confirmation = chart_data.get("confirmation", "Neutral")
        chart_note = chart_data.get("note", "")
        if close > ema9 > ema21 > ema50:
            score += 3
            trend = "Strong Uptrend"
            reasons.append("Giá trên EMA9/21/50")
        elif close > ema9 > ema21:
            score += 2
            trend = "Uptrend"
            reasons.append("Giá trên EMA9 và EMA21")
        elif close > ema21 and close > ema50:
            score += 1
            trend = "Neutral Up"
            reasons.append("Giá trên EMA21 và EMA50")
        elif close < ema21 and close < ema50:
            score -= 3
            trend = "Downtrend"
            reasons.append("Giá dưới EMA21 và EMA50")
        else:
            score -= 1
            trend = "Mixed"
            reasons.append("Xu hướng chưa rõ")

        if close > ema200:
            score += 1
            reasons.append("Giá trên EMA200")
        else:
            score -= 1
            reasons.append("Giá dưới EMA200")

        if 45 <= rsi <= 65:
            score += 2
            rsi_status = "Healthy"
            reasons.append("RSI đẹp")
        elif 65 < rsi <= 70:
            score += 1
            rsi_status = "Warm"
            reasons.append("RSI hơi nóng")
        elif 30 <= rsi < 45:
            score += 1
            rsi_status = "Low"
            reasons.append("RSI thấp, có thể tích lũy")
        elif rsi > 70:
            score -= 2
            rsi_status = "Overbought"
            reasons.append("RSI quá mua")
        else:
            score -= 1
            rsi_status = "Oversold"
            reasons.append("RSI quá bán")

        if macd_hist > 0 and macd_hist > prev_macd_hist:
            score += 2
            momentum = "Strong"
            reasons.append("MACD đang mạnh lên")
        elif macd_hist > 0:
            score += 1
            momentum = "Positive"
            reasons.append("MACD dương")
        elif macd_hist > prev_macd_hist:
            momentum = "Improving"
            reasons.append("MACD âm nhưng cải thiện")
        else:
            score -= 1
            momentum = "Weak"
            reasons.append("MACD yếu")

        if avg_volume > 0 and volume > avg_volume * 1.5:
            score += 2
            volume_status = "Strong Volume"
            reasons.append("Volume mạnh")
        elif avg_volume > 0 and volume > avg_volume:
            score += 1
            volume_status = "Above Average"
            reasons.append("Volume trên trung bình")
        else:
            volume_status = "Normal/Weak"
            reasons.append("Volume chưa mạnh")

        distance_to_support = ((close - support) / close) * 100
        distance_to_resistance = ((resistance - close) / close) * 100

        if distance_to_support <= 4:
            score += 2
            setup = "Near Support"
            reasons.append("Giá gần support")
        elif distance_to_resistance <= 4:
            score -= 2
            setup = "Near Resistance"
            reasons.append("Giá gần resistance")
        else:
            setup = "Middle Zone"

        if market_condition == "Bullish":
            score += 1
            market_filter = "Supportive"
            reasons.append("Thị trường chung ủng hộ")
        elif market_condition == "Neutral":
            market_filter = "Neutral"
            reasons.append("Thị trường chung trung lập")
        else:
            score -= 2
            market_filter = "Risky"
            reasons.append("Thị trường chung rủi ro")
        # -------------------------------
        # News + SEC filter
        # -------------------------------
        news_sec = get_news_sec_summary(ticker)
        news_sentiment = news_sec["News Sentiment"]
        news_risk = news_sec["News Risk"]
        recent_catalyst = news_sec["Recent Catalyst"]
        sec_filing_risk = news_sec["SEC Filing Risk"]
        recent_sec_filing = news_sec["Recent SEC Filing"]

        if news_sentiment == "Positive":
            score += 1
            reasons.append("News sentiment positive")
        elif news_sentiment == "Negative":
            score -= 2
            reasons.append("News sentiment negative")

        if news_risk == "High":
            score -= 2
            reasons.append("News risk high")
        elif news_risk == "Medium":
            score -= 1
            reasons.append("News risk medium")

        if sec_filing_risk == "High":
            score -= 3
            reasons.append("SEC filing risk high")
        elif sec_filing_risk == "Medium":
            score -= 1
            reasons.append("SEC filing risk medium")

        # -------------------------------
        # Entry Zones - STANDARD ZONE FORMULA v26 - Manual Chart Style Logic
        # -------------------------------
        # Mô phỏng cách vẽ chart của người dùng:
        # 1) Resistance gần nhất phía trên = vùng chốt lời / breakout line
        # 2) Support gần nhất phía dưới = vùng pullback entry đẹp
        # 3) Deep support = vùng an toàn hơn nếu giá điều chỉnh sâu
        # Không mua giữa vùng nếu giá không gần support và chưa breakout.

        support_distance_pct = ((close - support) / close) * 100 if close > 0 else 999
        resistance_distance_pct = ((resistance - close) / close) * 100 if close > 0 else 999
        range_position_pct = ((close - support) / max(resistance - support, 0.01)) * 100 if resistance > support else 50
        range_position_pct = max(0, min(100, range_position_pct))

        # Entry chính quanh support gần: ví dụ support 19.76 thì entry khoảng 19.55 - 20.05
        aggressive_entry_low = support * 0.99
        aggressive_entry_high = support * 1.015

        # Deep Safe Entry quanh support sâu hơn: ví dụ deep support 18.81 thì khoảng 18.62 - 19.00
        safe_entry_low = support_deep * 0.99
        safe_entry_high = support_deep * 1.01

        if support_distance_pct <= 3:
            entry_quality = "Near Support - Best Watch"
            reasons.append("Giá đang gần support, có thể canh nến xác nhận để vào")
        elif support_distance_pct <= 7:
            entry_quality = "Good Pullback Entry"
            reasons.append("Entry Zone nằm quanh support gần, risk/reward hợp lý nếu có nến xác nhận")
        elif support_distance_pct <= 12:
            entry_quality = "Wait Pullback"
            reasons.append("Giá đang cao hơn support, nên chờ pullback thay vì mua đuổi")
        else:
            entry_quality = "Entry Too Far"
            reasons.append("Giá cách support khá xa, mua ngay sẽ rủi ro hơn")


        # Manual chart rule: avoid buying middle zone unless breakout is confirmed.
        # 0-35% of range = near support; 35-70% = middle zone; 70%+ = near resistance.
        if 35 < range_position_pct < 70 and resistance_distance_pct > 3:
            reasons.append("Giá đang ở giữa vùng support/resistance, không phải điểm mua đẹp")
            if final_decision not in ["NO TRADE - SEC Filing Risk", "NO TRADE - Earnings Risk"]:
                final_decision = "WAIT / NO SETUP"
        elif range_position_pct >= 70:
            reasons.append("Giá đang ở nửa trên của range, cần tránh mua đuổi gần resistance")

        # Nếu giá đang sát resistance, không nâng entry lên theo giá hiện tại.
        if resistance_distance_pct <= 4:
            setup = "Near Resistance"
            score -= 2
            reasons.append("Giá gần resistance, ưu tiên chờ pullback hoặc breakout rõ ràng")

        # Breakout entry chỉ dùng khi đóng cửa vượt resistance với volume mạnh.
        breakout_entry = resistance * 1.005

        # Bảo vệ lỗi dữ liệu
        aggressive_entry_low = max(aggressive_entry_low, 0)
        aggressive_entry_high = max(aggressive_entry_high, aggressive_entry_low)
        safe_entry_low = max(safe_entry_low, 0)
        safe_entry_high = max(safe_entry_high, safe_entry_low)

        entry_mid = (aggressive_entry_low + aggressive_entry_high) / 2
        entry_distance_pct = ((close - entry_mid) / close) * 100

        if entry_distance_pct <= 0:
            entry_quality = "Below Entry"
        elif entry_distance_pct <= 3:
            entry_quality = "Near Entry"
        elif entry_distance_pct <= 8:
            entry_quality = "Good Pullback Entry"
        elif entry_distance_pct <= 12:
            entry_quality = "Acceptable Entry"
        else:
            entry_quality = "Entry Too Far"

        aggressive_stop = min(aggressive_entry_low * 0.965, support * 0.965, ema21 * 0.97)
        safe_stop = support_deep * 0.965

        if trend in ["Strong Uptrend", "Uptrend", "Neutral Up"]:
            stop_loss = aggressive_stop
        else:
            stop_loss = safe_stop

        atr_stop = close - (atr * 1.5)
        stop_loss = min(stop_loss, atr_stop)

        risk = close - stop_loss
        if risk <= 0:
            target_1 = resistance
            target_2 = resistance
            risk_reward = 0
        else:
            rr_target_1 = close + risk * 1.5
            rr_target_2 = close + risk * 2.5

            # Manual chart style target logic:
            # Target 1 = Resistance 1 / nearest resistance.
            # Target 2 = Resistance 2 / stronger resistance, or measured R/R target if higher.
            if resistance > close:
                target_1 = resistance
                target_2 = max(resistance_2, rr_target_1)
            else:
                target_1 = rr_target_1
                target_2 = max(rr_target_2, resistance_2)

            if target_2 < target_1:
                target_1, target_2 = target_2, target_1

            # Guardrail for long trades: Target 1 must be meaningfully ABOVE current price.
            # If a minor pivot/support slipped into Target 1, replace it with the next resistance / RR target.
            min_valid_t1 = close * 1.035  # Target 1 for LONG must be above current price by a meaningful margin
            if target_1 < min_valid_t1:
                old_target_1 = target_1
                target_1 = max(resistance_2 if resistance_2 > min_valid_t1 else 0, rr_target_1, min_valid_t1)
                target_2 = max(target_2, target_1 * 1.04, rr_target_2)
                reasons.append(f"Target 1 {old_target_1:.2f} quá gần/thấp so với giá hiện tại; dùng kháng cự chính phía trên làm target")

            reward = target_2 - close
            risk_reward = reward / risk if risk > 0 else 0

        earnings_warning = get_earnings_warning(ticker)
        if "High Risk" in earnings_warning:
            score -= 2
            reasons.append("Earnings rất gần")
        elif "Watch" in earnings_warning:
            score -= 1
            reasons.append("Earnings sắp tới")

                # Reversal Watch logic
                # -------------------------------
        # Advanced Watch Logic
        # -------------------------------

        reversal_watch = False
        breakout_watch = False
        pullback_watch = False
        falling_knife = False

        # FALLING KNIFE FILTER:
        # Nếu giá đang rơi mạnh, nhất là trong downtrend/mixed trend, không bắt đáy ngay tại support.
        if (
            (day_change_pct <= -5)
            or (
                trend in ["Downtrend", "Mixed"]
                and day_change_pct <= -3
                and volume_status in ["Above Average", "Strong Volume", "Normal/Weak"]
            )
        ):
            falling_knife = True
            score -= 1
            reasons.append("Falling knife risk: giá đang rơi mạnh, cần nến xanh xác nhận trước khi vào")

        # REVERSAL WATCH:
        # Dành cho cổ phiếu Downtrend/Mixed nhưng có dấu hiệu hồi từ vùng thấp.
        if (
            market_condition != "Bearish"
            and trend in ["Downtrend", "Mixed"]
            and 35 <= rsi <= 58
            and momentum in ["Strong", "Positive", "Improving"]
            and setup in ["Middle Zone", "Near Support"]
            and chart_confirmation != "Weak / Breakdown"
            and score >= 0
        ):
            reversal_watch = True
            score += 1
            reasons.append("Reversal setup: cần nến xanh giữ support / higher low xác nhận")

        # BREAKOUT WATCH:
        # Dành cho cổ phiếu đang khá lên, market ủng hộ, momentum tốt,
        # nhưng chưa breakout rõ ràng hoặc volume chưa đủ mạnh.
        if (
            market_condition != "Bearish"
            and trend in ["Strong Uptrend", "Uptrend", "Neutral Up"]
            and 45 <= rsi <= 68
            and momentum in ["Strong", "Positive", "Improving"]
            and setup in ["Middle Zone", "Near Support"]
            and volume_status in ["Normal/Weak", "Above Average", "Strong Volume"]
            and score >= 2
        ):
            breakout_watch = True
            score += 1
            reasons.append("Đang hồi tốt nhưng cần breakout/volume xác nhận")

        # PULLBACK WATCH:
        # Dành cho cổ phiếu uptrend nhưng đang gần kháng cự hoặc hơi nóng,
        # không nên mua đuổi, nên chờ kéo về vùng vào.
        if (
            market_condition != "Bearish"
            and trend in ["Strong Uptrend", "Uptrend", "Neutral Up"]
            and setup == "Near Resistance"
            and rsi < 72
        ):
            pullback_watch = True
            reasons.append("Xu hướng tốt nhưng gần kháng cự, nên chờ pullback")
        # -------------------------------
        # Chart Confirmation Filter
        # -------------------------------

        chart_adjustment = 0

        if chart_confirmation == "Bullish Breakout":
            score += 2
            chart_adjustment = 2
            reasons.append("Chart xác nhận bullish breakout")

        elif chart_confirmation == "Pullback Holding":
            score += 1
            chart_adjustment = 1
            reasons.append("Chart xác nhận pullback đang giữ support")

        elif chart_confirmation == "Reversal Setup":
            score += 1
            chart_adjustment = 1
            reversal_watch = True
            reasons.append("Chart xác nhận reversal setup")

        elif chart_confirmation == "Near Resistance":
            score -= 1
            chart_adjustment = -1
            pullback_watch = True
            reasons.append("Chart đang gần resistance, không nên mua đuổi")

        elif chart_confirmation == "Weak / Breakdown":
            score -= 3
            chart_adjustment = -3
            reasons.append("Chart xác nhận breakdown/yếu")

        else:
            chart_confirmation = "Neutral"
            chart_adjustment = 0
        # -------------------------------
        # Final Signal
        # -------------------------------
        # Extra rule: Uptrend recovery should be Breakout Watch, not No Trade
        if (
            market_condition != "Bearish"
            and trend in ["Strong Uptrend", "Uptrend", "Neutral Up"]
            and 45 <= rsi <= 68
            and momentum in ["Strong", "Positive", "Improving"]
            and setup in ["Middle Zone", "Near Support"]
            and volume_status in ["Normal/Weak", "Above Average", "Strong Volume"]
            and chart_confirmation != "Weak / Breakdown"
        ):
            breakout_watch = True

            if score < 3:
                score = 3

            reasons.append("Recovery/uptrend setup: nên theo dõi breakout thay vì bỏ qua")
        if falling_knife:
            signal = "AVOID"
            action = "Support Test - chờ nến xanh xác nhận giữ support, không bắt dao rơi"
        elif score >= 8 and entry_quality in ["Near Entry", "Good Pullback Entry"] and market_condition != "Bearish":
            signal = "BUY WATCH"
            action = "Có thể canh vào nếu nến xác nhận"
        elif reversal_watch:
            signal = "WAIT"
            action = "Reversal Watch - chờ vượt kháng cự hoặc giữ support"
        elif breakout_watch:
            signal = "WAIT"
            action = "Breakout Watch - chờ vượt resistance hoặc volume xác nhận"
        elif pullback_watch:
            signal = "WAIT"
            action = "Pullback Watch - không mua đuổi, chờ giá kéo về"
        elif score >= 5 and market_condition != "Bearish":
            signal = "WAIT"
            action = "Đợi pullback hoặc breakout"
        else:
            signal = "AVOID"
            action = "Chưa nên vào"

        # -------------------------------
        # Final Decision
        # -------------------------------

        earnings_info = get_earnings_status(ticker)
        earnings_status = earnings_info["status"]
        earnings_date = earnings_info["date"]
        earnings_note = earnings_info["note"]

        if earnings_status == "High Risk":
            final_decision = "NO TRADE - Earnings Risk"
        elif earnings_status == "Caution":
            final_decision = "CAUTION - Earnings Soon"
        elif sec_filing_risk == "High":
            final_decision = "NO TRADE - SEC Filing Risk"
        elif chart_confirmation == "Weak / Breakdown":
            final_decision = "NO TRADE - Chart Weak"
        elif falling_knife:
            final_decision = "NO TRADE - Falling Knife"
        elif news_risk == "High" and signal != "BUY WATCH":
            final_decision = "NO TRADE - News Risk"
        elif market_condition == "Bearish" and signal != "BUY WATCH":
            final_decision = "NO TRADE - Market Weak"
        elif chart_confirmation == "Bullish Breakout" and score >= 6 and risk_reward >= 1.3:
            final_decision = "WATCH TO ENTER"
        elif chart_confirmation == "Pullback Holding" and score >= 5 and risk_reward >= 1.3:
            final_decision = "WATCH TO ENTER"
        elif signal == "BUY WATCH" and risk_reward >= 1.5:
            final_decision = "WATCH TO ENTER"
        elif reversal_watch:
            final_decision = "REVERSAL WATCH"
        elif breakout_watch:
            final_decision = "BREAKOUT WATCH"
        elif pullback_watch:
            final_decision = "PULLBACK WATCH"
        elif signal == "WAIT":
            final_decision = "WAIT FOR SETUP"
        else:
            final_decision = "BREAKOUT WATCH" if market_condition != "Bearish" and trend in ["Strong Uptrend", "Uptrend", "Neutral Up"] and 45 <= rsi <= 68 and momentum in ["Strong", "Positive", "Improving"] and setup in ["Middle Zone", "Near Support"] and chart_confirmation != "Weak / Breakdown" else "REVERSAL WATCH" if market_condition != "Bearish" and trend == "Downtrend" and 40 <= rsi <= 58 and momentum in ["Strong", "Positive", "Improving"] and volume_status in ["Above Average", "Strong Volume", "Normal/Weak"] and setup in ["Middle Zone", "Near Support"] and chart_confirmation != "Weak / Breakdown" else "WAIT / NO SETUP"     
            signal = "WAIT" if final_decision in ["BREAKOUT WATCH", "REVERSAL WATCH"] else signal
            action = "Breakout Watch - chờ vượt resistance hoặc volume xác nhận" if final_decision == "BREAKOUT WATCH" else "Reversal Watch - chờ vượt kháng cự hoặc giữ support" if final_decision == "REVERSAL WATCH" else action
            score = max(score, 3) if final_decision in ["BREAKOUT WATCH", "REVERSAL WATCH"] else score


        # Aggressive Entry = vùng vào sớm hơn, gần giá hiện tại hơn Entry Zone.
        # Entry Zone vẫn là vùng chính theo support. Deep Safe Entry là vùng pullback sâu hơn.
        if setup == "Near Resistance":
            early_entry_low = close * 0.95
            early_entry_high = close * 0.98
        elif aggressive_entry_low <= close <= aggressive_entry_high:
            early_entry_low = max(aggressive_entry_low, close * 0.985)
            early_entry_high = min(aggressive_entry_high, close * 1.01)
        elif close > aggressive_entry_high:
            early_entry_low = close * 0.975
            early_entry_high = close * 1.005
        else:
            early_entry_low = aggressive_entry_low
            early_entry_high = min(aggressive_entry_high, close * 1.01)

        early_entry_low = max(early_entry_low, 0)
        early_entry_high = max(early_entry_high, early_entry_low)

        return {
            "Ticker": ticker,
            "Price": round(close, 2),
            "Score": score,
            "Signal": signal,
            "Final Decision": final_decision,
            "Earnings Status": earnings_status,
            "Earnings Date": earnings_date,
            "Market Condition": market_condition,
            "Market Filter": market_filter,
            "Trend": trend,
            "Setup": setup,
            "Action": action,
            "RSI": round(rsi, 2),
            "Day Change %": round(day_change_pct, 2),
            "RSI Status": rsi_status,
            "Momentum": momentum,
            "Volume Status": volume_status,
            "Support": round(support, 2),
            "Deep Support": round(support_deep, 2),
            "Resistance": round(resistance, 2),
            "Resistance 2": round(resistance_2, 2),
            "Failed Intraday Pivot": round(failed_intraday_pivot, 2) if failed_intraday_pivot else "",
            "Breakout Entry": round(breakout_entry, 2),
            "Range Position %": round(range_position_pct, 2),
            "Entry Zone": f"{round(aggressive_entry_low, 2)} - {round(aggressive_entry_high, 2)}",
            "Deep Safe Entry": f"{round(safe_entry_low, 2)} - {round(safe_entry_high, 2)}",
            "Aggressive Entry": f"{round(early_entry_low, 2)} - {round(early_entry_high, 2)}",
            "Entry Distance %": round(entry_distance_pct, 2),
            "Entry Quality": entry_quality,
            "Aggressive Stop": round(aggressive_stop, 2),
            "Safe Stop": round(safe_stop, 2),
            "Stop Loss": round(stop_loss, 2),
            "Target 1": round(target_1, 2),
            "Target 2": round(target_2, 2),
            "Risk/Reward": round(risk_reward, 2),
            "Earnings Warning": earnings_warning,
            "Chart Confirmation": chart_confirmation,
            "Chart Adjustment": chart_adjustment,
            "Chart Note": chart_note,
            "News Sentiment": news_sentiment,
            "News Risk": news_risk,
            "Recent Catalyst": recent_catalyst,
            "SEC Filing Risk": sec_filing_risk,
            "Recent SEC Filing": recent_sec_filing,
            "Reasons": "; ".join(reasons),
            "Status": "OK"
        }
    except Exception as e:
        return {"Ticker": ticker, "Status": f"Lỗi: {str(e)}"}


def save_to_excel(df):
    columns = [
        "Ticker", "Price", "Score", "Signal", "Final Decision",
        "Earnings Status", "Earnings Date",
        "Market Condition", "Market Filter",
        "Trend", "Setup", "Action",
        "RSI", "Day Change %", "RSI Status", "Momentum", "Volume Status",
        "Support", "Deep Support", "Resistance", "Resistance 2", "Failed Intraday Pivot", "Range Position %",
        "Entry Zone", "Deep Safe Entry",
        "Aggressive Entry", "Entry Distance %",
        "Entry Quality",
        "Aggressive Stop", "Safe Stop", "Stop Loss",
        "Target 1", "Target 2", "Risk/Reward",
        "Chart Confirmation", "Chart Adjustment", "Chart Note",
        "News Sentiment", "News Risk", "Recent Catalyst",
        "SEC Filing Risk", "Recent SEC Filing",
        "Reasons", "Status"
    ]

    for col in columns:
        if col not in df.columns:
            df[col] = ""

    if os.path.exists(TEMPLATE_FILE) and load_workbook is not None:
        wb = load_workbook(TEMPLATE_FILE)
        ws = wb.active

        # Clear old headers/data first so removed columns do not remain in Excel
        clear_max_col = max(ws.max_column, len(columns) + 5)
        for row in ws.iter_rows(min_row=1, max_row=1000, min_col=1, max_col=clear_max_col):
            for cell in row:
                cell.value = None

        # Update header row so Excel columns match the new PC/iPhone logic
        for c_idx, col_name in enumerate(columns, start=1):
            ws.cell(row=1, column=c_idx).value = col_name

        

        template_row = 2
        for r_idx, row_data in enumerate(df[columns].values, start=2):
            for c_idx, value in enumerate(row_data, start=1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.value = value

                source_cell = ws.cell(row=template_row, column=c_idx)
                cell.font = copy(source_cell.font)
                cell.fill = copy(source_cell.fill)
                cell.border = copy(source_cell.border)
                cell.alignment = copy(source_cell.alignment)
                cell.number_format = source_cell.number_format
                cell.protection = copy(source_cell.protection)

        wb.save(OUTPUT_FILE)
    else:
        df[columns].to_excel(OUTPUT_FILE, index=False)





# ============================================================
# Streamlit / iPhone Web App UI - Stock App Pro PC-synced logic
# ============================================================

st.set_page_config(
    page_title="Stock App Pro",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed"
)

APP_COLUMNS = [
    "Ticker", "Price", "Score", "Signal", "Final Decision",
    "Earnings Status", "Earnings Date",
    "Market Condition", "Market Filter",
    "Trend", "Setup", "Action",
    "RSI", "Day Change %", "RSI Status", "Momentum", "Volume Status",
    "Support", "Deep Support", "Resistance", "Resistance 2", "Failed Intraday Pivot", "Range Position %",
    "Entry Zone", "Deep Safe Entry", "Aggressive Entry", "Entry Distance %",
    "Entry Quality", "Aggressive Stop", "Safe Stop", "Stop Loss",
    "Target 1", "Target 2", "Risk/Reward",
    "Chart Confirmation", "Chart Adjustment", "Chart Note",
    "News Sentiment", "News Risk", "Recent Catalyst",
    "SEC Filing Risk", "Recent SEC Filing",
    "Reasons", "Status", "Market Notes"
]

PREVIEW_COLUMNS = [
    "Ticker", "Price", "Score", "Final Decision", "Support", "Resistance",
    "Entry Zone", "Deep Safe Entry", "Breakout Entry", "Stop Loss", "Target 1", "Target 2", "Risk/Reward"
]

st.markdown(
    """
    <style>
    :root {
        --bg:#f4f7fb; --card:#ffffff; --text:#0f172a; --muted:#64748b;
        --line:#e2e8f0; --blue:#2563eb; --green:#16a34a; --yellow:#ca8a04; --red:#dc2626;
    }
    .main .block-container { padding-top: 1rem; padding-bottom: 3rem; max-width: 1180px; }
    .hero {
        background: linear-gradient(135deg, #7F1D1D 0%, #991B1B 58%, #7F1D1D 100%);
        color:white; border-radius:24px; padding:22px 20px; margin-bottom:14px;
        box-shadow:0 12px 30px rgba(15,23,42,.22);
    }
    .hero h1 { font-size:30px; margin:0 0 6px 0; line-height:1.1; }
    .hero p { margin:0; color:rgba(255,255,255,.86); font-size:15px; }
    .mini-note { background:#eff6ff; color:#1e3a8a; border:1px solid #bfdbfe; border-radius:16px; padding:10px 13px; margin:8px 0 12px 0; font-size:14px; }
    .section-title { font-size:21px; font-weight:800; margin:16px 0 8px 0; color:var(--text); }
    .panel-card { background:var(--card); border:1px solid var(--line); border-radius:20px; padding:14px 15px; box-shadow:0 8px 22px rgba(15,23,42,.05); margin:8px 0 12px 0; }
    .panel-title { font-size:16px; font-weight:850; color:var(--text); margin:0 0 8px 0; display:flex; align-items:center; gap:7px; }
    .summary-text { color:#1f2937; font-size:15px; line-height:1.6; }
    .bullet-list { margin:0; padding-left:18px; }
    .bullet-list li { margin:0 0 7px 0; color:#1f2937; line-height:1.5; }
    .level-line { margin:0 0 7px 0; color:#111827; line-height:1.5; }
    .level-line b { color:#0f172a; }
    .hint-box { background:#fff7ed; color:#9a3412; border:1px solid #fed7aa; border-radius:14px; padding:10px 12px; margin:8px 0 12px 0; font-size:14px; }
    .ai-hero-card { background:linear-gradient(135deg,#ffffff 0%,#f8fbff 100%); border:1px solid #dbe7fb; border-radius:22px; padding:14px 16px; box-shadow:0 10px 26px rgba(15,23,42,.06); margin:8px 0 10px 0; }
    .ai-main-reason { font-size:16px; line-height:1.6; color:#111827; margin-top:8px; }
    .ai-card-accent-blue { border-left:5px solid #2563eb; }
    .ai-card-accent-orange { border-left:5px solid #f97316; }
    .ai-card-accent-green { border-left:5px solid #16a34a; }
    .ai-card-accent-red { border-left:5px solid #dc2626; }
    .compact-section { margin-top:4px; }
    .pill-row { display:flex; flex-wrap:wrap; gap:8px; margin:8px 0 12px 0; }
    .pill { display:inline-block; padding:7px 11px; border-radius:999px; background:#f1f5f9; color:#334155; font-size:13px; font-weight:700; }
    .stock-card { background:var(--card); border:1px solid var(--line); border-radius:22px; padding:16px; margin-bottom:14px; box-shadow:0 8px 24px rgba(15,23,42,.06); }
    .stock-head { display:flex; justify-content:space-between; align-items:flex-start; gap:10px; margin-bottom:12px; }
    .ticker-title { font-size:25px; font-weight:850; color:var(--text); margin:0; letter-spacing:-.02em; }
    .price-line { color:var(--muted); font-size:14px; margin-top:2px; }
    .badge { display:inline-flex; align-items:center; justify-content:center; border-radius:999px; padding:7px 11px; font-size:12px; font-weight:850; white-space:nowrap; text-align:center; }
    .badge-good { background:#dcfce7; color:#166534; }
    .badge-watch { background:#fef3c7; color:#92400e; }
    .badge-bad { background:#fee2e2; color:#991b1b; }
    .badge-info { background:#dbeafe; color:#1e40af; }
    .badge-neutral { background:#f1f5f9; color:#334155; }
    .metric-grid { display:grid; grid-template-columns:repeat(2, minmax(0,1fr)); gap:10px; margin:10px 0 12px 0; }
    .metric-box { background:#f8fafc; border-radius:16px; padding:12px; border:1px solid #eef2f7; }
    .metric-label { color:var(--muted); font-size:12px; font-weight:800; text-transform:uppercase; letter-spacing:.03em; margin-bottom:4px; }
    .metric-value { color:var(--text); font-size:18px; font-weight:850; line-height:1.15; word-break:break-word; }
    .action-box { border-radius:16px; padding:13px 14px; background:#f8fafc; border-left:5px solid #2563eb; color:#0f172a; margin-top:8px; font-size:14px; }
    .small-muted { color:var(--muted); font-size:13px; }

    .chart-preview-fit { margin:10px 0 12px 0; padding:8px; background:#ffffff; border:1px solid var(--line); border-radius:18px; box-shadow:0 8px 22px rgba(15,23,42,.04); }
    .chart-preview-fit img { display:block; width:100%; height:auto; max-height:68vh; object-fit:contain; margin:0 auto; border-radius:14px; }
    .preview-caption { color:var(--muted); font-size:13px; text-align:center; margin-top:6px; }
    @media (orientation: landscape) and (max-height: 520px) {
        .main .block-container { padding-top:.35rem; }
        .hero { padding:12px 14px; margin-bottom:8px; }
        .hero h1 { font-size:22px; }
        .hero p, .mini-note { font-size:12px; }
        .chart-preview-fit img { max-height:74vh; width:auto; max-width:100%; }
        .section-title { font-size:18px; margin:10px 0 6px 0; }
    }
    div.stButton > button { border-radius:16px; min-height:48px; font-weight:850; }
    div[data-testid="stTextInput"] input, textarea { border-radius:14px !important; }
    @media (max-width: 640px) {
        .main .block-container { padding-left:.85rem; padding-right:.85rem; }
        .hero { border-radius:20px; padding:18px 16px; }
        .hero h1 { font-size:27px; }
        .metric-grid { grid-template-columns:1fr; }
        .panel-card { padding:13px 13px; border-radius:18px; }
        .stock-head { flex-direction:column; }
        .ticker-title { font-size:23px; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

def parse_tickers(text):
    tickers = []
    for raw in str(text).replace(",", "\n").replace(";", "\n").splitlines():
        t = raw.strip().upper().replace("$", "")
        if t:
            tickers.append(t)
    return list(dict.fromkeys(tickers))


def dataframe_to_excel_bytes(df):
    output = io.BytesIO()
    df2 = df.copy()
    for col in APP_COLUMNS:
        if col not in df2.columns:
            df2[col] = ""
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df2[APP_COLUMNS].to_excel(writer, index=False, sheet_name="Analysis")
    output.seek(0)
    return output.getvalue()


def load_default_text():
    return "\n".join(load_watchlist())


def badge_class(decision):
    d = str(decision).upper()
    if "WATCH TO ENTER" in d or "BUY WATCH" in d:
        return "badge-good"
    if "BREAKOUT" in d or "REVERSAL" in d or "PULLBACK" in d or "WAIT" in d or "CAUTION" in d:
        return "badge-watch"
    if "NO TRADE" in d or "AVOID" in d or "WEAK" in d or "RISK" in d or "FALLING" in d:
        return "badge-bad"
    return "badge-neutral"


def market_badge_class(market):
    m = str(market).upper()
    if "BULL" in m:
        return "badge-good"
    if "BEAR" in m:
        return "badge-bad"
    return "badge-info"


def safe_money(value):
    try:
        if value == "" or pd.isna(value):
            return "—"
        return f"${float(value):.2f}"
    except Exception:
        return str(value) if value not in [None, ""] else "—"


def safe_value(value, suffix=""):
    try:
        if value == "" or pd.isna(value):
            return "—"
    except Exception:
        pass
    return f"{value}{suffix}" if value not in [None, ""] else "—"


def render_stock_card(row):
    ticker = row.get("Ticker", "")
    decision = row.get("Final Decision", "")
    price = safe_money(row.get("Price", ""))
    score = safe_value(row.get("Score", ""))
    day_change = safe_value(row.get("Day Change %", ""), "%")
    action = row.get("Action", "")
    trend = row.get("Trend", "")
    setup = row.get("Setup", "")
    rsi = safe_value(row.get("RSI", ""))
    rr = safe_value(row.get("Risk/Reward", ""))
    entry = row.get("Entry Zone", "")
    deep_entry = row.get("Deep Safe Entry", "")
    aggressive_entry = row.get("Aggressive Entry", "")
    stop = safe_money(row.get("Stop Loss", ""))
    target1 = safe_money(row.get("Target 1", ""))
    target2 = safe_money(row.get("Target 2", ""))
    support = safe_money(row.get("Support", ""))
    resistance = safe_money(row.get("Resistance", ""))

    st.markdown(
        f"""
        <div class="stock-card">
            <div class="stock-head">
                <div>
                    <div class="ticker-title">{ticker}</div>
                    <div class="price-line">Price {price} · Score {score} · Day {day_change}</div>
                </div>
                <span class="badge {badge_class(decision)}">{decision}</span>
            </div>
            <div class="metric-grid">
                <div class="metric-box"><div class="metric-label">Entry Zone</div><div class="metric-value">{entry or '—'}</div></div>
                <div class="metric-box"><div class="metric-label">Stop Loss</div><div class="metric-value">{stop}</div></div>
                <div class="metric-box"><div class="metric-label">Target</div><div class="metric-value">{target1} / {target2}</div></div>
                <div class="metric-box"><div class="metric-label">Risk / Reward</div><div class="metric-value">{rr}</div></div>
            </div>
            <div class="pill-row">
                <span class="pill">Trend: {trend}</span>
                <span class="pill">Setup: {setup}</span>
                <span class="pill">RSI: {rsi}</span>
                <span class="pill">Support: {support}</span>
                <span class="pill">Resistance: {resistance}</span>
            </div>
            <div class="action-box"><b>Action:</b> {action or '—'}<br>
                <span class="small-muted">Aggressive Entry: {aggressive_entry or '—'}</span><br>
                <span class="small-muted">Deep Safe Entry: {deep_entry or '—'}</span><br>
                <span class="small-muted">Entry zone only. Check Final Decision and chart before buying.</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.expander(f"Chi tiết {ticker}: lý do, news, SEC, earnings"):
        earnings_status = row.get("Earnings Status", "")
        earnings_date = row.get("Earnings Date", "")
        if earnings_status or earnings_date:
            st.write(f"**Earnings:** {earnings_status} | {earnings_date}")
        if row.get("Chart Note", ""):
            st.write(f"**Chart note:** {row.get('Chart Note', '')}")
        if row.get("Recent Catalyst", ""):
            st.write(f"**News:** {row.get('Recent Catalyst', '')}")
        if row.get("Recent SEC Filing", ""):
            st.write(f"**SEC:** {row.get('Recent SEC Filing', '')}")
        if row.get("Reasons", ""):
            st.write("**Lý do phân tích:**")
            st.write(row.get("Reasons", ""))


def run_analysis_for_tickers(analysis_tickers):
    progress = st.progress(0)
    status = st.empty()
    status.info("Đang kiểm tra thị trường chung...")
    market_condition, market_notes = get_market_condition()
    chart_confirmations = load_chart_confirmations()
    results = []

    for i, ticker in enumerate(analysis_tickers, start=1):
        status.info(f"Đang phân tích {ticker}... ({i}/{len(analysis_tickers)})")
        result = analyze_stock(ticker, market_condition=market_condition, chart_confirmations=chart_confirmations)
        result["Market Notes"] = market_notes
        results.append(result)
        progress.progress(i / len(analysis_tickers))

    df = pd.DataFrame(results)
    if "Score" in df.columns:
        df = df.sort_values(by="Score", ascending=False)
    status.empty()
    progress.empty()
    return market_condition, market_notes, df


def get_openai_api_key():
    try:
        key = st.secrets.get("OPENAI_API_KEY", "")
    except Exception:
        key = ""
    if not key:
        key = st.session_state.get("openai_api_key", "")
    return str(key).strip()


def compact_app_data(row, market_notes=""):
    keys = [
        "Ticker", "Price", "Score", "Signal", "Final Decision",
        "Earnings Status", "Earnings Date", "Market Condition", "Market Filter",
        "Trend", "Setup", "Action", "RSI", "Day Change %", "RSI Status",
        "Momentum", "Volume Status", "Support", "Deep Support", "Resistance", "Resistance 2", "Breakout Entry", "Range Position %", "Entry Zone",
        "Deep Safe Entry", "Aggressive Entry", "Entry Distance %", "Entry Quality",
        "Aggressive Stop", "Safe Stop", "Stop Loss", "Target 1", "Target 2",
        "Risk/Reward", "Chart Confirmation", "Chart Adjustment", "Chart Note",
        "News Sentiment", "News Risk", "Recent Catalyst", "SEC Filing Risk",
        "Recent SEC Filing", "Reasons", "Status"
    ]
    data = {}
    for k in keys:
        v = row.get(k, "")
        try:
            if pd.isna(v):
                v = ""
        except Exception:
            pass
        data[k] = str(v)
    data["Market Notes"] = market_notes
    return data


AI_FINAL_SCHEMA = {
    "ai_final_decision": "WAIT / NO TRADE / WATCH TO ENTER / BREAKOUT WATCH / PULLBACK WATCH / REVERSAL WATCH / CAUTION",
    "confidence": "High / Medium / Low",
    "app_decision_review": "Agree / More cautious / More bullish / Mixed",
    "chart_confirmation": "Confirmed / Not confirmed / Weak / Unclear",
    "trade_status": "Active setup / Reference only / No trade",
    "main_reason": "Lý do chính ngắn gọn bằng tiếng Việt",
    "key_points": ["Điểm chính 1 bằng tiếng Việt", "Điểm chính 2 bằng tiếng Việt", "Điểm chính 3 bằng tiếng Việt"],
    "risk_warnings": ["Cảnh báo rủi ro 1 bằng tiếng Việt", "Cảnh báo rủi ro 2 bằng tiếng Việt"],
    "action_plan": ["Bước hành động 1 bằng tiếng Việt", "Bước hành động 2 bằng tiếng Việt", "Bước hành động 3 bằng tiếng Việt"],
    "final_entry_zone": "price zone. If decision is WAIT or NO TRADE, still provide Reference Entry Zone / Stop / Target from app data and label them Reference only",
    "final_stop_loss": "price. If decision is WAIT, still provide Reference Stop from app data and label it Reference only",
    "final_target_1": "price. If decision is WAIT, still provide Reference Target 1 from app data and label it Reference only",
    "final_target_2": "price. If decision is WAIT, still provide Reference Target 2 from app data and label it Reference only",
    "invalid_if": "Điều kiện làm setup mất hiệu lực, viết bằng tiếng Việt",
    "summary_vi": "Vietnamese practical summary"
}


def run_ai_final_analysis(app_data, image_bytes, mime_type, user_context, trading_style, api_key):
    client = OpenAI(api_key=api_key)
    b64 = base64.b64encode(image_bytes).decode("utf-8")

    system_prompt = """
You are the final analyst layer for a swing-trading stock app.
The app has already collected real data and calculated technical/risk levels.
Your job is NOT to replace the app calculations. Your job is to:
1) read the uploaded chart image,
2) compare the chart to the app data,
3) evaluate whether the app's chart-style levels make sense: Support, Deep Support, Resistance, Resistance 2, Entry Zone, Deep Safe Entry, Breakout Entry, Stop and Targets,
4) incorporate app-collected news, SEC, earnings, market condition, ETF filter, and falling-knife risk,
5) produce one practical final decision using this rule: prefer pullback entries near support, avoid buying in the middle zone, avoid chasing near resistance, and only accept breakout entries when chart/volume confirm.

Important rules:
- Use only the app data, chart image, and user context provided.
- Do not invent fresh news or unseen fundamentals.
- If chart quality is poor, lower confidence.
- If app data says earnings or SEC risk is high, be conservative.
- If price is near Support 1 and chart confirms support hold / green reversal / higher low, you may upgrade to WATCH TO ENTER or PULLBACK WATCH.
- If price is near support but the chart shows a red candle/falling knife/no confirmation, prefer WAIT or NO TRADE.
- If a minor resistance was broken intraday but price closed back below it, treat it as a failed intraday pivot, not the main resistance. Use the next stronger resistance above as Resistance 1 / Target 1.
- For LONG/BUY analysis, Target 1 must be above current price and should be the main resistance from the uploaded chart, not a support, not an intraday high, and not a tiny pivot that price already pierced intraday.
- If app data gives a Target 1 that is below current price or too close to current price, correct it using the uploaded chart's next clear resistance.
- If price is in the middle of Support 1 and Resistance 1, prefer WAIT / NO SETUP unless there is a very strong catalyst.
- If price is close to Resistance 1, do not chase; prefer WAIT unless price clearly breaks out with volume.
- If app data is bullish but chart shows rejection at resistance, prefer WAIT.
- If chart confirms breakout above Resistance 1 with strong volume, you may upgrade to BREAKOUT WATCH.
- Always give a practical action plan.
- Write ALL user-facing text in Vietnamese for these fields: main_reason, key_points, risk_warnings, action_plan, invalid_if, summary_vi.
- Keep only technical labels such as Entry Zone, Stop Loss, Target 1, Target 2 in English if needed.
- IMPORTANT: Never return N/A for Entry Zone / Stop Loss / Target 1 / Target 2 if app data provides those levels.
- Even when the final decision is NO TRADE, keep the app's Entry Zone, Stop Loss, Target 1, and Target 2 as REFERENCE ONLY levels for tracking.
- If final decision is NO TRADE, clearly label levels as "No trade - reference only, not active trade".
- If final decision is WAIT, CAUTION, BREAKOUT WATCH, PULLBACK WATCH, or REVERSAL WATCH, label levels as "Reference only - not active trade yet".
- If the chart is not confirmed, say "Chart not confirmed" but still show the app reference levels for tracking.
- Return valid JSON only with keys matching the provided schema.
"""

    user_prompt = f"""
Trading style: {trading_style}
User extra context: {user_context or 'None'}

APP DATA ENGINE OUTPUT:
{json.dumps(app_data, ensure_ascii=False, indent=2)}

Return JSON only using this schema:
{json.dumps(AI_FINAL_SCHEMA, ensure_ascii=False, indent=2)}
"""

    response = client.responses.create(
        model="gpt-4.1-mini",
        input=[
            {"role": "system", "content": [{"type": "input_text", "text": system_prompt}]},
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": user_prompt},
                    {"type": "input_image", "image_url": f"data:{mime_type};base64,{b64}"},
                ],
            },
        ],
    )
    raw = response.output_text.strip()
    try:
        return json.loads(raw)
    except Exception:
        cleaned = raw.replace("```json", "").replace("```", "").strip()
        return json.loads(cleaned)



def _is_missing_level(value):
    text = str(value or "").strip().lower()
    return text in ["", "—", "n/a", "na", "none", "null", "wait"]


def normalize_ai_result(result, app_data):
    """Keep reference levels visible when AI says WAIT/CAUTION/WATCH.

    This avoids showing Entry/Stop/Targets as N/A when the setup is not active yet,
    while still making clear that these are tracking levels, not a buy signal.
    """
    if not isinstance(result, dict):
        return result

    decision = str(result.get("ai_final_decision", "")).upper()
    no_trade = "NO TRADE" in decision
    active_trade = "WATCH TO ENTER" in decision

    entry_zone = app_data.get("Entry Zone") or app_data.get("Aggressive Entry") or ""
    deep_entry = app_data.get("Deep Safe Entry") or ""
    stop_loss = app_data.get("Stop Loss") or app_data.get("Safe Stop") or ""
    target_1 = app_data.get("Target 1") or ""
    target_2 = app_data.get("Target 2") or ""

    if active_trade:
        default_status = "Active setup"
        prefix = "Active setup"
    elif no_trade:
        default_status = "No trade"
        prefix = "No trade - reference only, not active trade"
    else:
        default_status = "Reference only"
        prefix = "Reference only - not active trade yet"

    if not result.get("trade_status"):
        result["trade_status"] = default_status

    # Always keep app-calculated levels visible as reference when available.
    # Even NO TRADE should show these levels for tracking, not as a buy signal.
    if _is_missing_level(result.get("final_entry_zone")):
        if deep_entry and deep_entry != entry_zone:
            result["final_entry_zone"] = f"{prefix}: Entry Zone {entry_zone}; Deep Safe Entry {deep_entry}"
        else:
            result["final_entry_zone"] = f"{prefix}: Entry Zone {entry_zone}"
    if _is_missing_level(result.get("final_stop_loss")):
        result["final_stop_loss"] = f"{prefix}: {stop_loss}"
    if _is_missing_level(result.get("final_target_1")):
        result["final_target_1"] = f"{prefix}: {target_1}"
    if _is_missing_level(result.get("final_target_2")):
        result["final_target_2"] = f"{prefix}: {target_2}"

    # If AI returned plain "Wait" as entry, make it clearer.
    if str(result.get("final_entry_zone", "")).strip().lower() == "wait":
        result["final_entry_zone"] = f"Vùng tham khảo - chưa phải lệnh mua: Entry Zone {entry_zone}; Deep Safe Entry {deep_entry}"

    # Make invalid_if practical if blank.
    if _is_missing_level(result.get("invalid_if")):
        result["invalid_if"] = f"Setup mất hiệu lực nếu giá thủng Stop Loss {stop_loss}, hoặc chart không có nến xác nhận tại Entry Zone."

    return result



def render_ai_result(result):
    def esc(v):
        return html.escape(str(v or "—"))

    def make_list(items):
        items = items or []
        if not items:
            return '<ul class="bullet-list"><li>—</li></ul>'
        li = ''.join(f'<li>{esc(x)}</li>' for x in items)
        return f'<ul class="bullet-list">{li}</ul>'

    decision = safe_value(result.get("ai_final_decision", "—"))
    confidence = safe_value(result.get("confidence", "—"))
    review = safe_value(result.get("app_decision_review", "—"))
    chart_conf = safe_value(result.get("chart_confirmation", "—"))
    trade_status = safe_value(result.get("trade_status", "—"))

    st.markdown(
        f'<div class="ai-hero-card">'
        f'<div class="pill-row">'
        f'<span class="badge {badge_class(decision)}">🤖 Kết luận: {esc(decision)}</span>'
        f'<span class="badge badge-info">Độ tin cậy: {esc(confidence)}</span>'
        f'<span class="badge badge-neutral">So với app: {esc(review)}</span>'
        f'<span class="badge badge-neutral">Chart: {esc(chart_conf)}</span>'
        f'<span class="badge badge-neutral">Trạng thái: {esc(trade_status)}</span>'
        f'</div>'
        f'<div class="ai-main-reason">{esc(result.get("main_reason", ""))}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    if str(trade_status).lower().startswith("reference"):
        st.markdown(
            '<div class="hint-box">📌 Các mức Entry / Stop / Target bên dưới chỉ là vùng tham khảo để theo dõi. Chưa phải lệnh mua vì AI chưa thấy chart xác nhận.</div>',
            unsafe_allow_html=True,
        )

    st.markdown('<div class="section-title compact-section">Kế hoạch hành động</div>', unsafe_allow_html=True)

    left_html = (
        '<div class="panel-card ai-card-accent-blue">'
        '<div class="panel-title">🎯 Vùng giá tham khảo</div>'
        f'<div class="level-line"><b>Entry:</b> {esc(result.get("final_entry_zone", "—"))}</div>'
        f'<div class="level-line"><b>Stop:</b> {esc(result.get("final_stop_loss", "—"))}</div>'
        f'<div class="level-line"><b>Target 1:</b> {esc(result.get("final_target_1", "—"))}</div>'
        f'<div class="level-line"><b>Target 2:</b> {esc(result.get("final_target_2", "—"))}</div>'
        f'<div class="level-line"><b>Mất hiệu lực nếu:</b> {esc(result.get("invalid_if", "—"))}</div>'
        '</div>'
    )
    right_html = (
        '<div class="panel-card ai-card-accent-green">'
        '<div class="panel-title">🧭 Tóm tắt</div>'
        f'<div class="summary-text">{esc(result.get("summary_vi", "—"))}</div>'
        '</div>'
    )

    c1, c2 = st.columns([1.05, 0.95])
    with c1:
        st.markdown(left_html, unsafe_allow_html=True)
    with c2:
        st.markdown(right_html, unsafe_allow_html=True)

    c3, c4, c5 = st.columns(3)
    with c3:
        st.markdown(
            f'<div class="panel-card ai-card-accent-blue"><div class="panel-title">✅ Điểm chính</div>{make_list(result.get("key_points", []))}</div>',
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            f'<div class="panel-card ai-card-accent-red"><div class="panel-title">⚠️ Cảnh báo rủi ro</div>{make_list(result.get("risk_warnings", []))}</div>',
            unsafe_allow_html=True,
        )
    with c5:
        st.markdown(
            f'<div class="panel-card ai-card-accent-orange"><div class="panel-title">➡️ Bước tiếp theo</div>{make_list(result.get("action_plan", []))}</div>',
            unsafe_allow_html=True,
        )


def render_uploaded_chart_preview(uploaded_file):
    """Show chart image in a responsive container that fits better on iPhone landscape."""
    if uploaded_file is None:
        return
    mime = uploaded_file.type or "image/png"
    b64 = base64.b64encode(uploaded_file.getvalue()).decode("utf-8")
    html_block = (
        f'<div class="chart-preview-fit">'
        f'<img src="data:{mime};base64,{b64}" alt="Uploaded chart" />'
        f'<div class="preview-caption">Uploaded chart - tự fit theo màn hình, nhất là khi xoay ngang iPhone</div>'
        f'</div>'
    )
    st.markdown(html_block, unsafe_allow_html=True)


def build_share_text(result, app_data=None):
    app_data = app_data or {}
    ticker = app_data.get("Ticker", "")
    price = app_data.get("Price", "")
    lines = []
    title = f"📈 AI Final Analysis {ticker}".strip()
    lines.append(title)
    if price:
        lines.append(f"Giá hiện tại: {price}")
    lines.append(f"Kết luận: {result.get('ai_final_decision', '—')}")
    lines.append(f"Độ tin cậy: {result.get('confidence', '—')}")
    lines.append(f"Chart: {result.get('chart_confirmation', '—')}")
    lines.append(f"Trạng thái: {result.get('trade_status', '—')}")
    lines.append("")
    lines.append("Lý do chính:")
    lines.append(str(result.get("main_reason", "—")))
    lines.append("")
    lines.append("Vùng giá tham khảo:")
    lines.append(f"Entry: {result.get('final_entry_zone', '—')}")
    lines.append(f"Stop: {result.get('final_stop_loss', '—')}")
    lines.append(f"Target 1: {result.get('final_target_1', '—')}")
    lines.append(f"Target 2: {result.get('final_target_2', '—')}")
    lines.append(f"Mất hiệu lực nếu: {result.get('invalid_if', '—')}")
    lines.append("")
    lines.append("Tóm tắt:")
    lines.append(str(result.get("summary_vi", "—")))
    lines.append("")
    lines.append("Điểm chính:")
    for item in result.get("key_points", []) or []:
        lines.append(f"- {item}")
    lines.append("")
    lines.append("Cảnh báo rủi ro:")
    for item in result.get("risk_warnings", []) or []:
        lines.append(f"- {item}")
    lines.append("")
    lines.append("Bước tiếp theo:")
    for item in result.get("action_plan", []) or []:
        lines.append(f"- {item}")
    lines.append("")
    lines.append("Ghi chú: Đây là phân tích tham khảo, không phải khuyến nghị tài chính.")
    return "\n".join(lines)


st.markdown(
    """
    <div class="hero">
        <h1>📈 Stock App Pro V26</h1>
        <p>V26 đồng bộ PC: App Data + Multi-Source News/Catalyst + Final Analysis.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="mini-note">
        Lọc mã → News/Catalyst Check → Phân tích chart → Final Decision.
    </div>
    """,
    unsafe_allow_html=True,
)

if "ticker_text" not in st.session_state:
    st.session_state.ticker_text = load_default_text()
if "quick_text" not in st.session_state:
    st.session_state.quick_text = ""
if "ai_result" not in st.session_state:
    st.session_state.ai_result = None
if "ai_app_data" not in st.session_state:
    st.session_state.ai_app_data = None
if "openai_api_key" not in st.session_state:
    st.session_state.openai_api_key = ""

tab_quick, tab_watchlist, tab_ai, tab_guide = st.tabs([
    "⚡ Quick Analyze", "📋 Watchlist", "📊 FINAL ANALYSIS", "📱 iPhone Guide"
])

with tab_quick:
    st.markdown('<div class="section-title">Phân tích nhanh mã mới</div>', unsafe_allow_html=True)
    st.caption("Gõ 1 mã hoặc nhiều mã. Ví dụ: HOOD hoặc NVDA, TSLA, SOUN")
    quick_text = st.text_input("Nhập ticker", value=st.session_state.quick_text, placeholder="HOOD, NVDA, TSLA", key="quick_input", label_visibility="collapsed")
    st.session_state.quick_text = quick_text
    quick_tickers = parse_tickers(quick_text)

    c1, c2 = st.columns([2, 1])
    with c1:
        quick_run = st.button("⚡ RUN QUICK ANALYSIS", type="primary", use_container_width=True)
    with c2:
        clear_quick = st.button("Clear", use_container_width=True)
    if clear_quick:
        st.session_state.quick_text = ""
        st.rerun()

    if quick_tickers:
        st.markdown("<div class='pill-row'>" + "".join([f"<span class='pill'>{t}</span>" for t in quick_tickers]) + "</div>", unsafe_allow_html=True)

    if quick_run:
        if not quick_tickers:
            st.error("Bạn chưa nhập ticker để phân tích nhanh.")
        else:
            try:
                market_condition, market_notes, df = run_analysis_for_tickers(quick_tickers)
                st.markdown(f"<span class='badge {market_badge_class(market_condition)}'>Market: {market_condition}</span>", unsafe_allow_html=True)
                st.caption(f"Market notes: {market_notes}")
                st.markdown('<div class="section-title">Kết quả</div>', unsafe_allow_html=True)
                for _, row in df.iterrows():
                    render_stock_card(row)
                with st.expander("📊 Bảng preview"):
                    cols = [c for c in PREVIEW_COLUMNS if c in df.columns]
                    st.dataframe(df[cols], use_container_width=True, hide_index=True)
                with st.expander("📋 Bảng chi tiết đầy đủ"):
                    st.dataframe(df, use_container_width=True, hide_index=True)
                st.download_button("⬇️ Download Excel", data=dataframe_to_excel_bytes(df), file_name="quick_analysis.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            except Exception as e:
                st.error(f"Lỗi: {e}")

with tab_watchlist:
    st.markdown('<div class="section-title">Watchlist chính</div>', unsafe_allow_html=True)
    st.caption("Danh sách này dùng cho những mã bạn theo dõi thường xuyên.")
    ticker_text = st.text_area("Watchlist", value=st.session_state.ticker_text, height=220, placeholder="ENVX\nSOUN\nWULF\nPATH")
    tickers = parse_tickers(ticker_text)

    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        if st.button("💾 Save list", use_container_width=True):
            save_watchlist(tickers)
            st.session_state.ticker_text = "\n".join(tickers)
            st.success("Đã lưu watchlist.")
    with c2:
        if st.button("↩️ Default", use_container_width=True):
            st.session_state.ticker_text = "\n".join(DEFAULT_WATCHLIST)
            st.rerun()
    with c3:
        run_watchlist = st.button("🚀 RUN WATCHLIST ANALYSIS", type="primary", use_container_width=True)

    st.markdown(f"<div class='small-muted'>Đang có {len(tickers)} mã trong watchlist.</div>", unsafe_allow_html=True)
    if tickers:
        st.markdown("<div class='pill-row'>" + "".join([f"<span class='pill'>{t}</span>" for t in tickers]) + "</div>", unsafe_allow_html=True)
    else:
        st.warning("Bạn chưa nhập ticker nào trong watchlist.")

    if run_watchlist:
        if not tickers:
            st.error("Bạn chưa nhập ticker trong watchlist.")
        else:
            try:
                market_condition, market_notes, df = run_analysis_for_tickers(tickers)
                st.markdown(f"<span class='badge {market_badge_class(market_condition)}'>Market: {market_condition}</span>", unsafe_allow_html=True)
                st.caption(f"Market notes: {market_notes}")
                st.markdown('<div class="section-title">Kết quả watchlist</div>', unsafe_allow_html=True)
                for _, row in df.iterrows():
                    render_stock_card(row)
                with st.expander("📊 Bảng preview"):
                    cols = [c for c in PREVIEW_COLUMNS if c in df.columns]
                    st.dataframe(df[cols], use_container_width=True, hide_index=True)
                with st.expander("📋 Bảng chi tiết đầy đủ"):
                    st.dataframe(df, use_container_width=True, hide_index=True)
                st.download_button("⬇️ Download Excel", data=dataframe_to_excel_bytes(df), file_name="watchlist_analysis.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            except Exception as e:
                st.error(f"Lỗi: {e}")

with tab_ai:
    st.markdown('<div class="section-title">FINAL ANALYSIS - 2 tầng</div>', unsafe_allow_html=True)
    st.caption("Tầng 1: app tự lấy dữ liệu và tính toán. Tầng 2: phân tích chart upload + dữ liệu app để đưa quyết định cuối cùng.")

    try:
        _secret_api_key = st.secrets.get("OPENAI_API_KEY", "")
    except Exception:
        _secret_api_key = ""

    if _secret_api_key:
        st.caption("API key đã được cấu hình an toàn trong Streamlit Secrets.")
    else:
        with st.expander("API key", expanded=False):
            st.write("Chỉ nhập key ở đây khi bạn chưa lưu OPENAI_API_KEY trong Streamlit Secrets.")
            st.session_state.openai_api_key = st.text_input(
                "OpenAI API Key",
                value=st.session_state.openai_api_key,
                type="password",
                placeholder="sk-...",
            )

    ai_ticker = st.text_input("Ticker cần phân tích", placeholder="VD: CIFR, KTOS, SOUN").upper().strip()
    trading_style = st.selectbox("Trading style", ["Swing trade", "Day trade", "Position trade", "Conservative swing", "Aggressive breakout"], index=0)
    uploaded_chart = st.file_uploader("Upload chart screenshot", type=["png", "jpg", "jpeg", "webp"])
    user_context = st.text_area("Context thêm nếu có", placeholder="VD: Tôi đang muốn vào swing 1-4 tuần, chờ pullback hoặc breakout xác nhận...", height=100)

    c1, c2 = st.columns(2)
    with c1:
        run_app_layer = st.button("1️⃣ Run App Data Engine", type="primary", use_container_width=True)
    with c2:
        run_ai_layer = st.button("2️⃣ Run Final Analysis", use_container_width=True)

    if uploaded_chart is not None:
        render_uploaded_chart_preview(uploaded_chart)

    if run_app_layer:
        if not ai_ticker:
            st.error("Bạn chưa nhập ticker.")
        else:
            try:
                with st.spinner("Đang chạy Tầng 1: App Data Engine..."):
                    market_condition, market_notes = get_market_condition()
                    chart_confirmations = load_chart_confirmations()
                    row = analyze_stock(ai_ticker, market_condition=market_condition, chart_confirmations=chart_confirmations)
                    row["Market Notes"] = market_notes
                    st.session_state.ai_app_data = compact_app_data(row, market_notes=market_notes)
                    st.session_state.ai_result = None
                st.success("Đã chạy xong Tầng 1. Kiểm tra dữ liệu bên dưới, rồi upload chart và bấm Final Analysis.")
            except Exception as e:
                st.error(f"Lỗi khi chạy App Data Engine: {e}")

    if st.session_state.ai_app_data:
        st.markdown('<div class="section-title">Tầng 1: App Data Engine Output</div>', unsafe_allow_html=True)
        app_data = st.session_state.ai_app_data
        preview_items = {
            "Ticker": app_data.get("Ticker"),
            "Price": app_data.get("Price"),
            "App Decision": app_data.get("Final Decision"),
            "Entry Zone": app_data.get("Entry Zone"),
            "Stop Loss": app_data.get("Stop Loss"),
            "Target 1": app_data.get("Target 1"),
            "Target 2": app_data.get("Target 2"),
            "Risk/Reward": app_data.get("Risk/Reward"),
            "Earnings": f"{app_data.get('Earnings Status')} | {app_data.get('Earnings Date')}",
            "SEC Risk": app_data.get("SEC Filing Risk"),
            "News Risk": app_data.get("News Risk"),
        }
        st.dataframe(pd.DataFrame([preview_items]), use_container_width=True, hide_index=True)

    if run_ai_layer:
        api_key = get_openai_api_key()
        if not api_key:
            st.error("Bạn chưa nhập API key hoặc chưa lưu OPENAI_API_KEY trong Streamlit Secrets.")
        elif not st.session_state.ai_app_data:
            st.error("Bạn cần bấm Run App Data Engine trước.")
        elif uploaded_chart is None:
            st.error("Bạn chưa upload chart.")
        else:
            try:
                with st.spinner("Đang chạy Final Analysis..."):
                    result = run_ai_final_analysis(
                        app_data=st.session_state.ai_app_data,
                        image_bytes=uploaded_chart.getvalue(),
                        mime_type=uploaded_chart.type or "image/png",
                        user_context=user_context,
                        trading_style=trading_style,
                        api_key=api_key,
                    )
                    result = normalize_ai_result(result, st.session_state.ai_app_data)
                    st.session_state.ai_result = result
                st.success("Final Analysis đã hoàn tất.")
            except Exception as e:
                st.error(f"Lỗi khi chạy Final Analysis: {e}")

    if st.session_state.ai_result:
        st.markdown('<div class="section-title">Tầng 2: Final Decision</div>', unsafe_allow_html=True)
        render_ai_result(st.session_state.ai_result)
        with st.expander("📤 Nội dung chia sẻ cho bạn bè", expanded=False):
            st.caption("Copy nội dung bên dưới rồi gửi qua text, Zalo, Messenger hoặc email.")
            st.text_area(
                "Share text",
                value=build_share_text(st.session_state.ai_result, st.session_state.ai_app_data),
                height=280,
                label_visibility="collapsed",
            )


with tab_guide:
    st.markdown('<div class="section-title">Cài như app trên iPhone</div>', unsafe_allow_html=True)
    st.markdown(
        """
        1. Mở link app bằng **Safari** trên iPhone.  
        2. Bấm nút **Share**.  
        3. Chọn **Add to Home Screen**.  
        4. Đặt tên **Stock App Pro** rồi bấm **Add**.  

        Bản này có kiến trúc 2 tầng: Tầng 1 Data Engine miễn phí như app hiện tại; Tầng 2 Final Analysis chỉ chạy khi bạn upload chart và bấm nút phân tích.
        """
    )
    st.info("Gợi ý: dùng tab Quick Analyze khi bạn muốn kiểm tra nhanh một mã mới. Dùng PC để phân tích sâu và lưu Excel.")
