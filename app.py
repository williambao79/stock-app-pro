import requests
import json
import time

SEC_USER_AGENT = "StockAnalyzerPro/1.0 contact@example.com"
import os
from datetime import datetime
from copy import copy

import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta
try:
    from openpyxl import load_workbook
except Exception:
    load_workbook = None

WATCHLIST_FILE = "watchlist.txt"
TEMPLATE_FILE = "market_template.xlsx"
CHART_CONFIRMATION_FILE = "chart_confirmation.json"
OUTPUT_FILE = "market_template.xlsx"
@st.cache_data(ttl=3600)
def get_earnings_status(ticker):
    """
    Lấy ngày earnings gần nhất bằng yfinance.
    Trả về:
    - status: trạng thái earnings
    - date_text: ngày earnings nếu có
    - days_until: số ngày còn lại
    """

    try:
        stock = yf.Ticker(ticker)

        earnings_date = None

        # Cách 1: lấy từ get_earnings_dates
        try:
            ed = stock.get_earnings_dates(limit=12)
            if ed is not None and not ed.empty:
                today = pd.Timestamp.today(tz=ed.index.tz) if ed.index.tz is not None else pd.Timestamp.today()
                future_dates = ed[ed.index >= today]

                if not future_dates.empty:
                    earnings_date = future_dates.index[0]
                else:
                    earnings_date = ed.index[0]
        except Exception:
            pass

        # Cách 2: fallback dùng calendar
        if earnings_date is None:
            try:
                cal = stock.calendar
                if cal is not None and len(cal) > 0:
                    if isinstance(cal, dict):
                        possible = cal.get("Earnings Date")
                        if possible is not None:
                            if isinstance(possible, list):
                                earnings_date = possible[0]
                            else:
                                earnings_date = possible
                    else:
                        if "Earnings Date" in cal.index:
                            possible = cal.loc["Earnings Date"][0]
                            earnings_date = possible
            except Exception:
                pass

        if earnings_date is None:
            return {
                "status": "Unknown",
                "date": "Unknown",
                "days_until": None,
                "note": "No confirmed earnings date"
            }

        earnings_date = pd.to_datetime(earnings_date).tz_localize(None)
        today = pd.Timestamp.today().normalize()
        days_until = (earnings_date.normalize() - today).days

        date_text = earnings_date.strftime("%Y-%m-%d")

        if 0 <= days_until <= 7:
            status = "High Risk"
            risk = f"Earnings in {days_until} day(s)"
        elif 8 <= days_until <= 14:
            status = "Caution"
            risk = f"Earnings soon: {days_until} day(s)"
        elif days_until < 0:
            status = "Reported"
            risk = f"Last earnings was {abs(days_until)} day(s) ago"
        else:
            status = "Clear"
            risk = f"Earnings is {days_until} day(s) away"

        return {
            "status": status,
            "date": date_text,
            "days_until": days_until,
            "note": risk
        }

    except Exception as e:
        return {
            "status": "Unknown",
            "date_text": "Unknown",
            "days_until": None,
            "note": "Earnings data unavailable"
        }
DEFAULT_WATCHLIST = [
    "ENVX", "SOUN", "WULF", "FRMI", "PATH", "RCAT", "QXO",
    "NVDA", "VGT", "VOO", "SLV", "GDX"
]

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


def find_support_resistance(data, lookback=30):
    recent = data.tail(lookback)
    support = float(recent["Low"].min())
    resistance = float(recent["High"].max())
    return support, resistance


def find_recent_swing_support(data, lookback=15):
    recent = data.tail(lookback)
    return float(recent["Low"].min())


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


def get_earnings_warning(ticker):
    try:
        stock = yf.Ticker(ticker)
        calendar = stock.calendar

        if calendar is None or len(calendar) == 0:
            return "Unknown"

        earnings_date = None

        if isinstance(calendar, pd.DataFrame):
            for idx in calendar.index:
                if "Earnings" in str(idx):
                    value = calendar.loc[idx].dropna()
                    if len(value) > 0:
                        earnings_date = pd.to_datetime(value.iloc[0]).date()
                        break
        elif isinstance(calendar, dict):
            for key, value in calendar.items():
                if "Earnings" in str(key):
                    if isinstance(value, (list, tuple)) and len(value) > 0:
                        earnings_date = pd.to_datetime(value[0]).date()
                    else:
                        earnings_date = pd.to_datetime(value).date()
                    break

        if earnings_date is None:
            return "Unknown"

        today = datetime.now().date()
        days = (earnings_date - today).days

        if 0 <= days <= 7:
            return f"High Risk: Earnings in {days} days"
        elif 8 <= days <= 14:
            return f"Watch: Earnings in {days} days"
        elif days > 14:
            return f"OK: Earnings in {days} days"
        else:
            return "Past/Unknown"
    except Exception:
        return "Unknown"

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


def get_yahoo_news_signal(ticker):
    try:
        stock = yf.Ticker(ticker)
        news = stock.news

        if not news:
            return {
                "News Sentiment": "Neutral",
                "News Risk": "Unknown",
                "Recent Catalyst": "No Yahoo news found"
            }

        titles = []
        now_ts = datetime.now().timestamp()
        recent_count = 0

        for item in news[:10]:
            title = ""

            # Old yfinance format
            if isinstance(item, dict):
                title = safe_text(item.get("title", ""))

                # New yfinance format sometimes stores title inside content
                if title == "" and isinstance(item.get("content"), dict):
                    content = item.get("content", {})
                    title = safe_text(content.get("title", ""))

                # Another possible nested format
                if title == "" and isinstance(item.get("content"), dict):
                    content = item.get("content", {})
                    title = safe_text(content.get("headline", ""))

                publish_time = item.get("providerPublishTime")

                if publish_time is None and isinstance(item.get("content"), dict):
                    content = item.get("content", {})
                    publish_time = content.get("pubDate")

                if isinstance(publish_time, str):
                    try:
                        publish_dt = pd.to_datetime(publish_time)
                        days_old = (datetime.now() - publish_dt.to_pydatetime().replace(tzinfo=None)).days
                    except Exception:
                        days_old = 999
                elif publish_time:
                    days_old = (now_ts - float(publish_time)) / 86400
                else:
                    days_old = 999

                if title:
                    if days_old <= 21:
                        recent_count += 1
                    titles.append(title)

        if not titles:
            return {
                "News Sentiment": "Neutral",
                "News Risk": "Unknown",
                "Recent Catalyst": "No readable Yahoo news title"
            }

        combined = " | ".join(titles[:2])
        short_catalyst = combined[:120]
        sentiment = simple_news_sentiment(combined)

        if sentiment == "Negative":
            news_risk = "High"
        elif sentiment == "Positive":
            news_risk = "Low"
        else:
            news_risk = "Medium" if recent_count > 0 else "Unknown"

        return {
            "News Sentiment": sentiment,
            "News Risk": news_risk,
            "Recent Catalyst": short_catalyst
        }

    except Exception as e:
        return {
            "News Sentiment": "Neutral",
            "News Risk": "Unknown",
            "Recent Catalyst": f"News check error: {str(e)[:120]}"
        }


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
    try:
        ticker_map = get_sec_cik_map()
        info = ticker_map.get(ticker.upper())

        if not info:
            return {
                "SEC Filing Risk": "Unknown",
                "Recent SEC Filing": "Ticker not found in SEC ticker list"
            }

        cik = str(info["cik"]).zfill(10)

        headers = {
            "User-Agent": SEC_USER_AGENT,
            "Accept-Encoding": "gzip, deflate",
            "Host": "data.sec.gov"
        }

        url = f"https://data.sec.gov/submissions/CIK{cik}.json"
        r = requests.get(url, headers=headers, timeout=15)
        r.raise_for_status()

        data = r.json()
        recent = data.get("filings", {}).get("recent", {})

        forms = recent.get("form", [])
        filing_dates = recent.get("filingDate", [])
        descriptions = recent.get("primaryDocDescription", [])
        accession_nums = recent.get("accessionNumber", [])

        today = datetime.now().date()
        recent_items = []

        for i in range(min(len(forms), 25)):
            form = safe_text(forms[i])
            filing_date_raw = safe_text(filing_dates[i])
            desc = safe_text(descriptions[i]) if i < len(descriptions) else ""
            accession = safe_text(accession_nums[i]) if i < len(accession_nums) else ""

            try:
                filing_date = datetime.strptime(filing_date_raw, "%Y-%m-%d").date()
                days_old = (today - filing_date).days
            except Exception:
                days_old = 999

            if days_old <= 45:
                recent_items.append({
                    "form": form,
                    "date": filing_date_raw,
                    "desc": desc,
                    "accession": accession,
                    "days_old": days_old
                })

        if not recent_items:
            return {
                "SEC Filing Risk": "Low",
                "Recent SEC Filing": "No major recent SEC filing in 45 days"
            }

        risky_forms = {"S-1", "S-3", "F-1", "F-3", "424B3", "424B4", "424B5", "424B7", "424B8", "FWP"}
        medium_forms = {"8-K", "10-Q", "10-K", "6-K"}
        insider_forms = {"3", "4", "5"}

        risk = "Low"
        notes = []

        for item in recent_items[:8]:
            form = item["form"]
            desc = item["desc"]
            text = f"{form} {desc}".lower()
            label = f"{item['date']} {form} {desc}".strip()
            notes.append(label)

            if form in risky_forms:
                risk = "High"
            elif any(w in text for w in ["offering", "shelf", "prospectus", "atm", "warrant", "resale"]):
                risk = "High"
            elif form in medium_forms and risk != "High":
                risk = "Medium"
            elif form in insider_forms and risk == "Low":
                risk = "Low"

        return {
            "SEC Filing Risk": risk,
            "Recent SEC Filing": " | ".join(notes[:5])[:700]
        }

    except Exception as e:
        return {
            "SEC Filing Risk": "Unknown",
            "Recent SEC Filing": f"SEC check error: {str(e)[:150]}"
        }


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

        support, resistance = find_support_resistance(data, lookback=30)
        swing_support = find_recent_swing_support(data, lookback=15)

        score = 0
        reasons = []
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

        safe_entry_low = support * 1.01
        safe_entry_high = support * 1.04

        if trend in ["Strong Uptrend", "Uptrend", "Neutral Up"]:
            aggressive_entry_low = min(ema9, ema21) * 0.99
            aggressive_entry_high = max(ema9, ema21) * 1.01
        elif setup == "Near Support":
            aggressive_entry_low = swing_support * 1.00
            aggressive_entry_high = swing_support * 1.04
        else:
            aggressive_entry_low = safe_entry_low
            aggressive_entry_high = safe_entry_high

        if aggressive_entry_low > close:
            aggressive_entry_low = close * 0.98
            aggressive_entry_high = close * 1.01

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

        aggressive_stop = min(aggressive_entry_low * 0.96, ema21 * 0.97)
        safe_stop = support * 0.96

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
            target_1 = close + risk * 1.5
            target_2 = resistance
            reward = max(target_1, target_2) - close
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

        # REVERSAL WATCH:
        # Dành cho cổ phiếu còn Downtrend/Mixed nhưng đang hồi từ vùng thấp.
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
        if score >= 8 and entry_quality in ["Near Entry", "Good Pullback Entry"] and market_condition != "Bearish":
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
        elif earnings_status == "Caution" and signal == "BUY":
            final_decision = "CAUTION - Earnings Soon"
        elif sec_filing_risk == "High":
            final_decision = "NO TRADE - SEC Filing Risk"
        elif chart_confirmation == "Weak / Breakdown":
            final_decision = "NO TRADE - Chart Weak"
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
            final_decision = "BREAKOUT WATCH" if market_condition != "Bearish" and trend in ["Strong Uptrend", "Uptrend", "Neutral Up"] and 45 <= rsi <= 68 and momentum in ["Strong", "Positive", "Improving"] and setup in ["Middle Zone", "Near Support"] and chart_confirmation != "Weak / Breakdown" else "REVERSAL WATCH" if market_condition != "Bearish" and trend == "Downtrend" and 40 <= rsi <= 58 and momentum in ["Strong", "Positive", "Improving"] and volume_status in ["Above Average", "Strong Volume", "Normal/Weak"] and setup in ["Middle Zone", "Near Support"] and chart_confirmation != "Weak / Breakdown" else "NO TRADE"     
            signal = "WAIT" if final_decision in ["BREAKOUT WATCH", "REVERSAL WATCH"] else signal
            action = "Breakout Watch - chờ vượt resistance hoặc volume xác nhận" if final_decision == "BREAKOUT WATCH" else "Reversal Watch - chờ vượt kháng cự hoặc giữ support" if final_decision == "REVERSAL WATCH" else action
            score = max(score, 3) if final_decision in ["BREAKOUT WATCH", "REVERSAL WATCH"] else score

        return {
            "Ticker": ticker,
            "Price": round(close, 2),
            "Score": score,
            "Signal": signal,
            "Final Decision": final_decision,
            "Earnings Status": earnings_status,
            "Earnings Date": earnings_date,
            "Earnings Note": earnings_note,
            "Market Condition": market_condition,
            "Market Filter": market_filter,
            "Trend": trend,
            "Setup": setup,
            "Action": action,
            "RSI": round(rsi, 2),
            "RSI Status": rsi_status,
            "Momentum": momentum,
            "Volume Status": volume_status,
            "Support": round(support, 2),
            "Resistance": round(resistance, 2),
            "Aggressive Entry": f"{round(aggressive_entry_low, 2)} - {round(aggressive_entry_high, 2)}",
            "Safe Entry": f"{round(safe_entry_low, 2)} - {round(safe_entry_high, 2)}",
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
        "Earnings Status", "Earnings Date", "Earnings Note",
        "Market Condition", "Market Filter",
        "Trend", "Setup", "Action",
        "RSI", "RSI Status", "Momentum", "Volume Status",
        "Support", "Resistance",
        "Aggressive Entry", "Safe Entry", "Entry Distance %",
        "Entry Quality",
        "Aggressive Stop", "Safe Stop", "Stop Loss",
        "Target 1", "Target 2", "Risk/Reward",
        "Chart Confirmation", "Chart Adjustment", "Chart Note",
        "Earnings Warning",
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

        for row in ws.iter_rows(min_row=2, max_row=1000, min_col=1, max_col=len(columns)):
            for cell in row:
                cell.value = None

        

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
# Cloud/iPhone Web App UI - V3 Clean Mobile Design
# Deploy this file as app.py on Streamlit Community Cloud.
# ============================================================

import io
import streamlit as st

st.set_page_config(
    page_title="Stock App Pro",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown(
    """
    <style>
    :root {
        --card-bg: #ffffff;
        --soft-bg: #f6f8fb;
        --text-main: #0f172a;
        --text-muted: #64748b;
        --border: #e5e7eb;
        --green: #16a34a;
        --yellow: #ca8a04;
        --red: #dc2626;
        --blue: #2563eb;
        --purple: #7c3aed;
    }
    .main .block-container {
        padding-top: 1.1rem;
        padding-bottom: 3rem;
        max-width: 1120px;
    }
    section[data-testid="stSidebar"] .block-container {
        padding-top: 1rem;
    }
    .hero {
        background: linear-gradient(135deg, #0f172a 0%, #1d4ed8 55%, #7c3aed 100%);
        color: white;
        border-radius: 24px;
        padding: 22px 20px;
        box-shadow: 0 12px 35px rgba(15, 23, 42, 0.22);
        margin-bottom: 16px;
    }
    .hero h1 {
        font-size: 30px;
        margin: 0 0 6px 0;
        line-height: 1.1;
    }
    .hero p {
        margin: 0;
        color: rgba(255,255,255,0.86);
        font-size: 15px;
    }
    .mini-note {
        background: #eff6ff;
        color: #1e3a8a;
        border: 1px solid #bfdbfe;
        border-radius: 16px;
        padding: 11px 14px;
        margin: 8px 0 14px 0;
        font-size: 14px;
    }
    .stock-card {
        background: var(--card-bg);
        border: 1px solid var(--border);
        border-radius: 22px;
        padding: 16px;
        margin-bottom: 14px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
    }
    .stock-head {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 10px;
        margin-bottom: 12px;
    }
    .ticker-title {
        font-size: 25px;
        font-weight: 800;
        color: var(--text-main);
        margin: 0;
        letter-spacing: -0.02em;
    }
    .price-line {
        color: var(--text-muted);
        font-size: 14px;
        margin-top: 2px;
    }
    .badge {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        border-radius: 999px;
        padding: 7px 11px;
        font-size: 12px;
        font-weight: 800;
        white-space: nowrap;
        text-align: center;
    }
    .badge-good { background: #dcfce7; color: #166534; }
    .badge-watch { background: #fef3c7; color: #92400e; }
    .badge-bad { background: #fee2e2; color: #991b1b; }
    .badge-info { background: #dbeafe; color: #1e40af; }
    .badge-neutral { background: #f1f5f9; color: #334155; }
    .metric-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 10px;
        margin: 10px 0 12px 0;
    }
    .metric-box {
        background: var(--soft-bg);
        border-radius: 16px;
        padding: 12px;
        border: 1px solid #eef2f7;
    }
    .metric-label {
        color: var(--text-muted);
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: .03em;
        margin-bottom: 4px;
    }
    .metric-value {
        color: var(--text-main);
        font-size: 18px;
        font-weight: 800;
        line-height: 1.15;
        word-break: break-word;
    }
    .action-box {
        border-radius: 16px;
        padding: 13px 14px;
        background: #f8fafc;
        border-left: 5px solid #2563eb;
        color: #0f172a;
        margin-top: 8px;
        font-size: 14px;
    }
    .small-muted { color: var(--text-muted); font-size: 13px; }
    .section-title {
        font-size: 21px;
        font-weight: 800;
        margin: 16px 0 6px 0;
        color: var(--text-main);
    }
    .pill-row { display:flex; flex-wrap:wrap; gap:8px; margin: 8px 0 12px 0; }
    .pill {
        display:inline-block;
        padding:7px 11px;
        border-radius:999px;
        background:#f1f5f9;
        color:#334155;
        font-size:13px;
        font-weight:700;
    }
    div.stButton > button {
        border-radius: 16px;
        min-height: 48px;
        font-weight: 800;
    }
    div[data-testid="stTextInput"] input, textarea {
        border-radius: 14px !important;
    }
    @media (max-width: 640px) {
        .main .block-container { padding-left: 0.9rem; padding-right: 0.9rem; }
        .hero { border-radius: 20px; padding: 18px 16px; }
        .hero h1 { font-size: 27px; }
        .metric-grid { grid-template-columns: 1fr; }
        .ticker-title { font-size: 23px; }
        .stock-head { flex-direction: column; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------- Helpers for web version ----------

def parse_tickers(text):
    tickers = []
    for raw in str(text).replace(",", "\n").replace(";", "\n").splitlines():
        t = raw.strip().upper().replace("$", "")
        if t:
            tickers.append(t)
    return list(dict.fromkeys(tickers))


def dataframe_to_excel_bytes(df):
    output = io.BytesIO()
    columns = [
        "Ticker", "Price", "Score", "Signal", "Final Decision",
        "Earnings Status", "Earnings Date", "Earnings Note",
        "Market Condition", "Market Filter",
        "Trend", "Setup", "Action",
        "RSI", "RSI Status", "Momentum", "Volume Status",
        "Support", "Resistance",
        "Aggressive Entry", "Safe Entry", "Entry Distance %",
        "Entry Quality",
        "Aggressive Stop", "Safe Stop", "Stop Loss",
        "Target 1", "Target 2", "Risk/Reward",
        "Chart Confirmation", "Chart Adjustment", "Chart Note",
        "Earnings Warning",
        "News Sentiment", "News Risk", "Recent Catalyst",
        "SEC Filing Risk", "Recent SEC Filing",
        "Reasons", "Status", "Market Notes"
    ]

    for col in columns:
        if col not in df.columns:
            df[col] = ""

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df[columns].to_excel(writer, index=False, sheet_name="Analysis")

    output.seek(0)
    return output.getvalue()


def load_default_text():
    return "\n".join(load_watchlist())


def badge_class(decision):
    d = str(decision).upper()
    if "WATCH TO ENTER" in d or "BUY" in d:
        return "badge-good"
    if "BREAKOUT" in d or "REVERSAL" in d or "PULLBACK" in d or "WAIT" in d:
        return "badge-watch"
    if "NO TRADE" in d or "AVOID" in d or "WEAK" in d or "RISK" in d:
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
    action = row.get("Action", "")
    trend = row.get("Trend", "")
    setup = row.get("Setup", "")
    rsi = safe_value(row.get("RSI", ""))
    rr = safe_value(row.get("Risk/Reward", ""))
    entry = row.get("Aggressive Entry", "")
    safe_entry = row.get("Safe Entry", "")
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
                    <div class="price-line">Price {price} · Score {score}</div>
                </div>
                <span class="badge {badge_class(decision)}">{decision}</span>
            </div>
            <div class="metric-grid">
                <div class="metric-box"><div class="metric-label">Entry đẹp</div><div class="metric-value">{entry or '—'}</div></div>
                <div class="metric-box"><div class="metric-label">Stop loss</div><div class="metric-value">{stop}</div></div>
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
            <div class="action-box"><b>Action:</b> {action or '—'}<br><span class="small-muted">Safe Entry: {safe_entry or '—'}</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    reasons = row.get("Reasons", "")
    catalyst = row.get("Recent Catalyst", "")
    sec = row.get("Recent SEC Filing", "")
    earnings = row.get("Earnings Warning", "")
    earnings_status = row.get("Earnings Status", "")
    earnings_date = row.get("Earnings Date", "")
    earnings_note = row.get("Earnings Note", "")
    chart_note = row.get("Chart Note", "")

    with st.expander(f"Chi tiết {ticker}: lý do, news, SEC"):
        if reasons:
            st.write("**Lý do phân tích:**")
            st.write(reasons)
        if earnings:
            st.write(f"**Earnings:** {earnings}")
        if earnings_status or earnings_date or earnings_note:
            st.write(f"**Earnings Report:** {earnings_status} | {earnings_date} | {earnings_note}")
        if chart_note:
            st.write(f"**Chart note:** {chart_note}")
        if catalyst:
            st.write(f"**News:** {catalyst}")
        if sec:
            st.write(f"**SEC:** {sec}")


def run_analysis_for_tickers(analysis_tickers):
    progress = st.progress(0)
    status = st.empty()
    status.info("Đang kiểm tra thị trường chung...")
    market_condition, market_notes = get_market_condition()
    chart_confirmations = load_chart_confirmations()
    results = []

    for i, ticker in enumerate(analysis_tickers, start=1):
        status.info(f"Đang phân tích {ticker}... ({i}/{len(analysis_tickers)})")
        result = analyze_stock(
            ticker,
            market_condition=market_condition,
            chart_confirmations=chart_confirmations
        )
        result["Market Notes"] = market_notes
        results.append(result)
        progress.progress(i / len(analysis_tickers))

    df = pd.DataFrame(results)
    if "Score" in df.columns:
        df = df.sort_values(by="Score", ascending=False)
    status.empty()
    progress.empty()
    return market_condition, market_notes, df


# ---------- Header ----------

st.markdown(
    """
    <div class="hero">
        <h1>📈 Stock App Pro</h1>
        <p>Phân tích swing trade trên iPhone: điểm vào, stop loss, target, risk/reward, market filter, news và SEC.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="mini-note">
        Cách dùng nhanh: nhập mã ở <b>Quick Analyze</b> → bấm <b>RUN QUICK ANALYSIS</b>. Không cần sửa watchlist chính.
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------- Session defaults ----------

if "ticker_text" not in st.session_state:
    st.session_state.ticker_text = load_default_text()
if "quick_text" not in st.session_state:
    st.session_state.quick_text = ""

# ---------- Main tabs ----------

tab_quick, tab_watchlist, tab_chart, tab_guide = st.tabs([
    "⚡ Quick Analyze",
    "📋 Watchlist",
    "📌 Chart Signal",
    "📱 iPhone Guide",
])

with tab_quick:
    st.markdown('<div class="section-title">Phân tích nhanh mã mới</div>', unsafe_allow_html=True)
    st.caption("Gõ 1 mã hoặc nhiều mã. Ví dụ: HOOD hoặc NVDA, TSLA, SOUN")
    quick_text = st.text_input(
        "Nhập ticker",
        value=st.session_state.quick_text,
        placeholder="HOOD, NVDA, TSLA",
        key="quick_input",
        label_visibility="collapsed"
    )
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
                st.markdown(
                    f"<span class='badge {market_badge_class(market_condition)}'>Market: {market_condition}</span>",
                    unsafe_allow_html=True,
                )
                st.caption(f"Market notes: {market_notes}")

                st.markdown('<div class="section-title">Kết quả</div>', unsafe_allow_html=True)
                for _, row in df.iterrows():
                    render_stock_card(row)

                with st.expander("📊 Bảng chi tiết"):
                    st.dataframe(df, use_container_width=True, hide_index=True)

                excel_bytes = dataframe_to_excel_bytes(df)
                st.download_button(
                    label="⬇️ Download Excel",
                    data=excel_bytes,
                    file_name="quick_analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Lỗi: {e}")

with tab_watchlist:
    st.markdown('<div class="section-title">Watchlist chính</div>', unsafe_allow_html=True)
    st.caption("Danh sách này dùng cho những mã bạn theo dõi thường xuyên.")
    ticker_text = st.text_area(
        "Watchlist",
        value=st.session_state.ticker_text,
        height=220,
        placeholder="ENVX\nSOUN\nWULF\nPATH",
    )
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
                st.markdown(
                    f"<span class='badge {market_badge_class(market_condition)}'>Market: {market_condition}</span>",
                    unsafe_allow_html=True,
                )
                st.caption(f"Market notes: {market_notes}")

                st.markdown('<div class="section-title">Kết quả watchlist</div>', unsafe_allow_html=True)
                for _, row in df.iterrows():
                    render_stock_card(row)

                with st.expander("📊 Bảng chi tiết"):
                    st.dataframe(df, use_container_width=True, hide_index=True)

                excel_bytes = dataframe_to_excel_bytes(df)
                st.download_button(
                    label="⬇️ Download Excel",
                    data=excel_bytes,
                    file_name="watchlist_analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Lỗi: {e}")

with tab_chart:
    st.markdown('<div class="section-title">Chart Confirmation thủ công</div>', unsafe_allow_html=True)
    st.caption("Dùng khi bạn tự xem chart và muốn app cộng/trừ điểm theo nhận định của bạn.")
    chart_ticker = st.text_input("Ticker chart", value="", placeholder="Ví dụ: PATH").upper().strip()
    chart_confirmation = st.selectbox(
        "Chart signal",
        [
            "Neutral",
            "Bullish Breakout",
            "Pullback Holding",
            "Reversal Setup",
            "Near Resistance",
            "Weak / Breakdown"
        ],
        index=0
    )
    chart_note = st.text_input("Note", value="", placeholder="Ví dụ: giữ EMA21, gần resistance...")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("✅ Save chart signal", use_container_width=True):
            if not chart_ticker:
                st.warning("Bạn chưa nhập ticker.")
            else:
                confirmations = load_chart_confirmations()
                confirmations[chart_ticker] = {
                    "confirmation": chart_confirmation,
                    "note": chart_note
                }
                save_chart_confirmations(confirmations)
                st.success(f"Đã lưu chart signal cho {chart_ticker}.")
    with c2:
        if st.button("🧹 Clear chart signal", use_container_width=True):
            if not chart_ticker:
                st.warning("Bạn chưa nhập ticker cần xóa.")
            else:
                confirmations = load_chart_confirmations()
                confirmations.pop(chart_ticker, None)
                save_chart_confirmations(confirmations)
                st.success(f"Đã xóa chart signal cho {chart_ticker}.")

    confirmations = load_chart_confirmations()
    if confirmations:
        st.markdown('<div class="section-title">Chart signals đã lưu</div>', unsafe_allow_html=True)
        st.dataframe(
            pd.DataFrame([
                {"Ticker": k, "Signal": v.get("confirmation", ""), "Note": v.get("note", "")}
                for k, v in confirmations.items()
            ]),
            use_container_width=True,
            hide_index=True,
        )

with tab_guide:
    st.markdown('<div class="section-title">Cài như app trên iPhone</div>', unsafe_allow_html=True)
    st.markdown(
        """
        1. Mở link app bằng **Safari** trên iPhone.  
        2. Bấm nút **Share**.  
        3. Chọn **Add to Home Screen**.  
        4. Đặt tên **Stock App Pro** rồi bấm **Add**.  

        Từ lần sau bạn chỉ cần bấm icon ngoài màn hình iPhone. Không cần CMD, không cần ngrok, không cần máy tính ở nhà bật.
        """
    )
    st.info("Gợi ý: dùng tab Quick Analyze khi bạn muốn kiểm tra nhanh một mã mới như HOOD, NVDA, TSLA.")
