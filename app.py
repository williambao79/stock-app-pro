import requests
import json
import time

SEC_USER_AGENT = "StockAnalyzerPro/1.0 contact@example.com"
import os
from datetime import datetime
from copy import copy

import yfinance as yf
import pandas as pd

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

        if "High Risk" in earnings_warning:
            final_decision = "NO TRADE - Earnings Risk"
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
# Cloud/iPhone Web App UI
# Deploy this file as app.py on Streamlit Community Cloud.
# Then add the Streamlit URL to iPhone Home Screen.
# ============================================================

import io
import streamlit as st

st.set_page_config(
    page_title="Stock App Pro",
    page_icon="📈",
    layout="wide"
)

st.title("📈 Stock App Pro")
st.caption("Mở trên iPhone như app: trend, entry, risk/reward, market filter, earnings, news, SEC filings.")

# ---------- Helpers for web version ----------

def parse_tickers(text):
    tickers = []
    for raw in text.replace(",", "\n").splitlines():
        t = raw.strip().upper()
        if t:
            tickers.append(t)
    return list(dict.fromkeys(tickers))


def dataframe_to_excel_bytes(df):
    output = io.BytesIO()
    columns = [
        "Ticker", "Price", "Score", "Signal", "Final Decision",
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


# ---------- Sidebar ----------

with st.sidebar:
    st.header("⚙️ Settings")
    st.write("Nhập ticker, mỗi dòng một mã hoặc cách nhau bằng dấu phẩy.")

    if "ticker_text" not in st.session_state:
        st.session_state.ticker_text = load_default_text()

    ticker_text = st.text_area(
        "Watchlist",
        value=st.session_state.ticker_text,
        height=240,
        key="watchlist_area"
    )

    col_save, col_reset = st.columns(2)
    with col_save:
        if st.button("💾 Save list"):
            tickers_to_save = parse_tickers(ticker_text)
            save_watchlist(tickers_to_save)
            st.session_state.ticker_text = "\n".join(tickers_to_save)
            st.success("Đã lưu watchlist.")
    with col_reset:
        if st.button("↩️ Default"):
            st.session_state.ticker_text = "\n".join(DEFAULT_WATCHLIST)
            st.rerun()

    st.divider()
    st.subheader("Chart Confirmation")
    st.caption("Tùy chọn: nhập xác nhận chart thủ công để app chấm điểm chính xác hơn.")

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
        if st.button("✅ Save chart"):
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
        if st.button("🧹 Clear chart"):
            if not chart_ticker:
                st.warning("Bạn chưa nhập ticker cần xóa.")
            else:
                confirmations = load_chart_confirmations()
                confirmations.pop(chart_ticker, None)
                save_chart_confirmations(confirmations)
                st.success(f"Đã xóa chart signal cho {chart_ticker}.")

# ---------- Main ----------

tickers = parse_tickers(ticker_text)

top1, top2, top3 = st.columns(3)
top1.metric("Số ticker", len(tickers))
top2.metric("File watchlist", WATCHLIST_FILE)
top3.metric("Chart file", CHART_CONFIRMATION_FILE)

with st.expander("📌 Cách cài như app trên iPhone", expanded=True):
    st.markdown(
        """
        **Sau khi deploy lên Streamlit Cloud:**
        1. Mở link app trên Safari của iPhone.
        2. Bấm nút **Share**.
        3. Chọn **Add to Home Screen**.
        4. Từ lần sau chỉ cần bấm icon **Stock App Pro** trên màn hình iPhone.

        **Không cần chung Wi‑Fi, không cần ngrok, không cần mở CMD.**
        """
    )

st.subheader("⚡ Quick Analyze")
st.caption("Dùng khi bạn chỉ muốn xem nhanh 1 mã hoặc vài mã mới, không cần sửa watchlist chính.")
quick_text = st.text_input(
    "Nhập ticker cần phân tích nhanh",
    value="",
    placeholder="Ví dụ: HOOD hoặc NVDA, TSLA, SOUN",
)
quick_tickers = parse_tickers(quick_text)
quick_run = st.button("⚡ RUN QUICK ANALYSIS", use_container_width=True)

if not tickers:
    st.warning("Bạn chưa nhập ticker nào trong watchlist.")
else:
    st.subheader("Watchlist")
    st.write(", ".join(tickers))

run = st.button("🚀 RUN WATCHLIST ANALYSIS", type="primary", use_container_width=True)

analysis_tickers = quick_tickers if quick_run else tickers

if run or quick_run:
    if not analysis_tickers:
        if quick_run:
            st.error("Bạn chưa nhập ticker để phân tích nhanh.")
        else:
            st.error("Bạn chưa nhập ticker trong watchlist.")
    else:
        progress = st.progress(0)
        status = st.empty()

        try:
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

            st.success(f"Hoàn thành! Market Condition: {market_condition}")
            st.caption(f"Market notes: {market_notes}")

            # Mobile-friendly summary cards
            st.subheader("📱 Tóm tắt nhanh")
            for _, row in df.iterrows():
                ticker = row.get("Ticker", "")
                decision = row.get("Final Decision", "")
                score = row.get("Score", "")
                price = row.get("Price", "")
                action = row.get("Action", "")
                entry = row.get("Aggressive Entry", "")
                stop = row.get("Stop Loss", "")
                target1 = row.get("Target 1", "")
                rr = row.get("Risk/Reward", "")

                with st.container(border=True):
                    a, b, c = st.columns([1, 1, 1])
                    a.metric(str(ticker), f"${price}")
                    b.metric("Score", score)
                    c.metric("Decision", str(decision))

                    st.write(f"**Action:** {action}")
                    st.write(f"**Entry:** {entry} | **Stop:** {stop} | **Target 1:** {target1} | **R/R:** {rr}")

                    reasons = row.get("Reasons", "")
                    if reasons:
                        with st.expander("Lý do phân tích"):
                            st.write(reasons)

                    catalyst = row.get("Recent Catalyst", "")
                    sec = row.get("Recent SEC Filing", "")
                    if catalyst or sec:
                        with st.expander("News / SEC"):
                            if catalyst:
                                st.write(f"**News:** {catalyst}")
                            if sec:
                                st.write(f"**SEC:** {sec}")

            st.subheader("📊 Bảng chi tiết")
            st.dataframe(df, use_container_width=True, hide_index=True)

            excel_bytes = dataframe_to_excel_bytes(df)
            st.download_button(
                label="⬇️ Download Excel",
                data=excel_bytes,
                file_name="market_analysis_iphone.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Lỗi: {e}")
