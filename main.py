import os
import re
from datetime import datetime, timezone
import requests
from bs4 import BeautifulSoup
import openpyxl
import matplotlib.pyplot as plt

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "").strip()
EXCEL_FILE = "World_Gold_Ratio_TimeSeries_Ready.xlsx"
CHART_FILE = "gold_fx_trend.png"
OZ_PER_TONNE = 1_000_000.0 / 31.1034768

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}


def fetch_fx_and_gold_rates():
    """دریافت نرخ برابری ارزها به دلار و قیمت لحظه‌ای اونس طلا (XAU/USD)."""
    rates = {}
    xau_usd = None

    try:
        resp = requests.get("https://open.er-api.com/v6/latest/USD", timeout=20)
        resp.raise_for_status()
        data = resp.json()
        if data.get("result") == "success" and "rates" in data:
            for code, usd_per_unit in data["rates"].items():
                if usd_per_unit and usd_per_unit > 0:
                    # تبدیل از نرخ "واحد ارز به ازای 1 دلار" به "ارزش دلاری هر 1 واحد ارز"
                    rates[code.upper()] = 1.0 / float(usd_per_unit)
            rates["USD"] = 1.0
    except Exception as exc:
        print(f"[Warn] Error fetching FX rates API: {exc}")

    # دریافت قیمت اونس طلا از Trading Economics
    try:
        resp = requests.get(
            "https://tradingeconomics.com/commodity/gold",
            headers=HEADERS,
            timeout=20,
        )
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, "html.parser")
            # جستجوی عدد قیمت در تگ‌های رایج یا جدول
            price_el = soup.find(id="p")
            if price_el and price_el.text.strip():
                cleaned = re.sub(r"[^\d.]", "", price_el.text.strip())
                if cleaned:
                    xau_usd = float(cleaned)
            if xau_usd is None:
                for td in soup.find_all(["td", "span", "div"]):
                    txt = td.get_text(" ", strip=True)
                    if "Gold" in txt and "$" not in txt:
                        continue
                    m = re.search(r"\b(\d{4}\.\d{2})\b", txt)
                    if m:
                        val = float(m.group(1))
                        if 1500.0 <= val <= 10000.0:
                            xau_usd = val
                            break
    except Exception as exc:
        print(f"[Warn] Error scraping XAU price: {exc}")

    return rates, xau_usd


def parse_trading_economics_table(url):
    """تلاش برای استخراج جدول کشورها از Trading Economics."""
    rows_out = {}
    try:
        resp = requests.get(url, headers=HEADERS, timeout=25)
        if resp.status_code != 200:
            return rows_out
        soup = BeautifulSoup(resp.text, "html.parser")
        for tr in soup.find_all("tr"):
            cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
            cells = [re.sub(r"\s+", " ", c).strip() for c in cells if c.strip()]
            if len(cells) >= 2:
                country = cells[0].replace("▲", "").replace("▼", "").strip()
                num_match = re.search(r"[-+]?\d[\d,]*\.?\d*", cells[1])
                if country and num_match and country.lower() != "country":
                    val = float(num_match.group(0).replace(",", ""))
                    unit = cells[4] if len(cells) >= 5 else ""
                    ref = cells[3] if len(cells) >= 4 else ""
                    rows_out[country] = {"last": val, "unit": unit, "ref": ref}
    except Exception as exc:
        print(f"[Warn] Error scraping {url}: {exc}")
    return rows_out


def update_workbook(filepath):
    wb = openpyxl.load_workbook(filepath)

    # ۱. به‌روزرسانی نرخ ارز و XAU در FX_Rates_&_Params
    rates, xau_live = fetch_fx_and_gold_rates()
    ws_rates = wb["FX_Rates_&_Params"]

    xau_row = None
    for r in range(6, ws_rates.max_row + 1):
        code = str(ws_rates.cell(r, 2).value or "").strip().upper()
        if not code:
            continue
        if code == "XAU":
            xau_row = r
            if xau_live is not None:
                ws_rates.cell(r, 3).value = xau_live
        elif code in rates:
            ws_rates.cell(r, 3).value = rates[code]

    # خواندن قیمت نهایی XAU (زنده یا مقدار موجود در شیت)
    current_xau = float(ws_rates.cell(xau_row, 3).value) if xau_row else 3852.26

    # ۲. تلاش برای به‌روزرسانی جداول ذخایر ارزی و طلا در صورت دسترسی
    fx_scraped = parse_trading_economics_table(
        "https://tradingeconomics.com/country-list/foreign-exchange-reserves"
    )
    gold_scraped = parse_trading_economics_table(
        "https://tradingeconomics.com/country-list/gold-reserves"
    )

    if fx_scraped and "PQ_FX_Reserves" in wb.sheetnames:
        ws_fx = wb["PQ_FX_Reserves"]
        for r in range(5, ws_fx.max_row + 1):
            c_name = str(ws_fx.cell(r, 1).value or "").strip()
            if c_name in fx_scraped:
                ws_fx.cell(r, 2).value = fx_scraped[c_name]["last"]
                if fx_scraped[c_name]["ref"]:
                    ws_fx.cell(r, 4).value = fx_scraped[c_name]["ref"]

    if gold_scraped and "PQ_Gold_Reserves" in wb.sheetnames:
        ws_gold = wb["PQ_Gold_Reserves"]
        for r in range(5, ws_gold.max_row + 1):
            c_name = str(ws_gold.cell(r, 1).value or "").strip()
            if c_name in gold_scraped:
                ws_gold.cell(r, 2).value = gold_scraped[c_name]["last"]
                if gold_scraped[c_name]["ref"]:
                    ws_gold.cell(r, 4).value = gold_scraped[c_name]["ref"]

    # ۳. محاسبه مقادیر کلان از Live_Analysis (یا مقادیر موجود)
    ws_live = wb["Live_Analysis"]
    fx_map = {}
    if "PQ_FX_Reserves" in wb.sheetnames:
        ws_fx = wb["PQ_FX_Reserves"]
        for r in range(5, ws_fx.max_row + 1):
            c_name = str(ws_fx.cell(r, 1).value or "").strip()
            last_v = ws_fx.cell(r, 2).value
            unit_s = str(ws_fx.cell(r, 5).value or "").strip()
            if c_name and isinstance(last_v, (int, float)):
                parts = unit_s.split()
                cur_code = parts[0].upper() if parts else "USD"
                scale_word = parts[1].capitalize() if len(parts) > 1 else "Million"
                scale_mult = {
                    "Million": 1.0,
                    "Billion": 1000.0,
                    "Thousand": 0.001,
                    "Hundred": 100.0,
                }.get(scale_word, 1.0)
                fx_usd_rate = rates.get(cur_code, 1.0 if cur_code == "USD" else None)
                if fx_usd_rate is not None:
                    fx_map[c_name] = float(last_v) * fx_usd_rate * scale_mult

    gold_map = {}
    if "PQ_Gold_Reserves" in wb.sheetnames:
        ws_gold = wb["PQ_Gold_Reserves"]
        for r in range(5, ws_gold.max_row + 1):
            c_name = str(ws_gold.cell(r, 1).value or "").strip()
            last_v = ws_gold.cell(r, 2).value
            if c_name and isinstance(last_v, (int, float)):
                gold_map[c_name] = float(last_v)

    total_fx_m = 0.0
    total_gold_tonnes = 0.0
    for r in range(5, ws_live.max_row + 1):
        inc = str(ws_live.cell(r, 1).value or "").strip()
        c_name = str(ws_live.cell(r, 2).value or "").strip()
        if inc.lower() == "yes" and c_name:
            fx_val = fx_map.get(c_name, ws_live.cell(r, 5).value)
            g_val = gold_map.get(c_name, ws_live.cell(r, 6).value)
            if isinstance(fx_val, (int, float)):
                ws_live.cell(r, 5).value = float(fx_val)
                total_fx_m += float(fx_val)
            if isinstance(g_val, (int, float)):
                ws_live.cell(r, 6).value = float(g_val)
                total_gold_tonnes += float(g_val)
            gold_m_row = (
                (float(g_val) * OZ_PER_TONNE * current_xau) / 1_000_000.0
                if isinstance(g_val, (int, float))
                else 0.0
            )
            ws_live.cell(r, 7).value = gold_m_row
            tot_row = (float(fx_val) if isinstance(fx_val, (int, float)) else 0.0) + gold_m_row
            ws_live.cell(r, 8).value = tot_row
            ws_live.cell(r, 9).value = (gold_m_row / tot_row) if tot_row > 0 else 0.0
            ws_live.cell(r, 10).value = (
                (gold_m_row / float(fx_val))
                if isinstance(fx_val, (int, float)) and float(fx_val) > 0
                else 0.0
            )

    total_gold_m = (total_gold_tonnes * OZ_PER_TONNE * current_xau) / 1_000_000.0
    gold_fx_ratio = (total_gold_m / total_fx_m) if total_fx_m > 0 else 0.0
    gold_total_ratio = (
        (total_gold_m / (total_fx_m + total_gold_m))
        if (total_fx_m + total_gold_m) > 0
        else 0.0
    )

    # ۴. ثبت ردیف جدید در Time_Series_Tracker
    ws_ts = wb["Time_Series_Tracker"]
    now_utc = datetime.now(timezone.utc)
    stamp_str = now_utc.strftime("%Y-%m-%d %H:%M")

    # پیدا کردن آخرین ردیف دارای مقدار در ستون B (ستون 2)
    last_data_row = 9
    for r in range(10, ws_ts.max_row + 1):
        val_b = ws_ts.cell(r, 2).value
        if val_b is not None and str(val_b).strip() != "":
            last_data_row = r

    new_r = last_data_row + 1
    ws_ts.cell(new_r, 2).value = stamp_str
    ws_ts.cell(new_r, 3).value = round(total_fx_m, 2)
    ws_ts.cell(new_r, 4).value = round(total_gold_tonnes, 2)
    ws_ts.cell(new_r, 5).value = round(current_xau, 2)
    ws_ts.cell(new_r, 6).value = f"=(D{new_r}*(1000000/31.1034768)*E{new_r})/1000000"
    ws_ts.cell(new_r, 7).value = f"=IF(C{new_r}>0,F{new_r}/C{new_r},0)"
    ws_ts.cell(new_r, 8).value = f"=IF((C{new_r}+F{new_r})>0,F{new_r}/(C{new_r}+F{new_r}),0)"
    ws_ts.cell(new_r, 9).value = f'=IF(G{new_r}<=0.15,"<= 15% (Target Met)","> 15% (Above Target)")'
    ws_ts.cell(new_r, 10).value = "ثبت خودکار ابری (GitHub Actions - 2x/Day)"

    # گسترش محدوده جدول tbl_Time_Series اگر وجود داشته باشد
    if "tbl_Time_Series" in ws_ts.tables:
        ws_ts.tables["tbl_Time_Series"].ref = f"B9:J{new_r}"

    # ۵. استخراج سری زمانی برای رسم نمودار
    dates_list, ratio_list, xau_list = [], [], []
    for r in range(10, new_r + 1):
        d_val = str(ws_ts.cell(r, 2).value or "").strip()
        fx_v = ws_ts.cell(r, 3).value
        ton_v = ws_ts.cell(r, 4).value
        xau_v = ws_ts.cell(r, 5).value
        if d_val and isinstance(fx_v, (int, float)) and isinstance(ton_v, (int, float)) and isinstance(xau_v, (int, float)):
            g_m = (float(ton_v) * OZ_PER_TONNE * float(xau_v)) / 1_000_000.0
            rt = (g_m / float(fx_v)) * 100.0 if float(fx_v) > 0 else 0.0
            dates_list.append(d_val)
            ratio_list.append(rt)
            xau_list.append(float(xau_v))

    wb.save(filepath)

    return {
        "timestamp": stamp_str,
        "fx_total_m": total_fx_m,
        "gold_tonnes": total_gold_tonnes,
        "gold_val_m": total_gold_m,
        "xau_price": current_xau,
        "gold_fx_ratio": gold_fx_ratio,
        "gold_total_ratio": gold_total_ratio,
        "dates": dates_list,
        "ratios": ratio_list,
        "xau_series": xau_list,
    }


def build_chart(summary, out_png):
    if not summary["dates"]:
        return False

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(9, 6), sharex=True)
    x_idx = range(len(summary["dates"]))

    ax1.plot(x_idx, summary["ratios"], marker="o", linewidth=2, color="#D97706", label="Gold / FX (%)")
    ax1.axhline(15.0, color="#16A34A", linestyle="--", linewidth=1.3, label="Target (15%)")
    ax1.set_ylabel("Gold / FX (%)")
    ax1.set_title("Global Gold-to-FX Ratio & XAU Price Trend")
    ax1.grid(True, linestyle=":", alpha=0.6)
    ax1.legend(loc="best")

    ax2.plot(x_idx, summary["xau_series"], marker="s", linewidth=2, color="#1B365D", label="XAU/USD ($/oz)")
    ax2.set_ylabel("XAU Price ($/oz)")
    ax2.set_xticks(list(x_idx))
    ax2.set_xticklabels(summary["dates"], rotation=30, ha="right", fontsize=8)
    ax2.grid(True, linestyle=":", alpha=0.6)
    ax2.legend(loc="best")

    plt.tight_layout()
    plt.savefig(out_png, dpi=150)
    plt.close(fig)
    return True


def send_telegram_report(summary, excel_path, chart_path):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("[Info] Telegram credentials not set; skipping Telegram delivery.")
        return

    ratio_pct = summary["gold_fx_ratio"] * 100.0
    total_pct = summary["gold_total_ratio"] * 100.0
    target_status = "✅ زیر هدف ۱۵٪" if ratio_pct <= 15.0 else "⏳ بالاتر از هدف ۱۵٪"

    text = (
        f"📊 *گزارش خودکار نسبت طلا به ذخایر ارزی*\n"
        f"🕒 زمان ثبت (UTC): `{summary['timestamp']}`\n\n"
        f"💵 مجموع ذخایر ارزی: *${summary['fx_total_m']:,.0f}M*\n"
        f"⚖️ تناژ طلای بانک‌های مرکزی: *{summary['gold_tonnes']:,.1f} تن*\n"
        f"🥇 ارزش دلاری ذخایر طلا: *${summary['gold_val_m']:,.0f}M*\n"
        f"📈 قیمت اونس طلا (XAU): *${summary['xau_price']:,.2f}*\n\n"
        f"🎯 *نسبت Gold/FX:* `{ratio_pct:.2f}%` ({target_status})\n"
        f"🧩 *سهم طلا از کل ذخایر:* `{total_pct:.2f}%`"
    )

    requests.post(
        f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage",
        json={"chat_id": TELEGRAM_CHAT_ID, "text": text, "parse_mode": "Markdown"},
        timeout=20,
    )

    if os.path.exists(chart_path):
        with open(chart_path, "rb") as img_f:
            requests.post(
                f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendPhoto",
                data={"chat_id": TELEGRAM_CHAT_ID, "caption": "📉 نمودار روند شاخص و اونس طلا"},
                files={"photo": img_f},
                timeout=30,
            )

    if os.path.exists(excel_path):
        with open(excel_path, "rb") as xl_f:
            requests.post(
                f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument",
                data={"chat_id": TELEGRAM_CHAT_ID, "caption": "📁 فایل اکسل به‌روزرسانی‌شده"},
                files={"document": xl_f},
                timeout=30,
            )


def main():
    summary = update_workbook(EXCEL_FILE)
    has_chart = build_chart(summary, CHART_FILE)
    send_telegram_report(summary, EXCEL_FILE, CHART_FILE if has_chart else "")
    print("Done:", summary["timestamp"], f"Gold/FX={summary['gold_fx_ratio']*100:.2f}%")


if __name__ == "__main__":
    main()