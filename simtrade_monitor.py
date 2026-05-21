import shioaji as sj
import time
import requests
import urllib.parse
from datetime import datetime
import os
import pytz
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ================= 配置設定 =================
API_KEY    = os.environ.get("SHIOAJI_API_KEY",    "")
SECRET_KEY = os.environ.get("SHIOAJI_SECRET_KEY", "")
IS_SIMULATION    = os.environ.get("SHIOAJI_SIMULATION", "false").lower() == "true"
VOLUME_THRESHOLD = 100    # 試撮量門檻 (張)
LIMIT_ALERT_PCT  = 0.02   # 距漲跌停 2% 以內觸發警示
SURGE_ALERT_PCT  = 8.0    # 漲跌幅超過此值（%）時特別標注

# 預設監測族群（可透過 GitHub Actions 手動觸發時的 INPUT_CATEGORIES 覆蓋）
DEFAULT_CATEGORIES = ["24","25","26","27","28","29","30","31","32","21","03","13","23"]

TZ_TW = pytz.timezone("Asia/Taipei")
MARKET_CLOSE_HOUR   = 13
MARKET_CLOSE_MINUTE = 35

_bark_env = os.environ.get("BARK_KEYS", "")
BARK_KEYS = [k.strip() for k in _bark_env.split(",") if k.strip()]

last_push_time    = {}
stock_state       = {}
today_sim_records = []

api = sj.Shioaji(simulation=IS_SIMULATION)


# ─────────────── 工具函式 ─────────────────────────────────

def send_bark_alert(title: str, content: str):
    enc_title   = urllib.parse.quote(title)
    enc_content = urllib.parse.quote(content)
    for key in BARK_KEYS:
        url = f"https://api.day.app/{key}/{enc_title}/{enc_content}"
        try:
            requests.get(url, timeout=3)
        except Exception as e:
            print(f"❌ Bark 推送失敗: {e}")


def tick_type_str(tick_type: int) -> str:
    return {1: "外盤", 2: "內盤"}.get(tick_type, "不明")


def check_near_limit(price: float, limit_up, limit_down) -> str:
    if limit_up   and price >= limit_up   * (1 - LIMIT_ALERT_PCT):
        return "漲停注意"
    if limit_down and price <= limit_down * (1 + LIMIT_ALERT_PCT):
        return "跌停注意"
    return ""


def calc_change_pct(price: float, reference) -> float | None:
    if reference and reference > 0:
        return round((price - reference) / reference * 100, 2)
    return None


def format_change_pct(pct) -> str:
    if pct is None:
        return "N/A"
    sign = "+" if pct >= 0 else ""
    tag  = " 🚨大幅異動" if abs(pct) >= SURGE_ALERT_PCT else ""
    return f"{sign}{pct:.2f}%{tag}"


def _init_state(code: str, limit_up=None, limit_down=None, reference=None):
    stock_state[code] = {
        # ── 正常盤資訊（試撮前最後一筆）──
        "last_normal_price"     : None,   # 進試搓前的最後一個價格
        "last_normal_tick_type" : 0,
        "last_normal_total_vol" : 0,
        # ── 試撮狀態 ──
        "in_sim"                : False,
        "sim_start_time"        : None,
        "sim_first_price"       : None,   # 進試搓後的第一個價格
        "sim_last_price"        : None,   # 結束試搓前的最後一個價格
        "sim_total_vol"         : 0,
        # ── 參考價格 ──
        "limit_up"              : limit_up,
        "limit_down"            : limit_down,
        "reference"             : reference,
        # ── 警示 ──
        "near_limit"            : "",
        "change_pct"            : None,
    }


# ─────────────── Tick 處理 ────────────────────────────────

@api.on_tick_stk_v1()
def on_tick_handler(exchange, tick):
    code     = tick.code
    time_int = tick.datetime.hour * 100 + tick.datetime.minute
    is_trading_time = (900 <= time_int < 1325)

    if code not in stock_state:
        _init_state(code)

    state = stock_state[code]

    # ── 正常成交（非試撮）────────────────────────────────
    if not tick.simtrade:
        if state["in_sim"]:
            # 試撮結束 → 記錄「結束試搓後的第一個價格」
            state["in_sim"] = False
            record = {
                "code"            : code,
                "start_time"      : state["sim_start_time"],
                "end_time"        : tick.datetime.strftime("%H:%M:%S"),
                # 四個價格欄位
                "pre_sim_price"   : state["last_normal_price"],   # 進試搓前最後價
                "sim_first_price" : state["sim_first_price"],      # 試搓後首價
                "sim_last_price"  : state["sim_last_price"],       # 試搓末價（結束前最後一筆）
                "post_sim_price"  : tick.close,                    # 結束後首價
                # 其他資訊
                "tick_type"       : tick_type_str(state["last_normal_tick_type"]),
                "pre_total_vol"   : state["last_normal_total_vol"],
                "sim_vol"         : state["sim_total_vol"],
                "near_limit"      : state["near_limit"],
                "change_pct"      : state["change_pct"],
            }
            today_sim_records.append(record)
            print(
                f"📝 [{code}] 試撮結束 | "
                f"試搓末:{state['sim_last_price']:.2f} → 結束後首:{tick.close:.2f} | "
                f"{record['end_time']}"
            )

        # 更新正常盤資訊
        state["last_normal_price"]     = tick.close
        state["last_normal_tick_type"] = getattr(tick, "tick_type", 0)
        state["last_normal_total_vol"] = getattr(tick, "total_volume", 0)
        return

    # ── 試撮 (simtrade=True) ──────────────────────────────
    if not is_trading_time or tick.volume < VOLUME_THRESHOLD:
        return

    near_limit = check_near_limit(tick.close, state["limit_up"], state["limit_down"])
    change_pct = calc_change_pct(tick.close, state["reference"])
    is_surge   = change_pct is not None and abs(change_pct) >= SURGE_ALERT_PCT

    if not state["in_sim"]:
        # ── 試撮開始：記錄「進試搓後的第一個價格」──
        state["in_sim"]          = True
        state["sim_start_time"]  = tick.datetime.strftime("%H:%M:%S")
        state["sim_first_price"] = tick.close    # 試搓後首價（固定不變）
        state["sim_last_price"]  = tick.close    # 試搓末價（持續更新）
        state["sim_total_vol"]   = tick.volume
        state["near_limit"]      = near_limit
        state["change_pct"]      = change_pct

        pre_price     = state["last_normal_price"]
        pre_type      = tick_type_str(state["last_normal_tick_type"])
        pre_vol       = state["last_normal_total_vol"]
        pre_price_str = f"{pre_price:.2f}" if pre_price is not None else "無前置"
        pct_str       = format_change_pct(change_pct)

        tags = []
        if near_limit:
            tags.append(near_limit)
        if is_surge:
            tags.append("🚨大幅異動")
        tag_str = "　" + "　".join(tags) if tags else ""

        msg = (
            f"{code} 試撮:{tick.close:.2f} 漲跌:{pct_str} 量:{tick.volume}張{tag_str}\n"
            f"前價:{pre_price_str} {pre_type} 累積量:{pre_vol}張"
        )
        print(f"🔥 【試撮警報】[{state['sim_start_time']}] {msg}")

        now = time.time()
        if code not in last_push_time or (now - last_push_time[code] > 60):
            if is_surge:
                bark_title = f"🚨大幅異動試撮 {code} {pct_str}"
            elif near_limit:
                bark_title = f"試撮警報　{near_limit}"
            else:
                bark_title = "試撮警報"
            send_bark_alert(bark_title, msg)
            last_push_time[code] = now

    else:
        # ── 試撮持續中：更新試搓末價（結束前的最後一筆會留在這）──
        state["sim_last_price"]  = tick.close    # 每次都更新，最後留下的就是末價
        state["sim_total_vol"]  += tick.volume
        if near_limit:
            state["near_limit"] = near_limit
        if change_pct is not None:
            state["change_pct"] = change_pct


# ─────────────── 強制結束懸空試撮 ────────────────────────

def finalize_ongoing_sims():
    """
    匯出前掃描所有 in_sim=True 的股票。
    這些試撮有開始卻沒有等到後續的非試撮 tick，
    直接補上紀錄，結束後首價標記為 None（未知）。
    """
    now_str = datetime.now(TZ_TW).strftime("%H:%M:%S")
    count   = 0
    for code, state in stock_state.items():
        if not state["in_sim"]:
            continue
        record = {
            "code"            : code,
            "start_time"      : state["sim_start_time"],
            "end_time"        : f"{now_str}※",   # ※ 代表非正常結束
            "pre_sim_price"   : state["last_normal_price"],
            "sim_first_price" : state["sim_first_price"],
            "sim_last_price"  : state["sim_last_price"],
            "post_sim_price"  : None,             # 未收到結束後的非試撮 tick
            "tick_type"       : tick_type_str(state["last_normal_tick_type"]),
            "pre_total_vol"   : state["last_normal_total_vol"],
            "sim_vol"         : state["sim_total_vol"],
            "near_limit"      : state["near_limit"],
            "change_pct"      : state["change_pct"],
        }
        today_sim_records.append(record)
        state["in_sim"] = False
        count += 1
        print(f"⚠️  [{code}] 試撮未正常結束，已補上紀錄（結束後首價未知）")
    if count:
        print(f"📋 共補上 {count} 筆懸空試撮紀錄")


# ─────────────── Excel 匯出 ───────────────────────────────

def export_to_excel() -> str:
    today_str  = datetime.now(TZ_TW).strftime("%Y%m%d")
    script_dir = os.path.dirname(os.path.abspath(__file__))
    filepath   = os.path.join(script_dir, f"試撮紀錄_{today_str}.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "試撮紀錄"

    headers = [
        "股票代碼",
        "試撮開始", "試撮結束",
        "進試搓前末價",    # 進試搓前的最後一個價格
        "試搓後首價",      # 進試搓後的第一個價格
        "試搓末價",        # 結束試搓前的最後一個價格
        "結束後首價",      # 結束試搓後的第一個價格
        "漲跌幅%",
        "最後盤型",
        "試撮前累積量(張)",
        "試撮量(張)",
        "漲跌停警示",
    ]
    col_widths = [10, 12, 12, 14, 12, 12, 12, 16, 10, 18, 12, 12]

    hdr_fill   = PatternFill("solid", fgColor="1F4E79")
    hdr_font   = Font(color="FFFFFF", bold=True, size=11)
    alt_fill   = PatternFill("solid", fgColor="DEEAF1")
    wht_fill   = PatternFill("solid", fgColor="FFFFFF")
    surge_fill = PatternFill("solid", fgColor="FFE699")
    surge_font = Font(color="C00000", bold=True)

    for col_idx, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell           = ws.cell(row=1, column=col_idx, value=h)
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.row_dimensions[1].height = 22

    for row_idx, r in enumerate(today_sim_records, 2):
        pct      = r.get("change_pct")
        pct_str  = format_change_pct(pct)
        is_surge = pct is not None and abs(pct) >= SURGE_ALERT_PCT

        def fp(v):
            return f"{v:.2f}" if v is not None else "-"

        values = [
            r["code"],
            r["start_time"],       r["end_time"],
            fp(r["pre_sim_price"]),
            fp(r["sim_first_price"]),
            fp(r["sim_last_price"]),
            fp(r["post_sim_price"]),
            pct_str,
            r["tick_type"],
            r["pre_total_vol"],
            r["sim_vol"],
            r["near_limit"],
        ]
        row_fill = alt_fill if row_idx % 2 == 0 else wht_fill
        for col_idx, val in enumerate(values, 1):
            cell           = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.fill      = row_fill
            cell.alignment = Alignment(horizontal="center")

        # 漲跌幅欄（第8欄）
        pct_cell = ws.cell(row=row_idx, column=8)
        if is_surge:
            pct_cell.fill = surge_fill
            pct_cell.font = surge_font
        elif pct is not None:
            color = "C00000" if pct >= 0 else "375623"
            pct_cell.font = Font(color=color, bold=True)

        # 漲跌停警示欄（第12欄）
        if r["near_limit"]:
            ws.cell(row=row_idx, column=12).font = Font(color="FF0000", bold=True)

    ws.freeze_panes = "A2"
    wb.save(filepath)
    print(f"\n📊 Excel 報表已儲存：{filepath}")
    return filepath


# ─────────────── 動態標的篩選 ────────────────────────────

def get_target_categories() -> list[str]:
    """
    族群代碼優先順序：
    1. GitHub Actions 手動觸發時的 INPUT_CATEGORIES 環境變數
    2. 程式碼預設值 DEFAULT_CATEGORIES
    """
    env_val = os.environ.get("INPUT_CATEGORIES", "").strip()
    if env_val:
        cats = [c.strip() for c in env_val.split(",") if c.strip()]
        print(f"📌 使用自訂族群：{cats}")
        return cats
    print(f"📌 使用預設族群：{DEFAULT_CATEGORIES}")
    return DEFAULT_CATEGORIES


def get_dynamic_market_list(api):
    official_excluded = []
    req_headers = {"User-Agent": "Mozilla/5.0"}

    try:
        urls = [
            "https://www.twse.com.tw/exchangeReport/TWTB4U?response=json",
            "https://www.twse.com.tw/exchangeReport/TWT11U?response=json",
        ]
        for url in urls:
            res  = requests.get(url, headers=req_headers, timeout=10)
            data = res.json()
            if "data" in data:
                official_excluded.extend([row[0].split(" ")[0] for row in data["data"]])
        print("📊 官方異常名單同步完成")
    except Exception as e:
        print(f"⚠️ 官方名單讀取失敗: {e}")

    target_categories = get_target_categories()
    MANUAL_BLACKLIST  = []

    candidate_contracts = []
    for contract in api.Contracts.Stocks.TSE:
        if contract.code in MANUAL_BLACKLIST or contract.code in official_excluded:
            continue
        if len(contract.code) != 4:
            continue
        if contract.category not in target_categories:
            continue
        if contract.day_trade != sj.constant.DayTrade.Yes:
            continue
        if hasattr(contract, "special_type") and contract.special_type != 0:
            continue
        candidate_contracts.append(contract)

    final_codes = []
    limit_info  = {}
    print(f"📈 正在分析 {len(candidate_contracts)} 檔標的的價格...")

    for i in range(0, len(candidate_contracts), 100):
        batch     = candidate_contracts[i:i+100]
        snapshots = api.snapshots(batch)
        for s in snapshots:
            if not (s.close and 15 <= s.close <= 300):
                continue

            # ── 修正：用 or None 避免欄位為 0 被當成有效值 ──
            ref = getattr(s, "reference",  None) or None
            lu  = getattr(s, "limit_up",   None) or None
            ld  = getattr(s, "limit_down", None) or None

            # 若 API 未直接提供，從昨收計算
            if ref and not lu:
                lu = round(ref * 1.1, 2)
            if ref and not ld:
                ld = round(ref * 0.9, 2)

            final_codes.append(s.code)
            limit_info[s.code] = (lu, ld, ref)

    print(f"✅ 篩選完成，共 {len(final_codes)} 檔符合條件")
    return final_codes[:254], limit_info


# ─────────────── 主程式 ───────────────────────────────────

def start_monitoring():
    api.login(api_key=API_KEY, secret_key=SECRET_KEY)

    now_str = datetime.now(TZ_TW).strftime("%H:%M:%S")
    send_bark_alert("系統公告", f"監控程式已於 {now_str} 成功啟動！")
    print("✅ 登入成功！")

    final_monitor_list, limit_info = get_dynamic_market_list(api)
    print(f"🚀 啟動監控！實際監控標的：{len(final_monitor_list)} 檔")

    for code in final_monitor_list:
        lu, ld, ref = limit_info.get(code, (None, None, None))
        _init_state(code, limit_up=lu, limit_down=ld, reference=ref)

    for code in final_monitor_list:
        contract = api.Contracts.Stocks[code]
        api.quote.subscribe(contract, quote_type=sj.constant.QuoteType.Tick)

    last_heartbeat_time = 0
    try:
        while True:
            now    = time.time()
            tw_now = datetime.now(TZ_TW)

            if (tw_now.hour > MARKET_CLOSE_HOUR or
                    (tw_now.hour == MARKET_CLOSE_HOUR and tw_now.minute >= MARKET_CLOSE_MINUTE)):
                print(f"\n🔔 [{tw_now.strftime('%H:%M:%S')}] 已過收盤時間，自動結束監控")
                raise SystemExit(0)

            if now - last_heartbeat_time >= 300:
                print(
                    f"💓 [心跳] {tw_now.strftime('%H:%M:%S')} "
                    f"監控中... 已記錄 {len(today_sim_records)} 筆試撮"
                )
                last_heartbeat_time = now
            time.sleep(1)

    except (KeyboardInterrupt, SystemExit):
        print("\n👋 監控結束，正在匯出報表...")
        finalize_ongoing_sims()   # ← 先補上所有懸空的試撮
        if today_sim_records:
            filepath = export_to_excel()
            send_bark_alert("試撮報表", f"今日記錄 {len(today_sim_records)} 筆，報表已儲存")
        else:
            print("📭 今日無試撮紀錄")
            send_bark_alert("試撮監控", "今日無符合條件的試撮紀錄")
        print("正在登出...")
        api.logout()


if __name__ == "__main__":
    start_monitoring()
