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
# 優先讀取環境變數（GitHub Actions Secrets），否則使用預設值
API_KEY    = os.environ.get("SHIOAJI_API_KEY",    "")
SECRET_KEY = os.environ.get("SHIOAJI_SECRET_KEY", "")
IS_SIMULATION    = os.environ.get("SHIOAJI_SIMULATION", "false").lower() == "false"
VOLUME_THRESHOLD = 100   # 試撮量門檻 (張)
LIMIT_ALERT_PCT  = 0.02  # 距漲跌停 2% 以內觸發警示

TZ_TW = pytz.timezone("Asia/Taipei")
MARKET_CLOSE_HOUR   = 13
MARKET_CLOSE_MINUTE = 35  # 13:35 後自動結束

# 從環境變數讀取，多支手機用逗號分隔（例如："key1,key2"）
_bark_env = os.environ.get("BARK_KEYS", "")
BARK_KEYS = [k.strip() for k in _bark_env.split(",") if k.strip()]

# ── 內部狀態 ──────────────────────────────────────────────
last_push_time   = {}   # code -> 上次推播 timestamp
stock_state      = {}   # code -> dict (見下方說明)
today_sim_records = []  # 今日所有試撮紀錄
# ──────────────────────────────────────────────────────────
# stock_state[code] keys:
#   last_normal_price       : 試撮前最後一筆正常成交價
#   last_normal_tick_type   : 試撮前最後一筆 tick_type (1=外盤 2=內盤 0=不明)
#   last_normal_total_vol   : 試撮前最後一筆當日累積成交量
#   in_sim                  : 目前是否處於試撮中
#   sim_start_time          : 試撮開始時間字串
#   sim_price               : 最新試撮價格
#   sim_total_vol           : 本次試撮累積量
#   limit_up / limit_down   : 漲停 / 跌停價
#   near_limit              : 漲跌停警示字串

api = sj.Shioaji(simulation=IS_SIMULATION)


# ─────────────── 工具函式 ─────────────────────────────────

def send_bark_alert(title: str, content: str):
    """同時推送到兩支手機"""
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
    """回傳漲跌停警示，距漲停/跌停 LIMIT_ALERT_PCT 以內才觸發"""
    if limit_up   and price >= limit_up   * (1 - LIMIT_ALERT_PCT):
        return "漲停注意"
    if limit_down and price <= limit_down * (1 + LIMIT_ALERT_PCT):
        return "跌停注意"
    return ""


def _init_state(code: str, limit_up=None, limit_down=None):
    stock_state[code] = {
        "last_normal_price"     : None,
        "last_normal_tick_type" : 0,
        "last_normal_total_vol" : 0,
        "in_sim"                : False,
        "sim_start_time"        : None,
        "sim_price"             : None,
        "sim_total_vol"         : 0,
        "limit_up"              : limit_up,
        "limit_down"            : limit_down,
        "near_limit"            : "",
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
        # 若先前處於試撮中 → 試撮結束，記錄一筆
        if state["in_sim"]:
            state["in_sim"] = False
            record = {
                "code"         : code,
                "start_time"   : state["sim_start_time"],
                "end_time"     : tick.datetime.strftime("%H:%M:%S"),
                "sim_price"    : state["sim_price"],
                "end_price"    : tick.close,
                "tick_type"    : tick_type_str(state["last_normal_tick_type"]),
                "pre_total_vol": state["last_normal_total_vol"],
                "sim_vol"      : state["sim_total_vol"],
                "near_limit"   : state["near_limit"],
            }
            today_sim_records.append(record)
            print(
                f"📝 [{code}] 試撮結束 | "
                f"結束價:{tick.close:.2f} | {record['end_time']}"
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

    if not state["in_sim"]:
        # ── 試撮開始 ──
        state["in_sim"]         = True
        state["sim_start_time"] = tick.datetime.strftime("%H:%M:%S")
        state["sim_price"]      = tick.close
        state["sim_total_vol"]  = tick.volume
        state["near_limit"]     = near_limit

        pre_price = state["last_normal_price"]
        pre_type  = tick_type_str(state["last_normal_tick_type"])
        pre_vol   = state["last_normal_total_vol"]
        pre_price_str = f"{pre_price:.2f}" if pre_price is not None else "無前置"

        limit_tag = f"　{near_limit}" if near_limit else ""
        msg = (
            f"{code} 試撮:{tick.close:.2f} 量:{tick.volume}張{limit_tag}\n"
            f"前價:{pre_price_str} {pre_type} 累積量:{pre_vol}張"
        )
        print(f"🔥 【試撮警報】[{state['sim_start_time']}] {msg}")

        # 手機推播（60 秒冷卻）
        now = time.time()
        if code not in last_push_time or (now - last_push_time[code] > 60):
            bark_title = f"試撮警報{limit_tag}"
            send_bark_alert(bark_title, msg)
            last_push_time[code] = now

    else:
        # ── 試撮持續中：累計量、更新最新試撮價、更新警示 ──
        state["sim_total_vol"] += tick.volume
        state["sim_price"]      = tick.close
        if near_limit:
            state["near_limit"] = near_limit


# ─────────────── Excel 匯出 ───────────────────────────────

def export_to_excel() -> str:
    today_str = datetime.now(TZ_TW).strftime("%Y%m%d")
    # 本機執行 → 存到程式所在資料夾；GitHub Actions → 存到工作目錄
    script_dir = os.path.dirname(os.path.abspath(__file__))
    filepath   = os.path.join(script_dir, f"試撮紀錄_{today_str}.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "試撮紀錄"

    headers    = ["股票代碼", "試撮開始", "試撮結束", "試撮價格", "結束價格",
                  "最後盤型", "試撮前累積量(張)", "試撮量(張)", "漲跌停警示"]
    col_widths = [10,         13,         13,         10,         10,
                  10,         18,                     12,          12]

    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    hdr_font = Font(color="FFFFFF", bold=True, size=11)
    alt_fill = PatternFill("solid", fgColor="DEEAF1")
    wht_fill = PatternFill("solid", fgColor="FFFFFF")

    # 標題列
    for col_idx, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.row_dimensions[1].height = 22

    # 資料列
    for row_idx, r in enumerate(today_sim_records, 2):
        values = [
            r["code"],          r["start_time"],   r["end_time"],
            r["sim_price"],     r["end_price"],    r["tick_type"],
            r["pre_total_vol"], r["sim_vol"],      r["near_limit"],
        ]
        row_fill = alt_fill if row_idx % 2 == 0 else wht_fill
        for col_idx, val in enumerate(values, 1):
            cell           = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.fill      = row_fill
            cell.alignment = Alignment(horizontal="center")

        # 漲跌停警示標紅
        if r["near_limit"]:
            ws.cell(row=row_idx, column=9).font = Font(color="FF0000", bold=True)

    ws.freeze_panes = "A2"
    wb.save(filepath)
    print(f"\n📊 Excel 報表已儲存：{filepath}")
    return filepath


# ─────────────── 動態標的篩選 ────────────────────────────

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

    target_categories = ["24","25","26","27","28","29","30","31","32","21","03","13","23"]
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
    limit_info  = {}  # code -> (limit_up, limit_down)
    print(f"📈 正在分析 {len(candidate_contracts)} 檔標的的價格...")

    for i in range(0, len(candidate_contracts), 100):
        batch     = candidate_contracts[i:i+100]
        snapshots = api.snapshots(batch)
        for s in snapshots:
            if s.close and 15 <= s.close <= 300:
                final_codes.append(s.code)
                # 優先取 snapshot 直接提供的漲跌停，否則從 reference 計算
                lu  = getattr(s, "limit_up",   None)
                ld  = getattr(s, "limit_down",  None)
                ref = getattr(s, "reference",   None)
                if lu is None and ref:
                    lu = round(ref * 1.1, 2)
                if ld is None and ref:
                    ld = round(ref * 0.9, 2)
                limit_info[s.code] = (lu, ld)

    return final_codes[:254], limit_info


# ─────────────── 主程式 ───────────────────────────────────

def start_monitoring():
    api.login(api_key=API_KEY, secret_key=SECRET_KEY)

    now_str = datetime.now().strftime("%H:%M:%S")
    send_bark_alert("系統公告", f"監控程式已於 {now_str} 成功啟動！")
    print("✅ 登入成功！")

    final_monitor_list, limit_info = get_dynamic_market_list(api)
    print(f"🚀 啟動監控！實際監控標的：{len(final_monitor_list)} 檔")

    # 預先初始化所有股票狀態（帶入漲跌停）
    for code in final_monitor_list:
        lu, ld = limit_info.get(code, (None, None))
        _init_state(code, limit_up=lu, limit_down=ld)

    for code in final_monitor_list:
        contract = api.Contracts.Stocks[code]
        api.quote.subscribe(contract, quote_type=sj.constant.QuoteType.Tick)

    last_heartbeat_time = 0
    try:
        while True:
            now        = time.time()
            tw_now     = datetime.now(TZ_TW)

            # 自動結束：13:35 台灣時間（收盤後）
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
