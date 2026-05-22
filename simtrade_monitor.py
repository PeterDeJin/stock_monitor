import shioaji as sj
import time
import requests
import urllib.parse
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ================= 配置設定 =================
API_KEY    = os.environ.get("SHIOAJI_API_KEY",    "JBkUV2v5c1D2psEVsv9iUnnZkhJvcP5chn7sGPJi6xzq")
SECRET_KEY = os.environ.get("SHIOAJI_SECRET_KEY", "8Wmrwcoz4Gc8r5gWsxg9HiMScPtFvvDp7L9m958DkHAe")
IS_SIMULATION    = True
VOLUME_THRESHOLD = 100   # 試撮量門檻
LIMIT_ALERT_PCT  = 0.02  # 距漲跌停 2% 以內觸發警示

_bark_env = os.environ.get("BARK_KEYS", "")
BARK_KEYS = [k.strip() for k in _bark_env.split(",") if k.strip()] or [
    "TW8hgBWttV99DiyXHLy63c",
    "X2HvykRh99Eb2mionnHkp6",
]

# 自動結束時間（給 GitHub Actions 用，本機按 Ctrl+C 也可以）
MARKET_CLOSE_HOUR   = 13
MARKET_CLOSE_MINUTE = 35

last_push_time    = {}
stock_state       = {}   # code -> 狀態
today_sim_records = []   # 今日所有試撮紀錄
# ===========================================

api = sj.Shioaji(simulation=IS_SIMULATION)


def send_bark_alert(title, content):
    encoded_title   = urllib.parse.quote(title)
    encoded_content = urllib.parse.quote(content)
    for key in BARK_KEYS:
        url = f"https://api.day.app/{key}/{encoded_title}/{encoded_content}"
        try:
            requests.get(url, timeout=3)
        except Exception as e:
            print(f"❌ Bark 推送失敗: {e}")


def on_tick_handler(exchange, tick):
    code     = str(tick.code)
    time_int = tick.datetime.hour * 100 + tick.datetime.minute
    is_trading_time = (900 <= time_int < 1325)

    # 初始化狀態
    if code not in stock_state:
        stock_state[code] = {
            "ref"             : None,   # 昨收
            "limit_up"        : None,
            "limit_down"      : None,
            "pre_sim_price"   : None,   # 進試搓前的最後一個價格
            "in_sim"          : False,
            "sim_first_price" : None,   # 進試搓後的第一個價格
            "sim_last_price"  : None,   # 結束試搓前的最後一個價格
            "sim_start_time"  : None,
            "sim_total_vol"   : 0,
        }
    state = stock_state[code]

    # ─── 非試撮 tick ───
    if not tick.simtrade:
        # 之前是試撮 → 試撮結束
        if state["in_sim"]:
            state["in_sim"] = False
            ref = state["ref"]
            sim_first = state["sim_first_price"]
            change_pct = round((sim_first - ref) / ref * 100, 2) if ref else None

            near_limit = ""
            if state["limit_up"] and sim_first >= state["limit_up"] * (1 - LIMIT_ALERT_PCT):
                near_limit = "漲停注意"
            elif state["limit_down"] and sim_first <= state["limit_down"] * (1 + LIMIT_ALERT_PCT):
                near_limit = "跌停注意"

            today_sim_records.append({
                "code"            : code,
                "sim_start"       : state["sim_start_time"],
                "sim_end"         : tick.datetime.strftime("%H:%M:%S"),
                "pre_sim_price"   : state["pre_sim_price"],     # 進試搓前最後價
                "sim_first_price" : sim_first,                   # 試搓後首價
                "sim_last_price"  : state["sim_last_price"],     # 試搓末價
                "post_sim_price"  : tick.close,                  # 試搓結束後首價
                "change_pct"      : change_pct,
                "near_limit"      : near_limit,
                "volume"          : state["sim_total_vol"],
            })
            print(f"📝 [{code}] 試撮結束 結束後首價:{tick.close:.2f}")

        # 持續更新「進試搓前最後價」
        state["pre_sim_price"] = tick.close
        return

    # ─── 試撮 tick ───
    if not is_trading_time or tick.volume < VOLUME_THRESHOLD:
        return

    if not state["in_sim"]:
        # 試撮開始
        state["in_sim"]          = True
        state["sim_first_price"] = tick.close
        state["sim_last_price"]  = tick.close
        state["sim_start_time"]  = tick.datetime.strftime("%H:%M:%S")
        state["sim_total_vol"]   = tick.volume

        # 計算漲跌幅 + 警示
        ref = state["ref"]
        change_pct = round((tick.close - ref) / ref * 100, 2) if ref else None
        pct_str = f"{change_pct:+.2f}%" if change_pct is not None else "N/A"

        near_limit = ""
        if state["limit_up"] and tick.close >= state["limit_up"] * (1 - LIMIT_ALERT_PCT):
            near_limit = "　漲停注意"
        elif state["limit_down"] and tick.close <= state["limit_down"] * (1 + LIMIT_ALERT_PCT):
            near_limit = "　跌停注意"

        msg = f"{code} | 價:{tick.close:.2f} | 漲跌:{pct_str} | 量:{tick.volume}張{near_limit}"
        print(f"🔥 【試撮警報】[{state['sim_start_time']}] {msg}")

        now = time.time()
        if code not in last_push_time or (now - last_push_time[code] > 60):
            send_bark_alert("台股試撮警報", msg)
            last_push_time[code] = now
    else:
        # 試撮持續中
        state["sim_last_price"] = tick.close
        state["sim_total_vol"] += tick.volume


def get_dynamic_market_list(api):
    """ 篩選：上市、指定產業、15-300元、可當沖、排除異常 """
    official_excluded = []
    headers = {"User-Agent": "Mozilla/5.0"}

    try:
        urls = [
            "https://www.twse.com.tw/exchangeReport/TWTB4U?response=json",
            "https://www.twse.com.tw/exchangeReport/TWT11U?response=json",
        ]
        for url in urls:
            res  = requests.get(url, headers=headers, timeout=10)
            data = res.json()
            if "data" in data:
                official_excluded.extend([row[0].split(" ")[0] for row in data["data"]])
        print("📊 官方異常名單同步完成")
    except Exception as e:
        print(f"⚠️ 官方名單讀取失敗: {e}")

    target_categories = ["24","25","26","27","28","29","30","31","32","21","03","13","23"]
    MANUAL_BLACKLIST  = []

    # 新版 shioaji 直接迭代 TSE 會因為某些合約 code 是 int 而炸掉
    # 改用 keys() 拿代碼清單，再逐一查詢
    tse = api.Contracts.Stocks.TSE
    try:
        tse_codes = [str(k) for k in tse.keys()]
    except Exception:
        tse_codes = [f"{n:04d}" for n in range(1000, 10000)]   # fallback 暴力查詢

    candidate_contracts = []
    for code_str in tse_codes:
        if len(code_str) != 4:
            continue
        if code_str in MANUAL_BLACKLIST or code_str in official_excluded:
            continue
        try:
            contract = tse[code_str]
        except Exception:
            continue
        if contract is None:
            continue
        try:
            if contract.category not in target_categories:
                continue
            if contract.day_trade != sj.constant.DayTrade.Yes:
                continue
            if hasattr(contract, "special_type") and contract.special_type != 0:
                continue
        except Exception:
            continue
        candidate_contracts.append(contract)

    final_codes = []
    snapshot_info = {}   # code -> (ref, limit_up, limit_down)
    print(f"📈 正在分析 {len(candidate_contracts)} 檔標的的價格...")

    for i in range(0, len(candidate_contracts), 100):
        batch     = candidate_contracts[i:i+100]
        snapshots = api.snapshots(batch)
        for s in snapshots:
            if not (s.close and 15 <= s.close <= 300):
                continue

            # ── 求昨收 ──
            ref = getattr(s, "reference", None) or None
            if not ref:
                # 從 change_price 反推（close - 漲跌價差 = 昨收）
                change = getattr(s, "change_price", None)
                if change is not None and s.close:
                    ref = round(s.close - change, 2)

            # ── 求漲跌停 ──
            lu = getattr(s, "limit_up",   None) or None
            ld = getattr(s, "limit_down", None) or None
            if ref and not lu:
                lu = round(ref * 1.1, 2)
            if ref and not ld:
                ld = round(ref * 0.9, 2)

            final_codes.append(str(s.code))
            snapshot_info[str(s.code)] = (ref, lu, ld)

    return final_codes[:254], snapshot_info


def export_to_excel():
    today_str  = datetime.now().strftime("%Y%m%d")
    script_dir = os.path.dirname(os.path.abspath(__file__))
    filepath   = os.path.join(script_dir, f"試撮紀錄_{today_str}.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "試撮紀錄"

    headers = [
        "股票代碼", "試撮開始", "試撮結束",
        "進試搓前末價", "試搓後首價", "試搓末價", "結束後首價",
        "漲跌幅%", "漲跌停警示", "試撮量(張)",
    ]
    col_widths = [10, 12, 12, 14, 12, 12, 12, 12, 12, 12]

    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    hdr_font = Font(color="FFFFFF", bold=True, size=11)
    alt_fill = PatternFill("solid", fgColor="DEEAF1")

    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[1].height = 22

    def fp(v):
        return f"{v:.2f}" if v is not None else "-"

    for r, rec in enumerate(today_sim_records, 2):
        pct = rec["change_pct"]
        pct_str = f"{pct:+.2f}%" if pct is not None else "N/A"
        values = [
            rec["code"], rec["sim_start"], rec["sim_end"],
            fp(rec["pre_sim_price"]), fp(rec["sim_first_price"]),
            fp(rec["sim_last_price"]), fp(rec["post_sim_price"]),
            pct_str, rec["near_limit"], rec["volume"],
        ]
        row_fill = alt_fill if r % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        for c, val in enumerate(values, 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.fill = row_fill
            cell.alignment = Alignment(horizontal="center")

        # 漲跌幅顏色
        if pct is not None:
            color = "C00000" if pct >= 0 else "375623"
            ws.cell(row=r, column=8).font = Font(color=color, bold=True)
        # 漲跌停警示標紅
        if rec["near_limit"]:
            ws.cell(row=r, column=9).font = Font(color="FF0000", bold=True)

    ws.freeze_panes = "A2"
    wb.save(filepath)
    print(f"📊 Excel 報表已儲存：{filepath}")


def start_monitoring():
    api.login(api_key=API_KEY, secret_key=SECRET_KEY)
    api.on_tick_stk_v1()(on_tick_handler)   # 登入後才註冊 callback

    now_str = datetime.now().strftime("%H:%M:%S")
    send_bark_alert("系統公告", f"監控程式已於 {now_str} 成功啟動！")
    print("✅ 登入成功！")

    final_monitor_list, snapshot_info = get_dynamic_market_list(api)
    print(f"🚀 啟動監控！實際監控標的：{len(final_monitor_list)} 檔")

    # 把昨收、漲跌停寫入每檔股票狀態
    for code in final_monitor_list:
        ref, lu, ld = snapshot_info.get(code, (None, None, None))
        stock_state[code] = {
            "ref"             : ref,
            "limit_up"        : lu,
            "limit_down"      : ld,
            "pre_sim_price"   : None,
            "in_sim"          : False,
            "sim_first_price" : None,
            "sim_last_price"  : None,
            "sim_start_time"  : None,
            "sim_total_vol"   : 0,
        }

    for code in final_monitor_list:
        contract = api.Contracts.Stocks[code]
        api.quote.subscribe(contract, quote_type=sj.constant.QuoteType.Tick)

    last_heartbeat_time = 0
    try:
        while True:
            now = time.time()
            tw_now = datetime.now()

            # 自動結束（含 GitHub Actions）
            if (tw_now.hour > MARKET_CLOSE_HOUR or
                    (tw_now.hour == MARKET_CLOSE_HOUR and tw_now.minute >= MARKET_CLOSE_MINUTE)):
                print(f"🔔 [{tw_now.strftime('%H:%M:%S')}] 已過收盤時間，自動結束監控")
                raise SystemExit(0)

            if now - last_heartbeat_time >= 300:
                print(f"💓 [心跳] {tw_now.strftime('%H:%M:%S')} 監控中... 已記錄 {len(today_sim_records)} 筆試撮")
                last_heartbeat_time = now
            time.sleep(1)
    except (KeyboardInterrupt, SystemExit):
        print("👋 監控結束，匯出報表中...")
        if today_sim_records:
            export_to_excel()
            send_bark_alert("試撮報表", f"今日共 {len(today_sim_records)} 筆試撮紀錄")
        else:
            print("📭 今日無試撮紀錄")
            send_bark_alert("試撮監控", "今日無試撮紀錄")
        api.logout()


if __name__ == "__main__":
    start_monitoring()
