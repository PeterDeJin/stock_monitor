"""
台股試撮監控程式
─────────────────────────────────────────────
功能：
  1. 監測上市股票試撮（simtrade tick），09:00 後才開始記錄
  2. 觸發試撮警報 → Bark 推播手機
  3. 13:35 自動結束 → 匯出當日 Excel 報表

Excel 欄位：
  日期 / 股票代碼 / 試撮開始 / 試撮結束
  進試搓前末價 / 試搓後首價 / 試搓末價 / 結束後首價
  漲跌幅% = (結束後首價 - 進試搓前末價) / 進試搓前末價 × 100
  最後盤型 = 進試搓前最後一筆是內盤 / 外盤
  試撮前累積量(張) / 試撮量(張) / 漲跌停警示
"""

import shioaji as sj
import requests
import urllib.parse
import os
import time
from datetime import datetime, timezone, timedelta
import openpyxl

# ── 台灣時區（UTC+8）──
TW_TZ = timezone(timedelta(hours=8))

def now_tw() -> datetime:
    """取得台灣當下時間"""
    return datetime.now(TW_TZ)
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ================= 配置 =================
API_KEY    = os.environ.get("SHIOAJI_API_KEY",    "")
SECRET_KEY = os.environ.get("SHIOAJI_SECRET_KEY", "")
IS_SIMULATION = os.environ.get("SHIOAJI_SIMULATION", "true").lower() == "true"

VOLUME_THRESHOLD     = 100   # 試撮量門檻（張）
SURGE_THRESHOLD_PCT  = 8.0   # 日內漲跌幅 ≥ 此值時觸發漲跌停警示

MARKET_OPEN_HOUR      = 9    # 09:00 後才記錄
MARKET_CLOSE_HOUR     = 13
MARKET_CLOSE_MINUTE   = 35   # 13:35 自動結束

_bark_env = os.environ.get("BARK_KEYS", "")
BARK_KEYS = [k.strip() for k in _bark_env.split(",") if k.strip()]

# ── 內部狀態 ──
stock_state       = {}    # code -> dict（每檔股票的試撮狀態）
today_sim_records = []    # 今日所有試撮紀錄
last_push_time    = {}    # code -> 上次推播時間（60 秒冷卻）

api = sj.Shioaji(simulation=IS_SIMULATION)


# ================= 工具函式 =================

def send_bark(title: str, content: str):
    """同時推送到所有 Bark Keys"""
    et = urllib.parse.quote(title)
    ec = urllib.parse.quote(content)
    for k in BARK_KEYS:
        try:
            requests.get(f"https://api.day.app/{k}/{et}/{ec}", timeout=3)
        except Exception as e:
            print(f"❌ Bark 推送失敗: {e}")


def tick_type_str(t) -> str:
    """tick_type: 1=外盤, 2=內盤"""
    return {1: "外盤", 2: "內盤"}.get(t, "不明")


# ================= Tick 處理 =================

def on_tick_handler(*args):
    """
    Tick callback。簽名用 *args 同時相容新舊版 shioaji
    （舊版傳 (exchange, tick)，新版只傳 (tick)）
    """
    tick = args[-1]
    code = str(tick.code)

    if code not in stock_state:
        return
    state = stock_state[code]

    # 09:00 前不記錄
    if tick.datetime.hour < MARKET_OPEN_HOUR:
        # 但要先更新「進試搓前的最後一筆」資訊，09:00 後試撮才有對照
        if not tick.simtrade:
            state["pre_sim_price"]     = tick.close
            state["pre_sim_tick_type"] = getattr(tick, "tick_type", 0)
            state["pre_sim_total_vol"] = getattr(tick, "total_volume", 0)
        return

    # ─── 非試撮 tick ─────────────────────────────
    if not tick.simtrade:
        # 之前是試撮 → 試撮結束，結算紀錄
        if state["in_sim"]:
            state["in_sim"] = False
            _record_sim(state, code, tick.close, tick.datetime)

        # 更新「進試搓前的最後一筆」資訊
        state["pre_sim_price"]     = tick.close
        state["pre_sim_tick_type"] = getattr(tick, "tick_type", 0)
        state["pre_sim_total_vol"] = getattr(tick, "total_volume", 0)
        return

    # ─── 試撮 tick ───────────────────────────────
    if tick.volume < VOLUME_THRESHOLD:
        return

    if not state["in_sim"]:
        # ── 試撮開始 ──
        state["in_sim"]          = True
        state["sim_start_time"]  = tick.datetime.strftime("%H:%M:%S")
        state["sim_first_price"] = tick.close
        state["sim_last_price"]  = tick.close
        state["sim_total_vol"]   = tick.volume

        # 計算日內漲跌幅（vs 昨收），判斷漲跌停警示
        ref = state["ref"]
        day_pct = ((tick.close - ref) / ref * 100) if ref else None

        near_limit = ""
        if day_pct is not None and abs(day_pct) >= SURGE_THRESHOLD_PCT:
            near_limit = f"⚠️ 日內已{day_pct:+.2f}%"
        state["near_limit"] = near_limit

        # 終端訊息
        day_str = f"{day_pct:+.2f}%" if day_pct is not None else "N/A"
        msg = f"{code} 試撮:{tick.close:.2f} 量:{tick.volume}張 日內:{day_str}"
        if near_limit:
            msg += f"\n{near_limit}"
        print(f"🔥 【試撮警報】[{state['sim_start_time']}] {msg}", flush=True)

        # Bark 推播（60 秒冷卻）
        now = time.time()
        if code not in last_push_time or (now - last_push_time[code] > 60):
            title = f"🚨試撮警報 {code}" if near_limit else f"試撮警報 {code}"
            send_bark(title, msg)
            last_push_time[code] = now

    else:
        # ── 試撮持續中 ──
        state["sim_last_price"] = tick.close
        state["sim_total_vol"] += tick.volume


def _record_sim(state, code, post_sim_price, dt):
    """試撮結束 → 寫入紀錄"""
    pre = state["pre_sim_price"]
    impact_pct = None
    if pre and pre > 0:
        impact_pct = round((post_sim_price - pre) / pre * 100, 2)

    today_sim_records.append({
        "date"           : dt.strftime("%Y-%m-%d"),
        "code"           : code,
        "sim_start"      : state["sim_start_time"],
        "sim_end"        : dt.strftime("%H:%M:%S"),
        "pre_sim_price"  : pre,
        "sim_first_price": state["sim_first_price"],
        "sim_last_price" : state["sim_last_price"],
        "post_sim_price" : post_sim_price,
        "impact_pct"     : impact_pct,
        "tick_type"      : tick_type_str(state["pre_sim_tick_type"]),
        "pre_total_vol"  : state["pre_sim_total_vol"],
        "sim_vol"        : state["sim_total_vol"],
        "near_limit"     : state["near_limit"],
    })
    impact_str = f"{impact_pct:+.2f}%" if impact_pct is not None else "N/A"
    print(f"📝 [{code}] 試撮結束 影響:{impact_str}", flush=True)


# ================= 動態標的篩選 =================

def get_market_list(api):
    """回傳 (代碼清單, {code: (昨收, 漲停, 跌停)})"""
    # 官方異常名單（TWSE 半夜或非交易日 API 會回空字串，失敗就略過不影響主流程）
    excluded = []
    for url in [
        "https://www.twse.com.tw/exchangeReport/TWTB4U?response=json",
        "https://www.twse.com.tw/exchangeReport/TWT11U?response=json",
    ]:
        try:
            res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
            data = res.json()
            if "data" in data:
                excluded.extend([row[0].split(" ")[0] for row in data["data"]])
        except Exception:
            pass   # 靜默略過
    if excluded:
        print(f"📊 官方異常名單 {len(excluded)} 檔")

    target_categories = ["24","25","26","27","28","29","30","31","32","21","03","13","23"]

    # 新版 shioaji 直接迭代 TSE 會炸（某些合約 code 是 int）
    # 改用 keys() 拿代碼清單，再逐一查
    tse = api.Contracts.Stocks.TSE
    try:
        all_codes = [str(k) for k in tse.keys()]
    except Exception:
        all_codes = [f"{n:04d}" for n in range(1000, 10000)]

    # 嘗試取新舊版的 DayTrade.Yes
    try:
        day_trade_yes = sj.DayTrade.Yes
    except AttributeError:
        day_trade_yes = sj.constant.DayTrade.Yes

    candidates = []
    for code in all_codes:
        if len(code) != 4 or code in excluded:
            continue
        try:
            c = tse[code]
            if c is None:
                continue
            if c.category not in target_categories:
                continue
            if c.day_trade != day_trade_yes:
                continue
            if getattr(c, "special_type", 0) != 0:
                continue
        except Exception:
            continue
        candidates.append(c)

    print(f"📈 候選 {len(candidates)} 檔，正在抓 snapshot...")

    final_codes = []
    info = {}
    for i in range(0, len(candidates), 100):
        snaps = api.snapshots(candidates[i:i+100])
        for s in snaps:
            if not (s.close and 15 <= s.close <= 300):
                continue

            # 求昨收（reference 拿不到時，用 close - change_price 反推）
            ref = getattr(s, "reference", None) or None
            if not ref:
                change = getattr(s, "change_price", None)
                if change is not None and s.close:
                    ref = round(s.close - change, 2)

            # 求漲跌停
            lu = getattr(s, "limit_up",   None) or None
            ld = getattr(s, "limit_down", None) or None
            if ref and not lu:
                lu = round(ref * 1.1, 2)
            if ref and not ld:
                ld = round(ref * 0.9, 2)

            final_codes.append(str(s.code))
            info[str(s.code)] = (ref, lu, ld)

    return final_codes[:254], info


# ================= Excel 匯出 =================

def export_excel():
    today  = now_tw().strftime("%Y%m%d")
    folder = os.path.dirname(os.path.abspath(__file__))
    path   = os.path.join(folder, f"試撮紀錄_{today}.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "試撮紀錄"

    headers = [
        "日期", "股票代碼", "試撮開始", "試撮結束",
        "進試搓前末價", "試搓後首價", "試搓末價", "結束後首價",
        "漲跌幅%", "最後盤型", "試撮前累積量(張)", "試撮量(張)", "漲跌停警示",
    ]
    widths = [12, 10, 11, 11, 13, 12, 12, 12, 10, 10, 16, 12, 16]

    hdr_fill   = PatternFill("solid", fgColor="1F4E79")
    hdr_font   = Font(color="FFFFFF", bold=True, size=11)
    alt_fill   = PatternFill("solid", fgColor="DEEAF1")
    surge_fill = PatternFill("solid", fgColor="FFE699")

    for c, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill, cell.font = hdr_fill, hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[1].height = 22

    def fp(v):
        return f"{v:.2f}" if v is not None else "-"

    for r, rec in enumerate(today_sim_records, 2):
        pct = rec["impact_pct"]
        pct_str = f"{pct:+.2f}%" if pct is not None else "N/A"
        row_fill = alt_fill if r % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")

        values = [
            rec["date"], rec["code"], rec["sim_start"], rec["sim_end"],
            fp(rec["pre_sim_price"]), fp(rec["sim_first_price"]),
            fp(rec["sim_last_price"]), fp(rec["post_sim_price"]),
            pct_str, rec["tick_type"], rec["pre_total_vol"], rec["sim_vol"],
            rec["near_limit"],
        ]
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.fill = row_fill
            cell.alignment = Alignment(horizontal="center")

        # 漲跌幅顏色
        if pct is not None:
            color = "C00000" if pct >= 0 else "375623"
            ws.cell(row=r, column=9).font = Font(color=color, bold=True)

        # 漲跌停警示
        if rec["near_limit"]:
            ws.cell(row=r, column=13).fill = surge_fill
            ws.cell(row=r, column=13).font = Font(color="C00000", bold=True)

    ws.freeze_panes = "A2"
    wb.save(path)
    print(f"📊 Excel 已儲存：{path}", flush=True)


# ================= 主程式 =================

def main():
    # 登入
    api.login(api_key=API_KEY, secret_key=SECRET_KEY)
    api.on_tick_stk_v1()(on_tick_handler)   # 登入後才註冊 callback

    now_str = now_tw().strftime("%H:%M:%S")
    send_bark("系統公告", f"監控程式於 {now_str} 啟動！")
    print(f"✅ 登入成功 {now_str}", flush=True)

    # 抓監測標的
    codes, info = get_market_list(api)
    print(f"🚀 監測 {len(codes)} 檔", flush=True)

    # 初始化每檔股票狀態
    for code in codes:
        ref, lu, ld = info.get(code, (None, None, None))
        stock_state[code] = {
            "ref"              : ref,
            "limit_up"         : lu,
            "limit_down"       : ld,
            "pre_sim_price"    : None,
            "pre_sim_tick_type": 0,
            "pre_sim_total_vol": 0,
            "in_sim"           : False,
            "sim_start_time"   : None,
            "sim_first_price"  : None,
            "sim_last_price"   : None,
            "sim_total_vol"    : 0,
            "near_limit"       : "",
        }

    # 訂閱（相容新舊版 API）
    try:
        qt = sj.QuoteType.Tick
    except AttributeError:
        qt = sj.constant.QuoteType.Tick

    sub_n = 0
    for code in codes:
        try:
            c = api.Contracts.Stocks[code]
            try:
                api.subscribe(c, quote_type=qt)
            except AttributeError:
                api.quote.subscribe(c, quote_type=qt)
            sub_n += 1
        except Exception as e:
            print(f"⚠️ 訂閱 {code} 失敗: {e}")
    print(f"✅ 已訂閱 {sub_n} 檔", flush=True)

    # 主迴圈
    last_hb = 0
    try:
        while True:
            ts  = time.time()
            now = now_tw()

            # 13:35 自動結束
            if now.hour > MARKET_CLOSE_HOUR or \
               (now.hour == MARKET_CLOSE_HOUR and now.minute >= MARKET_CLOSE_MINUTE):
                print(f"🔔 {now.strftime('%H:%M:%S')} 已過收盤時間，自動結束", flush=True)
                raise SystemExit(0)

            # 每 5 分鐘心跳
            if ts - last_hb >= 300:
                print(f"💓 [{now.strftime('%H:%M:%S')}] 監控中... 已記錄 {len(today_sim_records)} 筆", flush=True)
                last_hb = ts

            time.sleep(1)

    except (KeyboardInterrupt, SystemExit):
        # 補上「懸空試撮」：開始了但沒等到結束 tick
        now_str = now_tw().strftime("%H:%M:%S")
        today_str = now_tw().strftime("%Y-%m-%d")
        for code, state in stock_state.items():
            if state["in_sim"]:
                today_sim_records.append({
                    "date"           : today_str,
                    "code"           : code,
                    "sim_start"      : state["sim_start_time"],
                    "sim_end"        : f"{now_str}※",
                    "pre_sim_price"  : state["pre_sim_price"],
                    "sim_first_price": state["sim_first_price"],
                    "sim_last_price" : state["sim_last_price"],
                    "post_sim_price" : None,
                    "impact_pct"     : None,
                    "tick_type"      : tick_type_str(state["pre_sim_tick_type"]),
                    "pre_total_vol"  : state["pre_sim_total_vol"],
                    "sim_vol"        : state["sim_total_vol"],
                    "near_limit"     : state["near_limit"],
                })

        print(f"📊 共 {len(today_sim_records)} 筆紀錄", flush=True)
        if today_sim_records:
            export_excel()
            send_bark("試撮報表", f"今日 {len(today_sim_records)} 筆，報表已存")
        else:
            send_bark("試撮監控", "今日無試撮紀錄")
        api.logout()


if __name__ == "__main__":
    main()
