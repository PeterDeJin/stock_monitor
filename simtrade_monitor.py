import shioaji as sj
import time
import requests
import urllib.parse
from datetime import datetime

# ================= 配置設定 =================
API_KEY = "JBkUV2v5c1D2psEVsv9iUnnZkhJvcP5chn7sGPJi6xzq"
SECRET_KEY = "8Wmrwcoz4Gc8r5gWsxg9HiMScPtFvvDp7L9m958DkHAe"
IS_SIMULATION = True 
VOLUME_THRESHOLD = 100  # 試撮量門檻 (建議設 500 以上減少雜訊)

# 📱 填入你的兩支手機 Bark Key
BARK_KEYS = [
    "TW8hgBWttV99DiyXHLy63c",
    "X2HvykRh99Eb2mionnHkp6"
]

# 儲存每檔股票最後發送推播的時間 (冷卻機制用)
last_push_time = {}
# ===========================================

api = sj.Shioaji(simulation=IS_SIMULATION)

def send_bark_alert(title, content):
    """ 同時推送到兩支手機，並處理 URL 編碼 """
    encoded_title = urllib.parse.quote(title)
    encoded_content = urllib.parse.quote(content)
    for key in BARK_KEYS:
        url = f"https://api.day.app/{key}/{encoded_title}/{encoded_content}"
        try:
            requests.get(url, timeout=3)
        except Exception as e:
            print(f"❌ Bark 推送失敗: {e}")

@api.on_tick_stk_v1()
def on_tick_handler(exchange, tick):
    time_int = tick.datetime.hour * 100 + tick.datetime.minute
    
    # 交易時間判斷 (09:00 - 13:25)
    is_trading_time = (900 <= time_int < 1325)

    # 判斷：試撮、時間內、量達標
    if tick.simtrade and is_trading_time and tick.volume >= VOLUME_THRESHOLD:
        current_time_str = tick.datetime.strftime('%H:%M:%S')
        msg = f"{tick.code} | 價:{tick.close:.2f} | 量:{tick.volume}張"
        
        # 終端機印出
        print(f"🔥 【試撮警報】[{current_time_str}] {msg}")

        # --- 手機推播邏輯 (含 60 秒冷卻) ---
        now = time.time()
        if tick.code not in last_push_time or (now - last_push_time[tick.code] > 60):
            send_bark_alert("台股試撮警報", msg)
            last_push_time[tick.code] = now

def get_dynamic_market_list(api):
    """ 篩選：上市、指定產業、15-150元、可當沖、排除異常 """
    official_excluded = []
    headers = {'User-Agent': 'Mozilla/5.0'}

    try:
        urls = ["https://www.twse.com.tw/exchangeReport/TWTB4U?response=json", 
                "https://www.twse.com.tw/exchangeReport/TWT11U?response=json"]
        for url in urls:
            res = requests.get(url, headers=headers, timeout=10)
            data = res.json()
            if "data" in data:
                official_excluded.extend([row[0].split(' ')[0] for row in data["data"]])
        print(f"📊 官方異常名單同步完成")
    except Exception as e:
        print(f"⚠️ 官方名單讀取失敗: {e}")

    # 目標產業：電子、航運、油氣
    target_categories = ["24", "25", "26", "27", "28", "29", "30", "31", "32", "21", "03", "13", "23"]
    MANUAL_BLACKLIST = [] # 你提到的台虹

    candidate_contracts = []
    for contract in api.Contracts.Stocks.TSE:
        if contract.code in MANUAL_BLACKLIST or contract.code in official_excluded: continue
        if len(contract.code) != 4: continue 
        if contract.category not in target_categories: continue
        if contract.day_trade != sj.constant.DayTrade.Yes: continue
        if hasattr(contract, 'special_type') and contract.special_type != 0: continue
        candidate_contracts.append(contract)

    final_codes = []
    print(f"📈 正在分析 {len(candidate_contracts)} 檔標的的價格...")
    
    # 分批抓取 Snapshot 以符合 15-150 元篩選
    for i in range(0, len(candidate_contracts), 100):
        batch = candidate_contracts[i:i+100]
        snapshots = api.snapshots(batch)
        for s in snapshots:
            if s.close and 15 <= s.close <= 300:
                final_codes.append(s.code)

    # Shioaji 訂閱限制建議控制在 200 檔以內較穩定
    return final_codes[:254]

def start_monitoring():
    api.login(api_key=API_KEY, secret_key=SECRET_KEY)

    # --- 新增這兩行：啟動時先傳一則訊息到手機測試 ---
    now_str = datetime.now().strftime('%H:%M:%S')
    send_bark_alert("系統公告", f"監控程式已於 {now_str} 成功啟動！")
    # ------------------------------------------

    print(f"✅ 登入成功！當前環境：{'/Users/huangdejin/Desktop/程式集'}")

    final_monitor_list = get_dynamic_market_list(api) 
    print(f"🚀 啟動監控！實際監控標的：{len(final_monitor_list)} 檔")
    
    for code in final_monitor_list:
        contract = api.Contracts.Stocks[code]
        api.quote.subscribe(contract, quote_type=sj.constant.QuoteType.Tick)
    
    last_heartbeat_time = 0
    try:
        while True:
            current_timestamp = time.time()
            if current_timestamp - last_heartbeat_time >= 300:
                print(f"💓 [心跳] {datetime.now().strftime('%H:%M:%S')} 監控運行中...")
                last_heartbeat_time = current_timestamp
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n👋 使用者停止監控，正在登出...")
        api.logout()

if __name__ == "__main__":
    start_monitoring()
