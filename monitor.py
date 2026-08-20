"""
超短持仓监控脚本 — 清单4六时间窗口自动监控
基于东方财富实时行情API，交易日9:25-15:00运行

使用方法：
  1. 修改 HOLDINGS 配置你的持仓
  2. 修改 SECTOR_LEADERS 配置各板块高标
  3. 交易日上午 python monitor.py
  4. Ctrl+C 退出

依赖：pip install requests
"""
import requests
import json
import time
import datetime
import os
import sys
import re
from typing import Dict, List, Optional, Tuple

# Windows 控制台 GBK 编码兼容：强制 UTF-8 输出（emoji 不会报错）
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")

# ============================================================
# 配置区 — 按需修改
# ============================================================

# 持仓股票 (代码, 名称, 成本价, 连板数, 所属板块, 是否高位股)
# ── 加新持仓时照抄一行，改这6个字段即可 ─────────────────────────
#   code    : 6位股票代码
#   name    : 股票名称（仅用于显示）
#   cost    : 你的成本价（用于算浮盈/止损，>5%亏损会触发飞书推送）
#   board   : 连板数（0=非连板，1=首板，4=4连板...；触发断板检测）
#   sector  : 所属板块（须与 SECTOR_LEADERS 里的板块名一致才能做板块退潮检测）
#   high    : 是否高位股。True=启用"冰点冲高≥7%减仓"提醒；False/省略=不启用
#             建议：连板≥3板 或 短期内涨幅大 的持仓标 True
# ───────────────────────────────────────────────────────────────
HOLDINGS = [
    # 宇环数控（2026-08-20 首板涨停 30.67，成本 30.481，低位首板不启用冲高减仓）
    {"code": "002903", "name": "宇环数控", "cost": 30.481, "board": 1, "sector": "通用设备", "high": False},
]

# 板块高标（板块名 → 最高标代码）
SECTOR_LEADERS = {
    "通用设备": "002903",
}

# 高位股冲高减仓规则（情绪冰点时触发）
HIGH_PULLBACK_PCT = 7.0    # 冲高阈值：涨幅≥7%（主板）；创业板/科创自动翻倍为14%
TEMP_ICE_THRESHOLD = 40    # 情绪温度 ≤40（冰点）时该规则生效
TEMP_REFRESH_SEC = 300     # 情绪温度缓存刷新间隔（秒）

# 监控参数
POLL_INTERVAL = 3          # 轮询间隔（秒）
ALERT_SOUND = False        # 是否播放提示音（Windows Beep）
LOG_FILE = "monitor_log.txt"

# 飞书 webhook（复用 start-scan.py 的配置文件，避免重复配置）
WEBHOOK_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "start_scan_webhook.txt")
FEISHU_WEBHOOK = ""
PUSH_EXIT_SIGNALS = True   # 退出信号(止损/断板/炸板/板块退潮)是否推送飞书

def load_webhook_from_file():
    """从 start_scan_webhook.txt 读取飞书 webhook。
    格式：FEISHU:https://... 或 直接 https://...（无前缀默认飞书）"""
    global FEISHU_WEBHOOK
    try:
        if os.path.exists(WEBHOOK_FILE):
            with open(WEBHOOK_FILE, "r", encoding="utf-8") as f:
                line = f.read().strip()
            if not line:
                return
            if line.startswith("http"):
                FEISHU_WEBHOOK = line
            elif ":" in line and line.split(":", 1)[1].startswith("http"):
                kind, url = line.split(":", 1)
                if kind.upper() in ("FEISHU", "FEISHU_WEBHOOK"):
                    FEISHU_WEBHOOK = url
    except Exception:
        pass

load_webhook_from_file()

# ============================================================
# 东方财富实时行情API
# ============================================================

def get_realtime_quotes(codes: List[str]) -> Dict[str, dict]:
    """获取实时行情（东方财富 stock/get 接口，价格×100需除以100）"""
    result = {}
    for code in codes:
        secid = f"{_get_market(code)}.{code}"
        url = "https://push2.eastmoney.com/api/qt/stock/get"
        params = {
            "secid": secid,
            "fields": "f43,f44,f45,f46,f47,f48,f50,f51,f52,f57,f58,f60,f170,f171,f5,f6,f8,f168",
            "_": int(time.time() * 1000)
        }
        try:
            resp = requests.get(url, params=params, timeout=5)
            data = resp.json()
            if data.get("data"):
                d = data["data"]
                price = d.get("f43", 0) / 100.0 if d.get("f43") else 0
                prev_close = d.get("f60", 0) / 100.0 if d.get("f60") else 0
                limit_up_price = round(prev_close * 1.1, 2)  # 10%涨停下
                limit_down_price = round(prev_close * 0.9, 2)
                # 创业板/科创板 20%涨跌停
                if code.startswith(("30", "68")):
                    limit_up_price = round(prev_close * 1.2, 2)
                    limit_down_price = round(prev_close * 0.8, 2)

                result[code] = {
                    "name": d.get("f58", ""),
                    "price": price,
                    "pct": d.get("f170", 0) / 100.0 if d.get("f170") else 0,
                    "high": d.get("f44", 0) / 100.0 if d.get("f44") else 0,
                    "low": d.get("f45", 0) / 100.0 if d.get("f45") else 0,
                    "open": d.get("f46", 0) / 100.0 if d.get("f46") else 0,
                    "volume": d.get("f5", 0),            # 成交量(手)
                    "amount": d.get("f6", 0),            # 成交额
                    "turnover": d.get("f8", 0) / 100.0 if d.get("f8") else 0,  # 换手率%
                    "prev_close": prev_close,
                    "bid_vol": d.get("f47", 0),          # 涨停封单量(手)
                    "bid_amount": d.get("f48", 0),       # 涨停封单额
                    "limit_up": limit_up_price,
                    "limit_down": limit_down_price,
                }
        except Exception as e:
            log(f"[ERROR] 获取{code}行情失败: {e}")
    return result

def _get_market(code: str) -> str:
    """判断交易所：0=深圳, 1=上海"""
    if code.startswith(("6", "68")):
        return "1"
    return "0"

def get_index_quotes() -> dict:
    """获取三大指数"""
    indices = {"1.000001": "上证指数", "0.399001": "深证成指", "0.399006": "创业板指"}
    secids = ",".join(indices.keys())
    url = "https://push2.eastmoney.com/api/qt/ulist.np/get"
    params = {
        "fltt": 2,
        "secids": secids,
        "fields": "f2,f3,f12",
        "_": int(time.time() * 1000)
    }
    try:
        resp = requests.get(url, params=params, timeout=5)
        data = resp.json()
        result = {}
        if data.get("data") and data["data"].get("diff"):
            for item in data["data"]["diff"]:
                name = indices.get(item["f12"], item["f12"])
                result[name] = {
                    "price": item.get("f2", 0),
                    "pct": item.get("f3", 0),
                }
        return result
    except:
        return {}

# ============================================================
# 时间窗口判断
# ============================================================

def get_time_window() -> Tuple[str, str]:
    """返回当前时间窗口和描述"""
    now = datetime.datetime.now()
    t = now.hour * 3600 + now.minute * 60 + now.second

    if t < 9 * 3600 + 15 * 60:
        return "PRE", "盘前等待"
    elif t < 9 * 3600 + 25 * 60 + 30:
        return "BID", "竞价阶段 (9:15-9:25)"
    elif t < 9 * 3600 + 35 * 60:
        return "W1", "窗口1: 弱转强拉红确认 (9:30-9:35)"
    elif t < 10 * 3600:
        return "W2", "窗口2: 封板质量确认 (9:35-10:00)"
    elif t < 10 * 3600 + 30 * 60:
        return "W3", "窗口3: 早盘封板可追 (10:00-10:30)"
    elif t < 14 * 3600 + 30 * 60:
        return "W4", "窗口4: 持有观察 (10:30-14:30)"
    elif t < 15 * 3600:
        return "W5", "窗口5: 尾盘不碰 (14:30-15:00)"
    else:
        return "POST", "收盘后"

def is_trading_time() -> bool:
    """判断是否在交易时段"""
    window, _ = get_time_window()
    return window in ("BID", "W1", "W2", "W3", "W4", "W5")

# ============================================================
# 日志 & 报警
# ============================================================

def log(msg: str):
    """带时间戳日志"""
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def extract_stock_title(msg: str) -> str:
    """从信号消息提取「股票名(代码)」作为标题，找不到则回退空串。
    如：'[断板] 金螳螂(002081) 昨日4板...' → '金螳螂(002081)'"""
    m = re.search(r"([\u4e00-\u9fa5A-Za-z]{2,10})\((\d{6})\)", msg)
    if m:
        return f"{m.group(1)}({m.group(2)})"
    return ""

def push_feishu(title: str, content: str):
    """推送飞书富文本卡片（红色头部，突出退出信号）
    注意：Windows GBK 环境下 json= 参数会因 emoji 编码失败，
    必须手动序列化为 UTF-8 bytes 并显式指定 Content-Type。"""
    if not FEISHU_WEBHOOK:
        return
    try:
        payload = {
            "msg_type": "interactive",
            "card": {
                "header": {
                    "title": {"tag": "plain_text", "content": title},
                    "template": "red",
                },
                "elements": [{"tag": "div", "text": {"tag": "lark_md", "content": content}}],
            },
        }
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        requests.post(
            FEISHU_WEBHOOK,
            data=data,
            headers={"Content-Type": "application/json; charset=utf-8"},
            timeout=5,
        )
        log(f"[飞书] 已推送退出信号: {title}")
    except Exception as e:
        log(f"[ERROR] 飞书推送失败: {e}")

def alert(msg: str, level: str = "WARN"):
    """报警输出：控制台打印 + 飞书推送（仅 DANGER 级别的退出信号推送）"""
    prefix = {"WARN": "⚠️", "DANGER": "🚨", "INFO": "📢", "OK": "✅"}
    p = prefix.get(level, "🔔")
    line = f"{p} {p} {p}  {msg}"
    print("\n" + "=" * 60)
    print(line)
    print("=" * 60 + "\n")
    if ALERT_SOUND:
        try:
            import winsound
            freq = 800 if level == "DANGER" else 500
            winsound.Beep(freq, 500)
        except:
            pass
    # 飞书推送：DANGER = 止损/断板/炸板/板块退潮等退出信号
    if PUSH_EXIT_SIGNALS and level == "DANGER":
        stock = extract_stock_title(msg)
        title = f"🚨 退出信号 {stock}" if stock else "🚨 退出信号"
        push_feishu(title, msg)

# ============================================================
# 市场情绪温度（L8 简化版，盘中实时）
# ============================================================

_TEMP_CACHE = {"time": 0, "temp": None, "zt": 0, "dt": 0, "max_lbc": 0}
_POOL_UT = "7eea3edcaed734bea9cbfc24409ed989"

def fetch_pool(kind: str) -> list:
    """拉取东财涨停(ZT)/跌停(DT)池，失败返回空列表"""
    url = f"https://push2ex.eastmoney.com/getTopic{kind}Pool"
    params = {
        "ut": _POOL_UT, "dpt": "wz.ztzt",
        "Pageindex": 0, "pagesize": 1000, "sort": "fund:asc",
        "date": datetime.date.today().strftime("%Y%m%d"),
    }
    try:
        resp = requests.get(url, params=params, timeout=8)
        data = resp.json()
        if data.get("data") and data["data"].get("pool"):
            return data["data"]["pool"]
    except Exception as e:
        log(f"[ERROR] {kind}池拉取失败: {e}")
    return []

def get_market_temp() -> dict:
    """盘中情绪温度（L8 简化版：涨停数30分/跌停数20分/最高板15分/晋级率按中性10分），
    5分钟缓存。返回 {temp, zone, zt, dt, max_lbc}"""
    now = time.time()
    if now - _TEMP_CACHE["time"] < TEMP_REFRESH_SEC and _TEMP_CACHE["temp"] is not None:
        return _TEMP_CACHE
    zt_pool = fetch_pool("ZT")
    dt_pool = fetch_pool("DT")
    zt = len(zt_pool)
    dt = len(dt_pool)
    max_lbc = max((int(p.get("lbc") or 0) for p in zt_pool), default=0)
    # L8 打分
    s_zt = 30 if zt >= 100 else 20 if zt >= 60 else 10 if zt >= 40 else 0
    s_dt = 20 if dt == 0 else 15 if dt <= 5 else 8 if dt <= 15 else 0
    s_gd = 15 if max_lbc >= 7 else 10 if max_lbc >= 5 else 6 if max_lbc >= 3 else 2
    s_jj = 10  # 盘中晋级率未知，按中性
    temp = s_zt + s_dt + s_gd + s_jj
    zone = "高潮" if temp >= 80 else "正常" if temp >= 40 else "冰点"
    _TEMP_CACHE.update({"time": now, "temp": temp, "zone": zone,
                        "zt": zt, "dt": dt, "max_lbc": max_lbc})
    return _TEMP_CACHE


# ============================================================
# 规则检测
# ============================================================

def check_rules(holdings: list, quotes: dict, indices: dict, 
                prev_window: str, temp_info: dict = None) -> List[str]:
    """检查清单4规则，返回报警消息列表。temp_info: 情绪温度(可选)"""
    alerts = []
    window, desc = get_time_window()

    # 大盘检查（全天）
    for name, idx in indices.items():
        if name == "上证指数" and idx["pct"] < -0.5:
            alerts.append(f"[大盘] {name} {idx['pct']:+.2f}% 翻绿，减仓至3成！")

    temp = temp_info.get("temp") if temp_info else None

    # 逐持仓检查
    for h in holdings:
        code = h["code"]
        q = quotes.get(code)
        if not q:
            continue

        name = q["name"] or h["name"]
        pct = q["pct"]
        price = q["price"]
        open_p = q["open"]
        prev_close = q["prev_close"]
        limit_up = q["limit_up"]
        bid_vol = q["bid_vol"]
        bid_amount = q["bid_amount"]
        volume = q["volume"]
        cost = h["cost"]
        profit_pct = (price - cost) / cost * 100 if cost else 0

        # 铁律：亏损>5%
        if profit_pct <= -5:
            alerts.append(f"[止损] {name}({code}) 亏损{profit_pct:.1f}% 触发无条件止损！")

        # 铁律：断板
        if h["board"] >= 1 and pct < 9.5 and window in ("W2", "W3", "W4", "W5"):
            alerts.append(f"[断板] {name}({code}) 昨日{h['board']}板，今日{pct:+.2f}%未封板，断板！")

        # 高位股冲高减仓（情绪冰点 temp≤40 时，冲高≥7%/14%且未封板 → 无条件减仓至≤2成）
        if (h.get("high") and temp is not None and temp <= TEMP_ICE_THRESHOLD
                and window in ("W1", "W2", "W3", "W4")):
            is_20cm = code.startswith(("30", "68"))
            pull_thresh = HIGH_PULLBACK_PCT * 2 if is_20cm else HIGH_PULLBACK_PCT
            limit_pct = 19.5 if is_20cm else 9.5
            if pct >= pull_thresh and pct < limit_pct:
                alerts.append(f"[冲高减仓] {name}({code}) 冰点(温度{temp})冲高{pct:+.1f}%，减仓至≤2成！")

        # 窗口1: 弱转强确认 (9:30-9:35)
        if window == "W1" and prev_window != "W1":
            if open_p < prev_close and pct > 0 and price > open_p:
                alerts.append(f"[弱转强] {name}({code}) 竞价低开→拉红，可半仓试！")
            elif open_p < prev_close and pct < 0 and price < open_p:
                alerts.append(f"[弱转强失败] {name}({code}) 竞价低开→未拉红，放弃！")

        # 窗口2: 封板质量 (9:35-10:00)
        if window == "W2" and price >= limit_up * 0.995:
            # 封板中
            if volume > 0 and bid_amount > 0:
                ratio = bid_amount / (volume * price) if volume > 0 else 0
                if ratio >= 3:
                    alerts.append(f"[封板] {name}({code}) 封板质量好，封单/成交={ratio:.1f}x，可加仓")
                else:
                    alerts.append(f"[封板弱] {name}({code}) 封单/成交={ratio:.1f}x，观察")

        # 窗口3: 早盘封板 (10:00-10:30)
        if window == "W3" and prev_window != "W3":
            if price >= limit_up * 0.995:
                alerts.append(f"[早盘板] {name}({code}) 10:30前封板，质量好！")

        # 窗口5: 尾盘 (14:30+)
        if window == "W5" and prev_window != "W5":
            if price >= limit_up * 0.995:
                alerts.append(f"[尾盘板] {name}({code}) 14:30后封板，不参与！")
            else:
                alerts.append(f"[尾盘] {name}({code}) 尾盘未封板，注意风险")

        # 全天：涨停被砸开
        if pct < 9 and prev_window in ("W2", "W3", "W4"):
            if price < limit_up * 0.98:
                alerts.append(f"[炸板] {name}({code}) 涨停被砸，封单萎缩！")

    # 板块高标检查（全天）
    for sector, leader_code in SECTOR_LEADERS.items():
        lq = quotes.get(leader_code)
        if lq and lq["pct"] < -5:
            alerts.append(f"[板块] {sector}高标({lq['name']})断板，板块退潮！全清该板块！")

    return alerts

# ============================================================
# 排除竞价潜伏（一字板候选）
# ============================================================

def check_one_zi_board(quotes: dict) -> List[str]:
    """检查一字板 → 提醒不追"""
    alerts = []
    for code, q in quotes.items():
        if q["pct"] >= 9.9 and q["open"] == q["limit_up"]:
            alerts.append(f"[一字板] {q['name']}({code}) 一字板封死，买不进；能买进=开板分歧，不追！")
    return alerts

# ============================================================
# 主循环
# ============================================================

def main():
    log("=" * 50)
    log("超短持仓监控启动")
    log(f"持仓: {len(HOLDINGS)}只 | 轮询间隔: {POLL_INTERVAL}s")
    log(f"板块高标: {len(SECTOR_LEADERS)}个板块")
    log("=" * 50)

    all_codes = [h["code"] for h in HOLDINGS]
    all_codes += list(SECTOR_LEADERS.values())
    all_codes = list(set(all_codes))

    prev_window = ""
    alerted_rules = set()  # 避免重复报警
    last_alert_time = {}   # 每个报警类型的最小间隔

    while True:
        window, desc = get_time_window()

        if window == "PRE":
            # 盘前等待
            now = datetime.datetime.now()
            open_time = now.replace(hour=9, minute=15, second=0, microsecond=0)
            wait_sec = (open_time - now).total_seconds()
            if wait_sec > 0:
                print(f"\r⏳ 盘前等待... 距开盘 {int(wait_sec//60)}分{int(wait_sec%60)}秒", end="")
                time.sleep(min(30, wait_sec))
                continue
        elif window == "POST":
            log("📊 收盘，监控结束")
            break

        # 窗口切换时打印
        if window != prev_window:
            log(f"\n{'='*40}")
            log(f"🕐 进入 {desc}")
            log(f"{'='*40}")
            if window == "W5":
                alert("进入尾盘(14:30)，所有尾盘板不参与！", "WARN")
            alerted_rules.clear()

        # 获取实时数据
        quotes = get_realtime_quotes(all_codes)
        indices = get_index_quotes()
        temp_info = get_market_temp()

        if not quotes:
            print(f"\r⏳ 获取数据中...", end="")
            time.sleep(POLL_INTERVAL)
            continue

        # 打印持仓状态
        print(f"\n{'─'*60}")
        print(f"🕐 {datetime.datetime.now().strftime('%H:%M:%S')}  {desc}")
        print(f"  大盘: ", end="")
        for name, idx in indices.items():
            color = "🟢" if idx["pct"] > 0 else "🔴" if idx["pct"] < 0 else "⚪"
            print(f"{color}{name} {idx['pct']:+.2f}%  ", end="")
        print()
        print(f"  🌡️ 情绪温度: {temp_info['temp']}/100 ({temp_info['zone']})  "
              f"涨停{temp_info['zt']} 跌停{temp_info['dt']} 最高{temp_info['max_lbc']}板")
        for h in HOLDINGS:
            q = quotes.get(h["code"])
            if q:
                profit = (q["price"] - h["cost"]) / h["cost"] * 100 if h["cost"] else 0
                status = "🟢" if q["pct"] > 0 else "🔴"
                print(f"  {status} {q['name']}({h['code']}) {q['price']:.2f} {q['pct']:+.2f}%  "
                      f"浮盈{profit:+.1f}%  换手{q['turnover']:.1f}%")
            else:
                print(f"  ❓ {h['name']}({h['code']}) 数据获取失败")

        # 检查一字板
        for a in check_one_zi_board(quotes):
            alert(a, "INFO")

        # 检查规则
        alerts = check_rules(HOLDINGS, quotes, indices, prev_window, temp_info)
        for a in alerts:
            # 防重复（同类型在同窗口只报一次）
            key = a[:30]
            now_ts = time.time()
            if key not in alerted_rules:
                alerted_rules.add(key)
                if "止损" in a or "断板" in a or "炸板" in a or "板块" in a or "冲高" in a:
                    alert(a, "DANGER")
                elif "弱转强" in a or "封板" in a:
                    alert(a, "OK")
                else:
                    alert(a, "WARN")

        prev_window = window
        time.sleep(POLL_INTERVAL)

    log("监控结束")

if __name__ == "__main__":
    if "--test-notify" in sys.argv:
        # 测试飞书推送通道
        print("=" * 50)
        print("测试飞书推送...")
        if FEISHU_WEBHOOK:
            log(f"webhook: {FEISHU_WEBHOOK[:50]}...")
            push_feishu("🚨 退出信号测试", 
                        "**[测试]** monitor.py 飞书推送已接通\n"
                        "止损/断板/炸板/板块退潮信号将推送到此群\n"
                        f"时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        else:
            print("❌ 未检测到飞书 webhook，请检查 start_scan_webhook.txt 配置")
        sys.exit(0)
    try:
        main()
    except KeyboardInterrupt:
        log("\n用户中断，监控停止")
    except Exception as e:
        log(f"异常退出: {e}")
        import traceback
        traceback.print_exc()