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
from typing import Dict, List, Optional, Tuple

# ============================================================
# 配置区 — 按需修改
# ============================================================

# 持仓股票 (代码, 名称, 成本价, 连板数, 所属板块)
HOLDINGS = [
    {"code": "002081", "name": "金螳螂", "cost": 5.65, "board": 4, "sector": "装修"},
]

# 板块高标（板块名 → 最高标代码）
SECTOR_LEADERS = {
    "装修": "002081",
}

# 监控参数
POLL_INTERVAL = 3          # 轮询间隔（秒）
ALERT_SOUND = False        # 是否播放提示音（Windows Beep）
LOG_FILE = "monitor_log.txt"

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

def alert(msg: str, level: str = "WARN"):
    """报警输出"""
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

# ============================================================
# 规则检测
# ============================================================

def check_rules(holdings: list, quotes: dict, indices: dict, 
                prev_window: str) -> List[str]:
    """检查清单4规则，返回报警消息列表"""
    alerts = []
    window, desc = get_time_window()

    # 大盘检查（全天）
    for name, idx in indices.items():
        if name == "上证指数" and idx["pct"] < -0.5:
            alerts.append(f"[大盘] {name} {idx['pct']:+.2f}% 翻绿，减仓至3成！")

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
        alerts = check_rules(HOLDINGS, quotes, indices, prev_window)
        for a in alerts:
            # 防重复（同类型在同窗口只报一次）
            key = a[:30]
            now_ts = time.time()
            if key not in alerted_rules:
                alerted_rules.add(key)
                if "止损" in a or "断板" in a or "炸板" in a or "板块" in a:
                    alert(a, "DANGER")
                elif "弱转强" in a or "封板" in a:
                    alert(a, "OK")
                else:
                    alert(a, "WARN")

        prev_window = window
        time.sleep(POLL_INTERVAL)

    log("监控结束")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        log("\n用户中断，监控停止")
    except Exception as e:
        log(f"异常退出: {e}")
        import traceback
        traceback.print_exc()