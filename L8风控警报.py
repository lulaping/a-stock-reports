# -*- coding: utf-8 -*-
"""
L8 风控警报 — 上证3900点自动监测
=================================
功能：下午盘中持续监控上证指数，若跌破3900点则发出多级警报。
用法：
  python L8风控警报.py              # 默认模式（13:00-15:00 自动监测）
  python L8风控警报.py --test       # 测试模式（立即监测一次，不等待时间窗口）
  python L8风控警报.py --forever    # 全时段监测（不限下午，用于盘中测试）

依赖：pip install requests
"""
import requests
import time
import datetime
import os
import sys
import json
import winsound
import subprocess

H = {"User-Agent": "Mozilla/5.0", "Referer": "https://quote.eastmoney.com/"}
QUOTE_URL = "https://push2.eastmoney.com/api/qt/ulist.np/get"
SLEEP = 30  # 轮询间隔（秒）
THRESHOLD = 3900  # 上证阈值

# ─── L8 风控规则 ───
L8_RULES = """
╔══════════════════════════════════════════════════════╗
║              L8 风控警报 — 上证跌破3900              ║
╠══════════════════════════════════════════════════════╣
║ 触发条件: 上证指数 < 3900                           ║
║ 当前背景: 2026-08-19 退潮日(涨停29/跌停35/冰点11分) ║
╠══════════════════════════════════════════════════════╣
║ 一级行动: 已持仓 → 立即清仓                          ║
║ 二级行动: 空仓 → 禁止抄底，禁止左侧买入               ║
║ 三级行动: 关闭所有短线交易窗口，转为观望              ║
║ 四级行动: 明日竞价前不新增任何候选                   ║
╚══════════════════════════════════════════════════════╝
"""


def fetch_index():
    """获取上证指数实时数据（含重试）"""
    last_err = None
    for attempt in range(3):
        try:
            r = requests.get(QUOTE_URL, params={
                "secids": "1.000001",
                "fields": "f2,f3,f4,f12,f14,f62"
            }, headers=H, timeout=15)
            j = r.json()
            q = (j.get("data") or {}).get("diff") or [{}]
            q = q[0]
            price = (q.get("f2") or 0) / 100
            pct = (q.get("f3") or 0) / 100
            vol = (q.get("f4") or 0) / 10000
            fund = (q.get("f62") or 0) / 100000000  # 亿元
            return {"price": price, "pct": pct, "vol": vol, "fund": fund}
        except Exception as e:
            last_err = e
            time.sleep(2)
    return None


def show_alert(msg, level="danger"):
    """多级警报输出"""
    # 1. 控制台警报（带颜色标记）
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    icon = {"danger": "!!", "warning": "!!", "info": "ii"}.get(level, "!!")
    print(f"\n{'='*60}")
    print(f"{icon} [{ts}] {msg}")
    print(f"{'='*60}")

    # 2. 声音警报
    try:
        if level == "danger":
            winsound.Beep(800, 500)
            time.sleep(0.2)
            winsound.Beep(1000, 500)
            time.sleep(0.2)
            winsound.Beep(1200, 500)
        elif level == "warning":
            winsound.Beep(800, 300)
            time.sleep(0.1)
            winsound.Beep(800, 300)
        elif level == "info":
            winsound.Beep(600, 200)
    except:
        pass

    # 3. Windows 弹窗通知
    try:
        # 使用 PowerShell 弹出通知
        ps_script = f'''
        [System.Windows.MessageBox]::Show('{msg}', 'L8 风控警报', 'OK', 'Warning')
        '''
        subprocess.Popen(["powershell", "-Command", ps_script],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except:
        pass


def print_l8_banner():
    """打印L8风控规则横幅"""
    print(L8_RULES)


def is_trading_hours():
    """判断是否在交易时段内"""
    now = datetime.datetime.now()
    if now.weekday() >= 5:
        return False  # 周末
    t = now.hour * 100 + now.minute
    # 上午 9:30-11:30, 下午 13:00-15:00
    if 930 <= t <= 1130:
        return True
    if 1300 <= t <= 1500:
        return True
    return False


def is_afternoon():
    """判断是否在下午时段"""
    now = datetime.datetime.now()
    t = now.hour * 100 + now.minute
    return 1300 <= t <= 1500


def main():
    test_mode = "--test" in sys.argv
    forever_mode = "--forever" in sys.argv or "--now" in sys.argv

    print(f"{'='*60}")
    print(f"  L8 风控警报 — 上证3900点自动监测")
    print(f"  启动时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  监测阈值: 上证指数 < {THRESHOLD}")
    print(f"  轮询间隔: {SLEEP}秒")
    print(f"  模式: {'测试模式(单次)' if test_mode else '全时段' if forever_mode else '下午时段自动监测'}")
    print(f"{'='*60}\n")

    if not test_mode and not forever_mode:
        if not is_afternoon():
            now = datetime.datetime.now()
            if now.hour < 13:
                wait_min = (13 - now.hour) * 60 - now.minute
                print(f"  当前 {now.strftime('%H:%M')}，下午 13:00 开始自动监测")
                print(f"  等待约 {wait_min} 分钟后启动...\n")

    triggered = False
    first_alert = True

    while True:
        now = datetime.datetime.now()

        # 判断是否在监测窗口
        if test_mode:
            pass  # 单次执行
        elif forever_mode:
            pass  # 全时段
        elif not is_afternoon():
            if now.hour < 13:
                # 上午等待到下午
                remaining = (13 - now.hour) * 3600 - now.minute * 60 - now.second
                if remaining > 0:
                    if first_alert:
                        print(f"  [{now.strftime('%H:%M:%S')}] 等待下午开盘...")
                        first_alert = False
                    time.sleep(min(60, remaining))
                    continue
            else:
                # 收盘后退出
                print(f"\n  [{now.strftime('%H:%M:%S')}] 下午收盘，监测结束。")
                break

        if test_mode:
            # 单次测试
            idx = fetch_index()
            if idx:
                price = idx["price"]
                print(f"  上证指数: {price:.2f} ({idx['pct']:+.2f}%) 资金净流入{idx['fund']:.0f}亿")
                if price < THRESHOLD:
                    show_alert(f"上证指数 {price:.2f} 跌破 {THRESHOLD}！跌幅 {idx['pct']:+.2f}%", "danger")
                    print_l8_banner()
                else:
                    print(f"  上证 {price:.2f} > {THRESHOLD}，安全。")
            else:
                print("  [!] 获取指数失败，请检查网络。")
            break

        # 正常监测
        idx = fetch_index()
        if idx:
            price = idx["price"]
            if price < THRESHOLD and not triggered:
                triggered = True
                show_alert(
                    f"上证指数 {price:.2f} 跌破 {THRESHOLD} 点！跌幅 {idx['pct']:+.2f}%\n"
                    f"资金净流出 {idx['fund']:.0f}亿\n"
                    f"→ 触发L8风控：立即执行清仓/禁止抄底/关闭交易窗口",
                    "danger"
                )
                print_l8_banner()
            elif price >= THRESHOLD and triggered:
                # 反弹回3900以上，解除警报
                triggered = False
                show_alert(
                    f"上证指数 {price:.2f} 重新站上 {THRESHOLD} 点，警报解除。\n"
                    f"但L8退潮确认仍然有效，建议维持空仓观望。",
                    "info"
                )

            # 每10轮输出一次状态
            if int(time.time()) % (SLEEP * 10) < SLEEP:
                status = "!! 已触发" if triggered else "OK 安全"
                print(f"  [{now.strftime('%H:%M:%S')}] 上证 {price:.2f} ({idx['pct']:+.2f}%) 流出{idx['fund']:.0f}亿 | {status}")
        else:
            print(f"  [{now.strftime('%H:%M:%S')}] [!] 获取指数失败，10秒后重试...")
            time.sleep(10)
            continue

        time.sleep(SLEEP)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n\n  L8 风控警报已手动停止。")
        print(f"  [{datetime.datetime.now().strftime('%H:%M:%S')}] 监测结束。")