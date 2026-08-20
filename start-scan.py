"""
全市场涨停扫描 + 连板梯队监控脚本
=====================================
功能：
  1. 全市场涨停池扫描（首板/连板/炸板）
  2. 连板梯队晋级监控（昨日N板 → 今日N+1板 = 晋级）
  3. 断板监控（昨日涨停今日消失 = 断板）
  4. 板块异动检测（同板块涨停家数突增）
  5. 命中信号 → Windows 弹窗 + 微信/飞书推送

使用方法：
  1. 配置 WEBHOOK_URL（企业微信/飞书，可选）
  2. 交易日上午 python start-scan.py      # 持续运行
  3. python start-scan.py --once          # 单次扫描（测试）
  4. python start-scan.py --test-notify   # 测试通知通道

依赖：pip install requests
"""
import requests
import json
import time
import datetime
import subprocess
import sys
import os
import re
from typing import Dict, List, Optional, Set, Tuple

# Windows 控制台 GBK 编码兼容：强制 UTF-8 输出
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")

# ============================================================
# 配置区 — 按需修改
# ============================================================

# 推送 webhook（二选一或都填，留空则只弹窗）
WECHAT_WEBHOOK = ""   # 企业微信机器人 webhook，如 https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=xxx
FEISHU_WEBHOOK = ""   # 飞书机器人 webhook，如 https://open.feishu.cn/open-apis/bot/v2/hook/xxx
WEBHOOK_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "start_scan_webhook.txt")

def load_webhook_from_file():
    """从本地配置文件读取 webhook 地址（避免 token 写入公开脚本）
    格式：FEISHU:https://...  /  WECHAT:https://...
    无前缀则默认飞书
    """
    global FEISHU_WEBHOOK, WECHAT_WEBHOOK
    try:
        if os.path.exists(WEBHOOK_FILE):
            with open(WEBHOOK_FILE, "r", encoding="utf-8") as f:
                line = f.read().strip()
            if not line or not line.startswith("http"):
                # 带平台前缀格式
                if ":" in line and line.split(":", 1)[1].startswith("http"):
                    kind, url = line.split(":", 1)
                    if kind.upper() == "WECHAT":
                        WECHAT_WEBHOOK = url
                    else:
                        FEISHU_WEBHOOK = url
            else:
                FEISHU_WEBHOOK = line
    except Exception:
        pass

load_webhook_from_file()

# 监控参数
POLL_INTERVAL = 10        # 轮询间隔（秒）
POOL_URL = "https://push2ex.eastmoney.com/getTopicZTPool"
UT = "7eea3edcaed734bea9cbfc24409ed989"
LOG_FILE = "start_scan_log.txt"

# L8 过滤参数
EXCLUDE_ST = True         # 剔除 ST
MAX_TURNOVER = 20.0       # 换手率 >20% 剔除（L8 强制）
MIN_SEAL_AMOUNT = 0.3e8   # 封单资金下限（元），低于此不报（弱封）
SECTOR_MIN_COUNT = 3      # 板块异动阈值：同板块 ≥3 家涨停

# 放量拉升监控参数（盘中涨幅榜）
SURGE_PCT = 5.5           # 涨幅 ≥5.5% 触发（未涨停）
BREAK_PCT = 6.0           # 涨幅突破 6% 即推送（无条件，放量拉升之外的更强信号）
SURGE_VOL_RATIO = 1.5     # 量比 ≥1.5（放量标准，东财量比）
SURGE_TURNOVER_MIN = 2.0  # 换手率 ≥2%（排除无量一字微涨）
SURGE_URL = "https://push2.eastmoney.com/api/qt/clist/get"
UT2 = "bd1d9ddb04089700cf9c27f6f7426281"

# ============================================================
# 涨停池 API
# ============================================================

def fetch_limit_pool(date: str = "") -> List[dict]:
    """拉取涨停池。date=YYYYMMDD，空则取当日。返回 list[dict]"""
    if not date:
        date = datetime.date.today().strftime("%Y%m%d")
    params = {
        "ut": UT,
        "dpt": "wz.ztzt",
        "Pageindex": 0,
        "pagesize": 1000,
        "sort": "fund:asc",
        "date": date,
    }
    try:
        resp = requests.get(POOL_URL, params=params, timeout=8)
        data = resp.json()
        if data.get("data") and data["data"].get("pool"):
            return data["data"]["pool"]
    except Exception as e:
        log(f"[ERROR] 涨停池拉取失败: {e}")
    return []


def parse_stock(item: dict) -> dict:
    """解析涨停池单条数据"""
    def _t(t):
        """93001 → 9:30:01"""
        s = str(t)
        s = s.zfill(6)
        return f"{s[0:2]}:{s[2:4]}:{s[4:6]}"
    name = item.get("n", "")
    return {
        "code": item.get("c", ""),
        "name": name,
        "price": item.get("p", 0) / 1000.0 if item.get("p") else 0,
        "pct": item.get("zdp", 0),
        "turnover": item.get("hs", 0),
        "board": item.get("lbc", 1),          # 连板数
        "amount": item.get("amount", 0),      # 成交额
        "seal_amount": item.get("fund", 0),   # 封单资金
        "first_seal": _t(item.get("fbt", 0)), # 首次封板时间
        "last_seal": _t(item.get("lbt", 0)),  # 最后封板时间
        "zbc": item.get("zbc", 0),            # 炸板次数
        "sector": item.get("hybk", ""),       # 行业板块
        "ltsz": item.get("ltsz", 0),          # 流通市值
        "is_st": "ST" in name.upper(),
    }


def is_limit_up_first(pool: List[dict]) -> List[dict]:
    """标记是否一字板（首次封板在9:30:30前且无炸板）"""
    result = []
    for it in pool:
        it["_yizi"] = (it["first_seal"] <= "09:30:30" and it["zbc"] == 0
                       and it["turnover"] < 3.0)
        result.append(it)
    return result


# ============================================================
# 盘中涨幅榜（放量拉升监控）
# ============================================================

def fetch_surge_list() -> List[dict]:
    """拉取全市场涨幅榜，返回涨幅 ≥SURGE_PCT% 且未涨停的个股（含量比/换手）"""
    params = {
        "pn": 1, "pz": 300, "po": 1, "np": 1, "fltt": 2, "invt": 2,
        "fid": "f3", "fs": "m:0+t:6,m:0+t:80,m:1+t:2,m:1+t:23,m:0+t:81+s:2048",
        "fields": "f2,f3,f8,f10,f12,f14,f22",
        "ut": UT2, "_": int(time.time() * 1000),
    }
    try:
        resp = requests.get(SURGE_URL, params=params, timeout=8)
        data = resp.json()
        result = []
        if data.get("data") and data["data"].get("diff"):
            for it in data["data"]["diff"]:
                pct = it.get("f3", 0)
                if pct is None or pct < SURGE_PCT:
                    continue
                code = str(it.get("f12", ""))
                name = it.get("f14", "")
                # 判断是否涨停（主板10% / 创业板/科创20% / 北交所30%）
                limit = 30 if code.startswith(("4", "8", "92")) else \
                        20 if code.startswith(("30", "68")) else 10
                if pct >= limit - 0.05:
                    continue  # 已涨停，交给涨停池报
                result.append({
                    "code": code,
                    "name": name,
                    "pct": pct,
                    "price": it.get("f2", 0),
                    "turnover": it.get("f8", 0) or 0,
                    "vol_ratio": it.get("f10", 0) or 0,   # 东财量比
                    "speed": it.get("f22", 0) or 0,       # 涨速
                    "is_st": "ST" in name.upper(),
                })
        return result
    except Exception as e:
        log(f"[ERROR] 涨幅榜拉取失败: {e}")
        return []


# ============================================================
# 昨日基线（连板梯队）
# ============================================================

def get_prev_trade_date() -> str:
    """向前推算上一交易日（跳过周末），返回 YYYYMMDD"""
    d = datetime.date.today() - datetime.timedelta(days=1)
    for _ in range(7):
        if d.weekday() < 5:  # 周一至周五
            return d.strftime("%Y%m%d")
        d -= datetime.timedelta(days=1)
    return d.strftime("%Y%m%d")


def build_yesterday_chain() -> Dict[str, dict]:
    """获取昨日连板梯队 {code: {board, name, sector}}"""
    date_str = get_prev_trade_date()
    pool = fetch_limit_pool(date_str)
    result = {}
    for it in pool:
        s = parse_stock(it)
        result[s["code"]] = s
    log(f"[BASE] 昨日({date_str})涨停池基线 {len(result)} 只")
    return result


# ============================================================
# 通知
# ============================================================

def log(msg: str):
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except:
        pass


def push_webhook(title: str, content: str):
    """推送企业微信/飞书 webhook"""
    if WECHAT_WEBHOOK:
        try:
            payload = {
                "msgtype": "markdown",
                "markdown": {"content": f"**{title}**\n{content}"},
            }
            requests.post(WECHAT_WEBHOOK, json=payload, timeout=5)
        except Exception as e:
            log(f"[ERROR] 企业微信推送失败: {e}")
    if FEISHU_WEBHOOK:
        try:
            # 飞书富文本卡片：支持多条信号分行、加粗标题
            payload = {
                "msg_type": "interactive",
                "card": {
                    "header": {
                        "title": {"tag": "plain_text", "content": title},
                        "template": "turquoise",
                    },
                    "elements": [{"tag": "div", "text": {"tag": "lark_md", "content": content}}],
                },
            }
            requests.post(FEISHU_WEBHOOK, json=payload, timeout=5)
        except Exception as e:
            log(f"[ERROR] 飞书推送失败: {e}")


def extract_stock_title(msg: str) -> str:
    """从信号消息中提取「股票名(代码)」作为标题，找不到则回退到信号类型。
    如：'⭐ 晋级2板 志邦家居(603801) 1→2板 ...' → '志邦家居(603801)'"""
    # 匹配 名称(6位代码)
    m = re.search(r"([\u4e00-\u9fa5A-Za-z]{2,10})\((\d{6})\)", msg)
    if m:
        return f"{m.group(1)}({m.group(2)})"
    return ""


def notify(msg: str, title: str = "", level: str = "info"):
    """综合通知：控制台 + 弹窗 + webhook（无提示音）。
    title 为空时自动从消息提取股票名(代码)作为标题，让弹窗/卡片直接看到标的。"""
    log(msg)
    # 生成标题：优先股票名(代码)，否则用信号类型
    if not title:
        stock = extract_stock_title(msg)
        if stock:
            tag = ("⚡" if "急拉" in msg else
                   "📈" if "突破" in msg else
                   "🚀" if "拉升" in msg else
                   "⭐" if "晋级" in msg else
                   "💥" if "断板" in msg else
                   "🆕" if "首板" in msg else "📊")
            title = f"{tag} {stock}"
        else:
            title = "📈 启动信号"
    # 1. Windows 弹窗（异步，不阻塞）
    try:
        safe_msg = msg.replace("'", "").replace('"', "")
        ps = f"[System.Windows.MessageBox]::Show('{safe_msg}', '{title}', 'OK', 'Information')"
        subprocess.Popen(["powershell", "-Command", ps],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except:
        pass
    # 2. webhook 推送
    push_webhook(title, msg)


# ============================================================
# 差异检测
# ============================================================

def detect_changes(curr: List[dict], prev: List[dict],
                   yesterday: Dict[str, dict],
                   reported: Set[str], first_run: bool = False
                   ) -> Tuple[List[str], List[dict]]:
    """
    对比本轮 vs 上轮涨停池，返回 (通知列表, 板块统计)
    first_run=True 表示首轮基线：首板聚合为一条，避免刷屏
    """
    msgs = []
    curr_map = {s["code"]: s for s in curr}
    prev_map = {s["code"]: s for s in prev}

    # 板块统计（当前池内，用于板块异动）
    sector_stats: Dict[str, List[dict]] = {}
    for s in curr:
        if s["sector"]:
            sector_stats.setdefault(s["sector"], []).append(s)

    # 1. 晋级监控（昨日在梯队 + 今日在池 + 板数增加，报一次）
    for code, s in curr_map.items():
        if code in reported:
            continue
        y = yesterday.get(code)
        if y and s["board"] > y["board"]:
            tag = "🥇" if s["board"] >= 3 else "⭐"
            msgs.append(f"{tag} 晋级{s['board']}板 {s['name']}({code}) "
                        f"{y['board']}→{s['board']}板 · 板块:{s['sector']} "
                        f"换手{s['turnover']:.1f}% 封单{s['seal_amount']/1e8:.1f}亿")
            reported.add(code)

    # 2. 断板监控（昨日连板 → 今日不在池，报一次）
    for code, y in yesterday.items():
        if code in curr_map or code in reported:
            continue
        if y["board"] >= 2 and y["board"] <= 10:  # 只报 2 板以上断板
            msgs.append(f"💥 断板 {y['name']}({code}) {y['board']}板 → 断板 · 板块:{y['sector']}")
            reported.add(code)

    # 3. 首板监控（昨日不在梯队 + 今日新涨停）
    first_boards = []
    for code, s in curr_map.items():
        if code in yesterday or code in reported:
            continue
        if s["board"] != 1:            # 只看首板
            continue
        if EXCLUDE_ST and s["is_st"]:
            continue
        if s["turnover"] > MAX_TURNOVER:
            continue
        if s["_yizi"]:
            continue                  # 一字板不刷屏
        first_boards.append(s)
        reported.add(code)

    if first_boards:
        if first_run:
            # 首轮基线：聚合为一条，避免几十条刷屏
            names = "、".join(f"{s['name']}({s['code']})" for s in first_boards[:20])
            more = f"… 共{len(first_boards)}只" if len(first_boards) > 20 else f"共{len(first_boards)}只"
            msgs.append(f"🆕 今日首板 {more}: {names}")
        else:
            for s in first_boards:
                tag = "⚡秒板" if s["first_seal"] <= "09:35:00" else "🆕首板"
                msgs.append(f"{tag} {s['name']}({s['code']}) · 板块:{s['sector']} "
                            f"换手{s['turnover']:.1f}% 封单{s['seal_amount']/1e8:.1f}亿")

    return msgs, sector_stats


def check_sector_surge(sector_stats: Dict[str, List[dict]],
                       sector_reported: Set[str]) -> List[str]:
    """板块异动：某板块涨停家数 ≥ 阈值且未报过"""
    msgs = []
    for sector, stocks in sector_stats.items():
        if len(stocks) >= SECTOR_MIN_COUNT and sector not in sector_reported:
            names = "、".join(f"{s['name']}({s['board']}板)" for s in stocks[:6])
            msgs.append(f"🔥 板块异动【{sector}】{len(stocks)}家涨停: {names}")
            sector_reported.add(sector)
    return msgs


def detect_surge(curr: List[dict], prev: List[dict],
                 reported: Set[str], first_run: bool = False) -> List[str]:
    """盘中涨幅监控（两类信号，去重共用 reported）：
    ① 涨幅突破 BREAK_PCT(6%)：无条件推送（更强信号，最优先）
    ② 放量拉升：5.5% ≤ 涨幅 < 6% 且 量比/换手达标
    prev 为上一轮涨幅榜；reported 用于去重（同一股票只报一次）"""
    msgs = []
    curr_map = {s["code"]: s for s in curr}
    prev_map = {s["code"]: s for s in prev}

    new_breaks = []   # 突破6%信号
    new_surges = []   # 放量拉升信号（5.5%-6%区间）
    for code, s in curr_map.items():
        if code in prev_map or code in reported:
            continue
        if EXCLUDE_ST and s["is_st"]:
            continue
        if s["pct"] >= BREAK_PCT:
            # ① 突破6%：无条件（不要求量比/换手）
            new_breaks.append(s)
            reported.add(code)
        elif (s["pct"] >= SURGE_PCT
              and s["vol_ratio"] >= SURGE_VOL_RATIO
              and s["turnover"] >= SURGE_TURNOVER_MIN):
            # ② 放量拉升（5.5%-6%区间）
            new_surges.append(s)
            reported.add(code)

    if new_breaks:
        if first_run:
            names = "、".join(f"{s['name']}({s['code']})" for s in new_breaks[:15])
            more = f"… 共{len(new_breaks)}只" if len(new_breaks) > 15 else f"共{len(new_breaks)}只"
            msgs.append(f"📈 涨幅突破{BREAK_PCT}% {more}: {names}")
        else:
            for s in new_breaks[:5]:   # 每轮最多报5条，避免刷屏
                tag = "⚡急拉" if s["speed"] >= 3 else "📈突破"
                msgs.append(f"{tag} {s['name']}({s['code']}) {s['pct']:+.1f}% "
                            f"量比{s['vol_ratio']:.1f} 换手{s['turnover']:.1f}% 涨速{s['speed']:+.2f}")

    if new_surges:
        if first_run:
            names = "、".join(f"{s['name']}({s['code']})" for s in new_surges[:15])
            more = f"… 共{len(new_surges)}只" if len(new_surges) > 15 else f"共{len(new_surges)}只"
            msgs.append(f"🚀 放量拉升 {more}(5.5%-6%+量比≥{SURGE_VOL_RATIO}): {names}")
        else:
            for s in new_surges[:5]:
                msgs.append(f"🚀拉升 {s['name']}({s['code']}) {s['pct']:+.1f}% "
                            f"量比{s['vol_ratio']:.1f} 换手{s['turnover']:.1f}% 涨速{s['speed']:+.2f}")
    return msgs


# ============================================================
# 主循环
# ============================================================

def is_trading_time() -> bool:
    """交易时段：9:20-11:35 / 12:55-15:05"""
    now = datetime.datetime.now()
    if now.weekday() >= 5:
        return False
    t = now.hour * 100 + now.minute
    return (920 <= t <= 1135) or (1255 <= t <= 1505)


def is_after_close() -> bool:
    """是否已收盘（15:05 后）"""
    now = datetime.datetime.now()
    t = now.hour * 100 + now.minute
    return t > 1505


def install_task() -> str:
    """安装 Windows 计划任务（每日 9:20 盘中监控 + 15:10 收盘扫描）"""
    py = sys.executable
    script = os.path.abspath(__file__)
    cmd1 = (f'schtasks /Create /F /TN "L8StartScan" '
            f'/SC DAILY /ST 09:20 /TR "\\"{py}\\" \\"{script}\\"" '
            f'/RL LIMITED')
    cmd2 = (f'schtasks /Create /F /TN "L8CloseScan" '
            f'/SC DAILY /ST 15:10 /TR "\\"{py}\\" \\"{script}\\" --once --silent" '
            f'/RL LIMITED')
    out = []
    for cmd in (cmd1, cmd2):
        r = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        out.append(r.stdout.strip() + r.stderr.strip())
    return "\n".join(out)


def main():
    # 安装计划任务
    if "--install-task" in sys.argv:
        print(install_task())
        return

    # 测试通知
    if "--test-notify" in sys.argv:
        notify("✅ 通知通道测试成功！当前全市场涨停扫描已就绪。",
               title="📈 启动扫描器", level="info")
        print("测试通知已发送（若配置了 webhook 会同时推送）。")
        return

    # --once 单次模式：扫描一次并报告（供计划任务收盘后/手动用）
    if "--once" in sys.argv:
        log("=" * 60)
        log("📈 单次扫描模式")
        yesterday = build_yesterday_chain()
        pool = fetch_limit_pool()
        stocks = [parse_stock(it) for it in pool]
        stocks = is_limit_up_first(stocks)
        msgs, sector_stats = detect_changes(stocks, [], yesterday, set(), first_run=True)
        sector_msgs = check_sector_surge(sector_stats, set())
        all_msgs = msgs + sector_msgs
        log(f"今日涨停 {len(stocks)} 只 | 昨日连板 {len(yesterday)} 只")
        for m in all_msgs:
            log(m)
        if all_msgs and "--silent" not in sys.argv:
            push_webhook("📈 收盘扫描", "\n".join(all_msgs[:10]))
        return

    # 初始化昨日连板基线
    log("=" * 60)
    log("📈 全市场涨停扫描器启动")
    log(f"轮询间隔: {POLL_INTERVAL}s | 过滤: ST/换手>{MAX_TURNOVER}%/封单<{MIN_SEAL_AMOUNT/1e8:.1f}亿")
    yesterday = build_yesterday_chain()
    if yesterday:
        notify(f"昨日连板梯队 {len(yesterday)} 只已加载，开始监控晋级/断板/新首板。",
               title="📈 启动扫描器", level="info")

    prev_pool: List[dict] = []
    reported: Set[str] = set()   # 已通知的晋级/断板去重
    sector_reported: Set[str] = set()
    prev_surge: List[dict] = []
    reported_surge: Set[str] = set()  # 已通知的放量拉升去重

    while True:
        # 收盘后自动退出（避免收盘后重复误报断板）
        if is_after_close():
            log("🔚 已收盘（>15:05），扫描结束。")
            return

        # 非交易时段等待
        if not is_trading_time():
            log("⏸ 非交易时段，等待...")
            time.sleep(60)
            continue

        pool = fetch_limit_pool()
        if not pool:
            log("⚠️ 涨停池为空（可能接口限流），重试...")
            time.sleep(5)
            continue

        stocks = [parse_stock(it) for it in pool]
        stocks = is_limit_up_first(stocks)

        # 涨幅榜（放量拉升监控）
        surges = fetch_surge_list()

        # 首次轮询：标记 first_run，首板/拉升聚合为一条上报（不刷屏）
        if not prev_pool:
            prev_pool = stocks
            prev_surge = surges
            msgs, sector_stats = detect_changes(stocks, [], yesterday, reported, first_run=True)
            sector_msgs = check_sector_surge(sector_stats, sector_reported)
            surge_msgs = detect_surge(surges, [], reported_surge, first_run=True)
            for s in stocks:
                reported.add(s["code"])
            for s in surges:
                reported_surge.add(s["code"])
            all_msgs = msgs + sector_msgs + surge_msgs
            log(f"🔄 首轮基线: {len(stocks)} 只涨停, {len(surges)} 只放量拉升(≥{SURGE_PCT}%)")
            if all_msgs:
                for m in all_msgs[:10]:
                    notify(m, level="info")
            time.sleep(POLL_INTERVAL)
            continue

        msgs, sector_stats = detect_changes(stocks, prev_pool, yesterday, reported)
        sector_msgs = check_sector_surge(sector_stats, sector_reported)
        surge_msgs = detect_surge(surges, prev_surge, reported_surge)

        # 更新已报告集合（记录本池出现的代码）
        for s in stocks:
            reported.add(s["code"])
        for s in surges:
            reported_surge.add(s["code"])

        # 输出通知（合并，避免刷屏）
        all_msgs = msgs + sector_msgs + surge_msgs
        if all_msgs:
            for m in all_msgs[:10]:   # 每轮最多报 10 条
                level = "danger" if "断板" in m else "info"
                notify(m, level=level)
            if len(all_msgs) > 10:
                notify(f"…… 另有 {len(all_msgs)-10} 条信号，见 {LOG_FILE}", level="info")

        prev_pool = stocks
        prev_surge = surges
        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    main()
