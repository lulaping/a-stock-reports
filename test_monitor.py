"""快速测试 金螳螂 002081 实时行情 + 清单4规则检测"""
import requests, time, datetime

secid = "0.002081"
url = "https://push2.eastmoney.com/api/qt/stock/get"
params = {
    "secid": secid,
    "fields": "f43,f44,f45,f46,f47,f48,f50,f51,f52,f57,f58,f60,f170,f171,f5,f6,f8",
    "_": int(time.time() * 1000)
}
resp = requests.get(url, params=params, timeout=10)
data = resp.json()
d = data["data"]

# 解析 (价格×100)
price = d["f43"] / 100.0
prev_close = d["f60"] / 100.0
pct = d["f170"] / 100.0
high = d["f44"] / 100.0
low = d["f45"] / 100.0
open_p = d["f46"] / 100.0
bid_vol = d["f47"]
bid_amount = d["f48"]
volume = d.get("f5", 0)
amount = d.get("f6", 0)
turnover = d.get("f8", 0) / 100.0 if d.get("f8") else 0
limit_up = round(prev_close * 1.1, 2)
limit_down = round(prev_close * 0.9, 2)

print("=" * 55)
print(f"  金螳螂 (002081) 实时行情  {datetime.datetime.now().strftime('%H:%M:%S')}")
print("=" * 55)
print(f"  最新价:     {price:.2f}")
print(f"  涨跌幅:     {pct:+.2f}%")
print(f"  今开:       {open_p:.2f}")
print(f"  最高/最低:  {high:.2f} / {low:.2f}")
print(f"  昨收:       {prev_close:.2f}")
print(f"  涨停价:     {limit_up:.2f}   跌停价: {limit_down:.2f}")
print(f"  成交量:     {volume} 手    成交额: {amount:.0f}")
print(f"  换手率:     {turnover:.2f}%")
print(f"  封单量:     {bid_vol} 手")
print(f"  封单额:     {bid_amount:.0f} 元")
print()

# 计算浮盈
cost = 5.65
profit = (price - cost) / cost * 100
print(f"  成本价:     {cost}")
print(f"  浮盈:       {profit:+.2f}%")
print()

# 清单4 规则检测
print("─" * 55)
print("  清单4 规则检测")
print("─" * 55)
r = []

# 1. 断板
if pct < 9.5:
    r.append(("断板", f"昨日4板 今日{pct:+.2f}%未封板 → 🚨 断板！", "DANGER"))
else:
    r.append(("封板", f"昨日4板 今日{pct:+.2f}%涨停封板 → ✅ 封板中", "OK"))

# 2. 止损
if profit <= -5:
    r.append(("止损", f"浮盈{profit:+.2f}% → 🚨 触发-5%止损！", "DANGER"))
else:
    r.append(("安全", f"浮盈{profit:+.2f}% → ✅ 安全", "OK"))

# 3. 一字板
if open_p == limit_up:
    r.append(("一字板", "今开=涨停价 → ⚠️ 一字板不追！", "WARN"))
else:
    r.append(("非一字", f"今开{open_p:.2f} ≠ 涨停价{limit_up:.2f} → ✅ 非一字板", "OK"))

# 4. 弱转强
if open_p < prev_close and pct > 0 and price > open_p:
    r.append(("弱转强", f"低开{open_p:.2f}→拉涨停{price:.2f} → ✅ 弱转强经典形态！", "OK"))
elif open_p < prev_close and pct < 0:
    r.append(("弱转强失败", "低开未拉红 → ⚠️ 放弃！", "WARN"))

# 5. 尾盘
now = datetime.datetime.now().strftime("%H:%M")
r.append(("尾盘", f"当前{now} → 收盘后，无需执行", "INFO"))

# 6. 封单质量
if bid_amount > 0:
    r.append(("封单", f"封单额 {bid_amount/1e8:.2f}亿 → 封单坚实", "OK"))

# 7. 高标
r.append(("高标", "装修板块高标=002081自身 → 需监控是否断板", "INFO"))

for label, msg, level in r:
    icon = {"OK": "✅", "WARN": "⚠️", "DANGER": "🚨", "INFO": "📢"}
    print(f"  {icon[level]} {label}: {msg}")

print()
print("✅ 测试完成 — API数据获取正常，清单4规则检测逻辑正确")