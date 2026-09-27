# -*- coding: utf-8 -*-
"""
南网在线每日电费抓取脚本
================================
功能：根据 config.json 中的防溺水监控电表用户编号 + bindingId 映射，
      调用南网在线 95598.csg.cn 明文接口，抓取每日电量数据，落盘 Excel；
      当日电量 > 阈值 或 日估算电费 > 阈值 时，推送飞书提醒。
    2026-09-26 更新：改为读取 fangnishui_users（24户防溺水监控），
      meteringPointId 传空字符串即可，无需浏览器UI收集。

用法：
    python nw_daily_fetch.py                 # 抓取本月（截至昨天）每日电量
    python nw_daily_fetch.py 202609          # 抓取指定月份
    python nw_daily_fetch.py --price 0.696   # 指定电价估算电费
    python nw_daily_fetch.py --no-alert      # 跳过飞书提醒
    python nw_daily_fetch.py --alert-fee 2.0 --alert-power 3.0  # 覆盖告警阈值
    python nw_daily_fetch.py --config D:/path/config.json   # 指定配置文件（默认脚本同目录 config.json）

数据说明：
    - 接口返回的是【每日用电量】（千瓦时/度），T+3 天发布（例：1月1日数据最晚1月4日可查）
    - 日电费不直接提供，脚本按电价估算；如电价与实际不符，请用 --price 覆盖
    - 数据仅供参考，抄表算费以实际电费单为准

依赖：Python 3.8+（标准库 urllib；openpyxl 可选，无则自动落 CSV）
"""
import argparse
import datetime
import json
import os
import subprocess
import sys
import urllib.error
import urllib.request

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
API_BASE = "https://95598.csg.cn/ucs/ma/wt"
DEFAULT_PRICE = 0.696  # 元/千瓦时（参考均价，可 --price 覆盖）

# 默认提醒阈值：日估算电费 > 1.5 元 或 日用电量 > 1.5 度
DEFAULT_ALERT_FEE = 1.5
DEFAULT_ALERT_POWER = 1.5


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def api_raw(cfg, path, body):
    """调用南网在线接口（明文，x-auth-token 认证），返回解析后的 JSON"""
    url = API_BASE + path
    req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"), method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("x-auth-token", cfg["token"])
    req.add_header("custNumber", cfg.get("custNumber", ""))
    req.add_header("User-Agent",
                   "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36")
    req.add_header("Referer", "https://95598.csg.cn/")
    req.add_header("Accept", "application/json, text/plain, */*")
    with urllib.request.urlopen(req, timeout=20) as resp:
        return json.loads(resp.read().decode("utf-8"))


def query_day_electric(cfg, ele_cust_id, point_id, year_month):
    """
    查询单户某月每日电量
    ele_cust_id: bindingId（config 中 telecom_users.bindingId）
    point_id:    meteringPointId（config 中 telecom_users.pointId）
    返回: [{date:"-09-24", power:"29.15", ...}, ...]
    """
    body = {
        "areaCode": cfg.get("areaCode", "030800"),
        "eleCustId": ele_cust_id,
        "yearMonth": year_month,
        "meteringPointId": point_id,
    }
    try:
        r = api_raw(cfg, "/charge/queryDayElectricAndTemperature", body)
    except urllib.error.HTTPError as e:
        raise RuntimeError("HTTP %s: %s" % (e.code, e.read().decode("utf-8", "ignore")[:200]))
    if r.get("sta") != "00":
        raise RuntimeError("接口错误: %s" % json.dumps(r, ensure_ascii=False)[:200])
    return r.get("data") or []


def fetch_month(cfg, year_month, price):
    """抓取所有防溺水户号某月每日电量，返回明细行"""
    rows = []
    # 优先读取 fangnishui_users（防溺水监控），其次 telecom_users（兼容旧配置）
    users = cfg.get("fangnishui_users", []) or cfg.get("telecom_users", [])
    for u in users:
        bid = u.get("bindingId", "")
        pid = u.get("pointId", "")  # 可传空字符串
        try:
            data = query_day_electric(cfg, bid, pid, year_month)
            for d in data:
                # d["date"] 形如 "-09-24"，直接拼年份： "2026" + "-09-24"
                date_str = year_month[:4] + d.get("date", "")
                power = float(d.get("power") or 0)
                rows.append({
                    "用户编号": u["eleCustNumber"],
                    "用电地址": u.get("eleAddress", ""),
                    "日期": date_str,
                    "用电量(千瓦时)": power,
                    "估算电费(元)": round(power * (price if price is not None else DEFAULT_PRICE), 2),
                    "最低温": d.get("minTemperature", ""),
                    "最高温": d.get("maxTemperature", ""),
                    "均温": d.get("averageTemperature", ""),
                })
            print("  %s  %s  OK %d 天" % (u["eleCustNumber"], u.get("eleAddress", ""), len(data)))
        except Exception as e:
            rows.append({
                "用户编号": u["eleCustNumber"],
                "用电地址": u.get("eleAddress", ""),
                "日期": year_month,
                "用电量(千瓦时)": "",
                "估算电费(元)": "",
                "错误": str(e)[:120],
            })
            print("  %s %s  失败: %s" % (u["eleCustNumber"], u.get("eleAddress", ""), e))
    return rows


def to_excel(rows, out_path):
    """生成 Excel（openpyxl 可选）；无 openpyxl 时落 CSV"""
    try:
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "每日电量"
        if rows:
            headers = list(rows[0].keys())
            ws.append(headers)
            for r in rows:
                ws.append([r.get(h, "") for h in headers])
        wb.save(out_path)
        return "xlsx"
    except ImportError:
        import csv
        csv_path = out_path.replace(".xlsx", ".csv")
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else ["日期"])
            writer.writeheader()
            for r in rows:
                writer.writerow(r)
        return "csv"


def check_token(cfg):
    """
    检测 token 是否有效。
    南网在线接口 sta=04 表示未登录/登录失效。
    返回 (是否有效, 提示信息)
    """
    try:
        r = api_raw(cfg, "/eleCustNumber/queryBindEleUsers", {})
        if r.get("sta") == "00":
            return True, ""
        return False, json.dumps(r, ensure_ascii=False)[:200]
    except Exception as e:
        return False, str(e)[:200]


def show_reminder(title, message):
    """Windows 桌面弹窗提醒（PowerShell MessageBox）"""
    try:
        ps = (
            "[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null;"
            "[System.Windows.Forms.MessageBox]::Show('%s','%s','OK','Warning')" % (message, title)
        )
        subprocess.Popen(["powershell", "-NoProfile", "-Command", ps],
                         creationflags=subprocess.CREATE_NO_WINDOW)
    except Exception:
        pass


def send_feishu_alert(cfg, alerts):
    """
    通过飞书 Webhook 推送提醒消息。
    alerts: 需要提醒的记录列表（每条含 用户编号/用电地址/日期/用电量/估算电费/触发原因）
    """
    url = cfg.get("feishu_webhook_url", "")
    if not url:
        print("config.json 未配置 feishu_webhook_url，跳过飞书推送")
        return

    total = len(alerts)
    fee_over = sum(1 for a in alerts if "电费超标" in a.get("触发原因", ""))
    power_over = sum(1 for a in alerts if "电量超标" in a.get("触发原因", ""))
    both = sum(1 for a in alerts if "电费超标" in a.get("触发原因", "") and "电量超标" in a.get("触发原因", ""))

    lines = []
    for a in alerts[:20]:  # 最多展示 20 条，避免卡片过长
        lines.append(
            "**%s** %s\n"
            "  日期 %s｜电量 %s 度｜电费 %s 元｜%s" % (
                a.get("用户编号", ""),
                a.get("用电地址", ""),
                a.get("日期", ""),
                a.get("用电量(千瓦时)", ""),
                a.get("估算电费(元)", ""),
                a.get("触发原因", ""),
            )
        )
    if len(alerts) > 20:
        lines.append("… 等共 %d 条" % len(alerts))

    today = datetime.date.today().strftime("%Y-%m-%d")
    threshold = cfg.get("alert_threshold", {})
    fee_th = threshold.get("daily_fee", DEFAULT_ALERT_FEE)
    power_th = threshold.get("daily_power", DEFAULT_ALERT_POWER)
    payload = {
        "msg_type": "interactive",
        "card": {
            "header": {
                "title": {"tag": "plain_text", "content": "⚡ 南网电费超标提醒"},
                "template": "red",
            },
            "elements": [
                {
                    "tag": "div",
                    "text": {
                        "tag": "lark_md",
                        "content": (
                            "**日期：** %s\n"
                            "**触发规则：** 日电量 > %.1f 度 或 日估算电费 > %.1f 元\n"
                            "**触发户数：** %d 户（电费超标 %d，电量超标 %d，两者均超 %d）"
                        ) % (today, power_th, fee_th, total, fee_over, power_over, both),
                    },
                },
                {"tag": "hr"},
                {"tag": "div", "text": {"tag": "lark_md", "content": "\n".join(lines)}},
                {"tag": "note", "elements": [{"tag": "plain_text", "content": "数据来源：南网在线（T+3 发布，仅供参考）"}]},
            ],
        },
    }

    req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), method="POST")
    req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            body = resp.read().decode("utf-8")
            result = json.loads(body)
            if result.get("code") == 0:
                print("飞书提醒已发送，共 %d 条" % total)
            else:
                print("飞书推送失败: %s" % body[:200])
    except Exception as e:
        print("飞书推送异常: %s" % e)


def build_alerts(rows, fee_threshold, power_threshold):
    """
    从抓取结果中筛选需要提醒的记录（只看最新日期，避免历史数据全推）。
    规则：日估算电费 > fee_threshold 或 日用电量 > power_threshold。
    返回: 提醒列表；若当日已有提醒标记文件则跳过（防重复推送）。
    """
    if not rows:
        return []

    valid = [r for r in rows if r.get("用电量(千瓦时)") not in (None, "")]
    if not valid:
        return []
    latest_date = max(r["日期"] for r in valid)
    latest = [r for r in valid if r["日期"] == latest_date]

    # 去重标记：同一天只推一次（标记文件放 config 同目录）
    sent_file = os.path.join(os.path.dirname(os.path.abspath(CONFIG_PATH)),
                             ".last_alert_%s.json" % latest_date)
    if os.path.exists(sent_file):
        print("日期 %s 的提醒已发送过，跳过" % latest_date)
        return []

    alerts = []
    for r in latest:
        power = r.get("用电量(千瓦时)", 0) or 0
        fee = r.get("估算电费(元)", 0) or 0
        reasons = []
        if power > power_threshold:
            reasons.append("电量超标")
        if fee > fee_threshold:
            reasons.append("电费超标")
        if reasons:
            r["触发原因"] = "+".join(reasons)
            alerts.append(r)

    if alerts:
        with open(sent_file, "w", encoding="utf-8") as f:
            json.dump({"date": latest_date, "count": len(alerts)}, f, ensure_ascii=False)
    return alerts


def current_month():
    return datetime.datetime.now().strftime("%Y%m")


def main():
    global CONFIG_PATH
    parser = argparse.ArgumentParser(description="南网在线每日电费抓取")
    parser.add_argument("month", nargs="?", help="月份 YYYYMM，缺省为本月")
    parser.add_argument("--price", type=float, default=DEFAULT_PRICE, help="估算电价(元/度)")
    parser.add_argument("--no-alert", action="store_true", help="跳过飞书提醒")
    parser.add_argument("--config", default=CONFIG_PATH, help="config.json 路径（默认脚本同目录）")
    parser.add_argument("--alert-fee", type=float, default=None, help="日估算电费告警阈值(元)，覆盖 config.json")
    parser.add_argument("--alert-power", type=float, default=None, help="日用电量告警阈值(度)，覆盖 config.json")
    args = parser.parse_args()
    CONFIG_PATH = os.path.abspath(args.config)

    cfg = load_config()
    if not cfg.get("token"):
        show_reminder("南网在线抓取失败", "config.json 缺少 token，请登录南网在线后更新")
        print("config.json 缺少 token，请登录南网在线后更新")
        sys.exit(1)

    # Token 有效性检测
    ok, err = check_token(cfg)
    if not ok:
        msg = ("南网在线登录已失效（Token 过期）！\n"
               "请打开 https://95598.csg.cn 重新登录，\n"
               "然后把浏览器里的 token 更新到 config.json。\n\n错误: %s" % err)
        show_reminder("南网在线 Token 失效", msg)
        print("Token 失效: %s" % err)
        sys.exit(2)

    month = args.month or current_month()
    users = cfg.get("fangnishui_users", []) or cfg.get("telecom_users", [])
    print("抓取月份: %s  防溺水户数: %d" % (month, len(users)))
    if not users:
        print("config.json 的 fangnishui_users 为空，请先补充映射")
        sys.exit(1)

    rows = fetch_month(cfg, month, args.price)

    # 月初容错：本月数据尚未发布（T+3）时，自动回退抓上月
    if not rows or all(r.get("用电量(千瓦时)") in (None, "") for r in rows):
        prev = (datetime.date.today().replace(day=1) - datetime.timedelta(days=1)).strftime("%Y%m")
        print("本月 %s 暂无数据，回退抓取上月 %s" % (month, prev))
        month = prev
        rows = fetch_month(cfg, month, args.price)

    if not rows:
        print("无数据返回")
        return

    out_path = os.path.join(os.path.dirname(os.path.abspath(CONFIG_PATH)),
                            "南网每日电量_%s.xlsx" % month)
    fmt = to_excel(rows, out_path)
    print("已保存: %s (%s)" % (out_path, fmt))
    print("共 %d 条记录" % len(rows))

    # 飞书提醒：日电量 > 阈值 或 日估算电费 > 阈值
    if not args.no_alert:
        threshold = cfg.get("alert_threshold", {})
        fee_th = args.alert_fee if args.alert_fee is not None else threshold.get("daily_fee", DEFAULT_ALERT_FEE)
        power_th = args.alert_power if args.alert_power is not None else threshold.get("daily_power", DEFAULT_ALERT_POWER)
        print("告警阈值: 日电费 > %.1f 元 / 日电量 > %.1f 度" % (fee_th, power_th))
        alerts = build_alerts(rows, fee_th, power_th)
        if alerts:
            send_feishu_alert(cfg, alerts)
        else:
            print("最新日期无超标户号，不推送飞书")


if __name__ == "__main__":
    main()