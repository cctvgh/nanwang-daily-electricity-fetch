# -*- coding: utf-8 -*-
"""
南网在线 Token 更新小工具
================================
当抓取脚本弹出"Token 失效"提醒时，运行本脚本获取新 token。

使用步骤：
1. 用浏览器（建议 Edge/Chrome）打开 https://95598.csg.cn 并登录
2. 按 F12 打开开发者工具 → Console（控制台）
3. 粘贴下面这行并回车：
   copy(sessionStorage.getItem('token') + '|' + sessionStorage.getItem('custnumber'))
4. 此时剪贴板里就是 "token|custnumber" 格式
5. 运行本脚本：python nw_token_update.py
   按提示粘贴刚才复制的内容，即可自动写入 config.json
"""
import argparse
import json
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")


def main():
    global CONFIG_PATH
    parser = argparse.ArgumentParser(description="南网在线 Token 更新工具")
    parser.add_argument("--config", default=CONFIG_PATH, help="config.json 路径（默认脚本同目录）")
    args = parser.parse_args()
    CONFIG_PATH = os.path.abspath(args.config)

    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    print("===== 南网在线 Token 更新工具 =====")
    print("步骤1：浏览器打开 https://95598.csg.cn 并登录（手机号+验证码）")
    print("步骤2：F12 → Console，粘贴并回车：")
    print("  copy(sessionStorage.getItem('token') + '|' + sessionStorage.getItem('custnumber'))")
    print("步骤3：粘贴复制的 token 内容：")
    line = input("> ").strip()
    if "|" not in line:
        print("格式错误：应为 token|custnumber")
        sys.exit(1)
    token, cust = line.split("|", 1)
    token = token.strip()
    cust = cust.strip()
    if not token or len(token) < 8:
        print("Token 无效：%r" % token)
        sys.exit(1)

    cfg["token"] = token
    cfg["custNumber"] = cust
    cfg["token_updated"] = datetime_today()
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    print("已更新 config.json：")
    print("  token      : %s" % token)
    print("  custNumber : %s" % cust)
    print("下次运行 nw_daily_fetch.py 即可生效。")


def datetime_today():
    import datetime
    return datetime.date.today().strftime("%Y-%m-%d")


if __name__ == "__main__":
    main()