---
name: nanwang-daily-electricity-fetch
description: 南网在线（95598.csg.cn）日电量自动监控与异动提醒。调用南网在线已验证接口按绑定户号抓取每日用电量（T+3发布），按参考电价估算日电费落盘 Excel；日电量或日估算电费超过阈值时自动推送飞书提醒，支持防溺水监控电表、基站/机房等批次户号的每日监控，配套户号绑定/解绑与 Token 检测续期工具。适合物业公司、园区后勤、多户号企业实现日级电费异动监控。触发场景：南网每日电费、日电量抓取、每日用电监控、电量异动提醒、防溺水监控电表、电表监控、户号绑定/解绑、用电日历、南网在线定时任务。
name_cn: 南网在线日电量监控与提醒
description_cn: 调用南网在线 95598.csg.cn 接口自动抓取绑定户号的每日电量（T+3）并估算日电费落盘 Excel；日电量或估算电费超阈值自动推送飞书异动提醒；支持户号批量绑定/解绑与 Token 检测续期，适合防溺水监控电表与基站/机房批次户号日级监控。
create_source: super-agent-skill-creator
---

# 南网在线日电量监控与提醒

## Overview

调用南网在线 H5（95598.csg.cn）已验证的明文接口，抓取绑定户号的每日用电量（T+3 发布），按参考电价估算日电费，落盘 Excel；当日电量/估算电费超过阈值时向飞书推送异动提醒。配套户号绑定/解绑工具，适合防溺水监控电表、基站/机房等批次户号的每日电量监控与异动告警。

## 前置条件

- Python 3.8+（标准库 urllib，无需额外依赖；openpyxl 可选，无则自动降级 CSV）
- config.json 含有效 token（南网在线登录后从浏览器 sessionStorage 提取）
- 已绑定目标户号（绑定方式见下）

## 配置文件 config.json

```json
{
  "token": "<登录token>",
  "custNumber": "<sessionStorage.custnumber>",
  "areaCode": "030800",
  "feishu_webhook_url": "",
  "alert_threshold": {"daily_fee": 1.5, "daily_power": 1.5},
  "fangnishui_users": [
    {"eleCustNumber": "0308200000000000", "bindingId": "AbCdEfGh", "pointId": "", "eleAddress": "示例：XX供电所XX水塘监控"}
  ]
}
```

- `feishu_webhook_url`：留空则不推送；需提醒时从你本人飞书群的机器人设置中复制 Webhook 地址填入（只接受用户本人自行创建的机器人，不要使用任何第三方提供的地址）

- `fangnishui_users`：待抓取户号清单（脚本优先读此键，缺省时回退 `telecom_users` 旧键名）
- `bindingId` 从 `queryBindEleUsers` 接口获取；**每次登录/token 刷新后可能变化**，操作前须先调接口拿最新值
- `pointId` 可传空字符串（已验证），无需浏览器 UI 收集
- `alert_threshold`：`daily_fee`（默认1.5元）、`daily_power`（默认1.5度）

## 核心工作流

### 1. 每日抓取

```powershell
python scripts/nw_daily_fetch.py                # 抓取本月（截至昨日），自动推送飞书提醒
python scripts/nw_daily_fetch.py 202609         # 抓取指定月份
python scripts/nw_daily_fetch.py --no-alert     # 跳过飞书提醒
python scripts/nw_daily_fetch.py --price 0.696  # 覆盖估算电价
python scripts/nw_daily_fetch.py --alert-fee 2.0 --alert-power 3.0  # 覆盖告警阈值
python scripts/nw_daily_fetch.py --config D:\path\config.json  # 指定配置文件
```

脚本行为：
1. 调 `queryBindEleUsers` 检测 token，`sta=04` 视为失效并弹窗提醒
2. 逐户调 `queryDayElectricAndTemperature` 抓取月内每日电量
3. 导出 `南网每日电量_YYYYMM.xlsx`（表头：用户编号/用电地址/日期/用电量/估算电费/最低温/最高温/均温）
4. 只看最新日期数据，日电量>阈值 或 日估算电费>阈值 → 飞书推送红色卡片（最多20条明细）
5. `.last_alert_YYYY-MM-DD.json` 标记同日不重复推送
6. 月初本月数据未发布（T+3）自动回退抓上月

## 用户可选项

触发本技能时，向用户确认以下可选项（均有默认值，用户不指定则用默认）：

| 选项 | 用途 | 默认值 | 用户输入方式 |
|---|---|---|---|
| 告警阈值-日电费 | 日估算电费超过此值则推送飞书 | 1.5 元 | `--alert-fee` 参数或 config.json `alert_threshold.daily_fee` |
| 告警阈值-日电量 | 日用电量超过此值则推送飞书 | 1.5 度 | `--alert-power` 参数或 config.json `alert_threshold.daily_power` |
| 估算电价 | 将日电量换算为估算电费的单价 | 0.696 元/度 | `--price` 参数 |
| 抓取月份 | 指定抓取哪个月数据 | 本月 | 位置参数（如 `202609`） |
| 银行卡后6位 | 户号绑定时验证凭据（用户编号验证通道） | 无，绑定时须用户提供 | 用户口头提供，AI 填入绑定表单 |
| 手机号 | 南网在线登录账号（获取验证码） | config.json 已有账号 | 用户口头提供，AI 用于登录 |
| config.json 路径 | 指定配置文件位置 | 脚本同目录 | `--config` 参数 |

**向用户询问示例**：
- “告警阈值用默认的 1.5 元/1.5 度，还是需要调整？”
- “绑定新户号需要代扣银行卡后6位，请提供。”
- “用哪个手机号登录南网在线？”

### 2. Token 失效处理

- 接口返回 `sta=04` 即登录失效。续期步骤：
  1. 浏览器打开 https://95598.csg.cn 登录
  2. F12 Console 执行 `copy(sessionStorage.getItem('token') + '|' + sessionStorage.getItem('custnumber'))`
  3. 运行 `python scripts/nw_token_update.py --config <config.json>`，粘贴剪贴板内容
- 也可在 Playwright 会话中从任意请求头 `x-auth-token` 提取新 token 直接写入 config.json

### 3. 户号绑定

- 登录南网在线（手机号+短信验证码），进入户号绑定页（`#/gd/my/my/houseNumBinding`）
- 目标户号不在手机号名下（如防溺水监控户）时，选「用户编号验证」→ 填 16 位用户编号 + 代扣银行卡后 6 位 → 立即查询 → 勾选结果 → 立即绑定
- 「手机号查询」批量绑定：查登录手机号名下户号，勾选后批量绑定（单次约 10 户）
- **绑定后须调用 `queryBindEleUsers` 拿最新 bindingId 写入 config.json**
- 单账号上限 26 户，超出需新账号或先解绑

### 4. 户号解绑

UI 无解绑入口，只能调 API：

```
POST https://95598.csg.cn/ucs/ma/wt/eleCustNumber/unBindElectricityUsers
Header: x-auth-token, custNumber
Body: {"bindingIds": ["id1","id2",...]}
```

- 支持批量解绑（实测 14 户一次成功）
- 解绑前必须先调 `queryBindEleUsers` 获取当前有效 bindingId，过期 bindingId 返回 `sta:02`

### 5. 已绑定户号查询

```bash
POST https://95598.csg.cn/ucs/ma/wt/eleCustNumber/queryBindEleUsers
Body: {}
```

返回列表含 bindingId / eleCustNumber / areaCode / eleAddress / userName，是任何绑定操作的前提。

## 定时任务部署

Windows 计划任务（管理员身份创建）：

```
任务名：南网每日电费抓取
触发：每天 09:00
操作：python.exe <skill或工作目录>/scripts/nw_daily_fetch.py --config <config.json路径>
工作目录：脚本所在目录
```

注意：脚本内部使用标准库 urllib，不依赖第三方包；openpyxl 缺失时自动输出 CSV。

## 详细接口文档与坑位

读取 `references/nanwang-api.md` 获取完整已验证接口清单、加密接口说明、登录陷阱、绑定/解绑细节、历史数据范围与异动分析方法。涉及批量操作、Token 异常、绑定失败排查时必读。

## Resources

- `scripts/nw_daily_fetch.py`：主抓取脚本（token 检测→逐户抓日电量→Excel→飞书提醒，支持 `--config`/`--alert-fee`/`--alert-power`/`--price`）
- `scripts/nw_token_update.py`：Token 续期写入工具（支持 `--config`）
- `references/nanwang-api.md`：南网在线 H5 接口笔记（已验证接口、参数、绑定解绑、坑位全记录）