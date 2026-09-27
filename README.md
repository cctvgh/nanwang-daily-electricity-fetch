# 南网在线日电量监控与提醒

> 月度清单要等下个月？重点电表等不起。调南网在线（95598.csg.cn）已验证接口抓取绑定户号的每日用电量（T+3 发布），估算日电费落盘 Excel；日电量或日估算电费超阈值，飞书自动提醒。

## 为什么需要它

管防溺水监控电表、基站、机房这类重点户的电费管理员，最怕"事后才知道"：等月度清单出来发现某池塘电表日电量异常，可能已经过去二十天。日级监控靠人工每天登网页逐户查，几十户点下来半小时就没了。

本工具把这件事变成一条命令 + 一个定时任务：

1. **T+3 日级数据**：抓取当月截至昨日的每日电量（南网官方发布节奏）
2. **估算日电费**：按参考电价折算（可覆盖），落盘 Excel（含日均温参照）
3. **阈值提醒**：日电量或日估算电费超阈值 → 飞书机器人推送红色卡片（含地址明细）
4. **户号批量管理**：手机号查询批量绑 / 用户编号+代扣卡后6位逐户绑，解绑也已跑通

## 功能亮点

| 能力 | 说明 |
|------|------|
| 每日抓取 | `python scripts/nw_daily_fetch.py`，支持指定月份、跳过提醒、覆盖电价与阈值 |
| Token 管理 | 自动检测失效（sta=04 弹窗提醒），配套 `nw_token_update.py` 续期 |
| 户号绑定/解绑 | 单账号上限 26 户，绑定双通道实测打通，`pointId` 可传空 |
| 飞书提醒 | 同日不重复推送；推送内容仅为业务电费文本，不含本地配置 |
| 零依赖 | Python 标准库 urllib，openpyxl 缺失时自动降级 CSV |
| 月初回退 | 本月数据未发布时自动回退抓取上月 |

## 快速开始

```bash
git clone https://github.com/cctvgh/nanwang-daily-electricity-fetch.git
cd nanwang-daily-electricity-fetch
# 编辑 config.json：填入你的 token / custNumber / 户号清单 / 飞书 webhook
python scripts/nw_daily_fetch.py            # 抓取本月（截至昨日）
python scripts/nw_daily_fetch.py 202609     # 指定月份
```

配置项说明、Token 获取方法、接口清单与认证机制见 `references/nanwang-api.md`（21KB 实测笔记）。

> 安全约定：Webhook 地址仅从你本人飞书群机器人设置中复制；Token 仅保存在本机 config.json，脚本不外传任何数据。

## 适用人群

- 防溺水监控电表责任人（教育系统 / 供电所 / 乡镇安全管理）
- 基站、机房电费管理员（通信运营商、代维公司）
- 物业、园区后勤多户号日级电费监控
- 南网在线的其他深度用户

## 作为 TeleAgent 技能使用

本工具同时发布为「星辰超级智能体（TeleAgent）」技能 **`nanwang-daily-electricity-fetch`（南网在线日电量监控与提醒）**，技能市场安装后直接说"抓每日电费""绑定户号""防溺水电表监控"即可。

配套技能：[`electricity-bill-analysis`](https://github.com/cctvgh/electricity-bill-analysis)（企业电费异常分析报告）负责月度深度分析。

## License

MIT © 2026 何汉锋
