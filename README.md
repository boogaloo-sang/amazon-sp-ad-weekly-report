# 广告经营与广告周月报

该仓库每周一美国东部时间上午 09:00 自动运行：

1. 通过领星 MCP 获取 主账号数据。
2. 使用 DeepSeek 选择只读工具、整理数据并生成异常摘要。
3. 按确定性规则生成 Excel 周月报。
4. 把摘要和 Excel 附件发送到飞书应用机器人私聊。

## 报表口径

- 周：周一至周日，展示最近四个完整自然周。
- 月：当月 1 日至最近完整周日，对比上月相同日数。
- 金额：全部为美元。
- TACOS：广告花费 ÷ 净销售额。
- 高花费产品：最近完整周广告花费大于 200 美元，按国家 × 父 ASIN 聚合。
- 负责人：由加密的 `OWNER_CONFIG_JSON` Secret 提供，不公开姓名和 UID。
- 产品对比：A 本周、B 上周、C 本月、D 上月同期，纵向排列。

## GitHub Actions Secrets

必须在 **Settings → Secrets and variables → Actions** 配置：

| Secret | 用途 |
| --- | --- |
| `DEEPSEEK_API_KEY` | DeepSeek API 密钥 |
| `LINGXING_MCP_URL` | 可从 GitHub Runner 访问的领星 MCP Streamable HTTP 地址 |
| `LINGXING_MCP_TOKEN` | 领星 MCP Bearer Token；不需要时可留空 |
| `LINGXING_MCP_HEADERS_JSON` | 其他认证请求头的 JSON；不需要时可留空 |
| `OWNER_CONFIG_JSON` | 负责人姓名与 UID 数组，例如 `[{"name":"负责人A","uid":"123"}]` |
| `FEISHU_APP_ID` | 飞书自建应用 App ID |
| `FEISHU_APP_SECRET` | 飞书自建应用 App Secret |
| `FEISHU_RECEIVE_ID` | 当前应用下接收人的 open_id |

`LINGXING_MCP_URL` 必须是公网或 GitHub Runner 可达地址。本机 Codex 中显示的 MCP 工具不会自动出现在 GitHub 云端。

## 手动测试

进入 **Actions → ARCCAPTAIN weekly business and ads report → Run workflow**。

- 勾选 `dry_run`：只生成并校验，不发送飞书。
- 不勾选：生成后发送摘要与 Excel 附件。

仓库不保存销售数据、Excel 报告或任何密钥。

