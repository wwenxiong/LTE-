# 网优报表自动监控（SFTP 轮询 + 动态规则）

## 功能

- 读取本地 `Objects.xlsx`，获取需要监控的小区清单（列名：`小区名称`）
- 读取本地 `Rules.xlsx`，获取监控指标与阈值（列名：`监控指标`、`判断符`、`阈值`）
- 每 5 分钟轮询 SFTP 目录（默认 `/业务监控文件夹`），仅下载未处理过的 `.xlsx/.csv`
- 从报表中筛选匹配小区，并按规则逐条判定（指标不硬编码，完全由 `Rules.xlsx` 驱动）
- 若触发告警，输出到 `./output/alerts_<报表名>_<时间>.xlsx`
- 告警将通过钉钉机器人推送 Markdown 消息（需配置 webhook），通知中包含监控时段信息

## 依赖安装

```bash
pip install -r requirements.txt
```

## 表结构要求

### Objects.xlsx

- 必须包含列：`小区名称`

### Rules.xlsx

- 必须包含列：`监控指标`、`判断符`、`阈值`
- `判断符`支持：`>`、`<`、`>=`、`<=`、`==`、`!=`
- `监控指标`必须与报表里的**列名完全一致**（建议避免多余空格）

### 报表（新进的 .xlsx/.csv）

- 必须包含列：`小区名称`
- 必须包含 `Rules.xlsx` 中配置的各个 `监控指标` 列

## 运行

示例（按需替换 SFTP 参数）：

```bash
python main.py \
  --sftp-host your.host \
  --sftp-port 22 \
  --sftp-user your_user \
  --sftp-password your_password \
  --sftp-remote-path "/业务监控文件夹" \
  --interval 300
```

## 钉钉机器人通知

- 在 `config.json` 中设置 `dingtalk_webhook` 为钉钉机器人的 webhook 地址（或设置环境变量 `DINGTALK_WEBHOOK`）
- 告警会合并一个报表中所有异常小区，生成一条 Markdown 消息
- 通知内容包含：**监控时段**（从文件名自动解析）、报表文件名、处理时间、各小区违规指标详情
- 无异常时发送“✅ VIP区域指标监控正常”心跳消息，同样包含监控时段信息

## SFTP 轮询与增量识别

- 使用 `processed_files.txt` 记录已处理文件名，只对新文件触发下载与分析
- 轮询周期默认 300 秒，可用 `--interval` 调整
- 临时下载目录默认 `./temp_data/`

## 输出说明

- 如果报表中没有匹配到需要监控的小区，会提示并结束本次处理
- 如果匹配到小区但未触发任何规则，不输出告警文件
- 如果触发告警，会生成 `output/alerts_*.xlsx`，其中包含：
  - 原始报表字段（匹配小区的行）
  - 每条规则的判定列：`rule_<指标>_<判断符>_<阈值>`
  - 汇总列：`是否触发告警`

