"""本地测试切换数据处理逻辑（不发送钉钉通知）"""
import sys
sys.stdout.reconfigure(encoding='utf-8')

import logging
import json
from pathlib import Path

# 配置日志
logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")

# 模拟 send_dingtalk_notification，只打印不发送
import main
_original_send = main.send_dingtalk_notification
def mock_send(title, content, **kwargs):
    print(f"\n{'='*60}")
    print(f"[钉钉消息标题] {title}")
    print(f"{'='*60}")
    print(content)
    print(f"{'='*60}\n")
main.send_dingtalk_notification = mock_send

# 加载配置
with open("config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

# 测试参数
test_file = Path("性能管理-历史查询-点对点切换-ZTE_wuwenxiong-20260421094843.xlsx")
objects_path = Path("Objects.xlsx")
output_dir = Path("./output")

print("=== 测试1: 处理切换数据（使用 Objects.xlsx 过滤） ===")
main.process_handover_file(test_file, objects_path, output_dir, config)

# 恢复
main.send_dingtalk_notification = _original_send
