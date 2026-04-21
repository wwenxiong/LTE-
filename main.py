import argparse
import json
import logging
import os
import threading
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable, Dict, Optional, Tuple, List, Set, Any
import ctypes
import posixpath
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import queue

import pandas as pd
import requests
import paramiko
import pystray
from pystray import MenuItem as MenuItem
from PIL import Image, ImageDraw
from win10toast import ToastNotifier


# =========================
# 配置（如需可改为命令行参数）
# =========================
DEFAULT_OBJECTS_PATH = Path("./Objects.xlsx")
DEFAULT_RULES_PATH = Path("./Rules.xlsx")
DEFAULT_OUTPUT_DIR = Path("./output")
DEFAULT_TEMP_DIR = Path("./temp_data")
DEFAULT_PROCESSED_FILE = Path("./processed_files.txt")
CONFIG_PATH = Path("./config.json")

# SFTP 默认配置
DEFAULT_SFTP_HOST = ""
DEFAULT_SFTP_PORT = 22
DEFAULT_SFTP_USER = ""
DEFAULT_SFTP_PASSWORD = ""
DEFAULT_SFTP_REMOTE_PATH = "/业务监控文件夹"

# 是否在无异常时也推送“心跳”消息
SEND_NORMAL_HEARTBEAT = True

CELL_COL = "小区名称"  # Objects.xlsx 与报表中匹配小区的列名

# 运行状态（供托盘菜单查询/控制）
last_scan_time: Optional[datetime] = None
scan_stop_event = threading.Event()
scan_force_event = threading.Event()


def setup_logger(level: str = "INFO") -> None:
    logging.basicConfig(
        level=getattr(logging, level.upper(), logging.INFO),
        format="%(asctime)s | %(levelname)s | %(message)s",
    )


def validate_config(data: Dict[str, Any]) -> Dict[str, Any]:
    """
    校验并规范化 config.json 内容。
    """
    required_keys = [
        "sftp_host",
        "sftp_port",
        "sftp_user",
        "sftp_password",
        "sftp_remote_path",
        "sleep_minutes",
        "dingtalk_webhook",
    ]
    # 兼容旧配置：wechat_webhook -> dingtalk_webhook
    if "wechat_webhook" in data and "dingtalk_webhook" not in data:
        data["dingtalk_webhook"] = data.pop("wechat_webhook")
    for key in required_keys:
        if key not in data:
            raise ValueError(f"配置缺少字段：{key}")

    cfg: Dict[str, Any] = {}
    cfg["sftp_host"] = str(data["sftp_host"]).strip()
    cfg["sftp_port"] = int(data["sftp_port"])
    cfg["sftp_user"] = str(data["sftp_user"]).strip()
    cfg["sftp_password"] = str(data["sftp_password"])
    cfg["sftp_remote_path"] = str(data["sftp_remote_path"]).strip() or DEFAULT_SFTP_REMOTE_PATH
    cfg["sleep_minutes"] = float(data["sleep_minutes"])
    if cfg["sleep_minutes"] <= 0:
        raise ValueError("休眠时间必须为正数")
    
    # 调度模式校验
    cfg["schedule_mode"] = str(data.get("schedule_mode", "interval"))
    cfg["schedule_minute"] = int(data.get("schedule_minute", 0))
    cfg["schedule_fixed_times"] = str(data.get("schedule_fixed_times", "")).strip()
    
    if cfg["schedule_mode"] == "hourly":
        if not (0 <= cfg["schedule_minute"] <= 59):
            raise ValueError("整点模式下，分钟数必须在 0-59 之间")
            
    if cfg["schedule_mode"] == "fixed_times":
        times = cfg["schedule_fixed_times"].replace("，", ",").split(",")
        valid_count = 0
        for t in times:
            t = t.strip()
            if not t: continue
            try:
                datetime.strptime(t, "%H:%M")
                valid_count += 1
            except ValueError:
                raise ValueError(f"时间格式错误：'{t}'，请使用 HH:MM 格式（如 08:00）")
        if valid_count == 0:
            raise ValueError("固定时间点模式下，必须至少指定一个有效时间")

    cfg["retention_days"] = int(data.get("retention_days", 7))
    if cfg["retention_days"] < 0:
        raise ValueError("文件保留天数不能为负数")

    cfg["dingtalk_webhook"] = str(data["dingtalk_webhook"]).strip()
    cfg["dingtalk_secret"] = str(data.get("dingtalk_secret", "")).strip()
    cfg["send_normal_heartbeat"] = bool(data.get("send_normal_heartbeat", True))

    # 切换监控配置
    cfg["handover_enabled"] = bool(data.get("handover_enabled", True))
    cfg["handover_success_rate_threshold"] = float(data.get("handover_success_rate_threshold", 0.95))
    cfg["handover_min_attempts"] = int(data.get("handover_min_attempts", 5))
    if not (0 < cfg["handover_success_rate_threshold"] <= 1):
        raise ValueError("切换成功率阈值必须在 0~1 之间（如 0.95 表示 95%）")
    if cfg["handover_min_attempts"] < 0:
        raise ValueError("切换最低请求次数不能为负数")

    return cfg


def get_default_config() -> Dict[str, Any]:
    return {
        "sftp_host": DEFAULT_SFTP_HOST or "127.0.0.1",
        "sftp_port": DEFAULT_SFTP_PORT,
        "sftp_user": DEFAULT_SFTP_USER or "user",
        "sftp_password": DEFAULT_SFTP_PASSWORD or "",
        "sftp_remote_path": DEFAULT_SFTP_REMOTE_PATH,
        "sleep_minutes": 5.0,
        "dingtalk_webhook": "",
        "dingtalk_secret": "",
        "send_normal_heartbeat": True,
        "schedule_mode": "interval",  # interval 或 hourly
        "schedule_minute": 0,
        "schedule_fixed_times": "08:00,12:00,18:00",
        "retention_days": 7,
        "handover_enabled": True,
        "handover_success_rate_threshold": 0.95,
        "handover_min_attempts": 5,
    }


@dataclass(frozen=True)
class Rule:
    metric: str
    op: str
    threshold_raw: object


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_objects(objects_path: Path) -> pd.DataFrame:
    """
    从 Objects.xlsx 读取小区与区域清单。
    要求列名包含：小区名称。如果有“区域”列，则保留，否则用“未知区域”填充。
    返回去重的 DataFrame，包含：['小区名称', '区域']
    """
    df = pd.read_excel(objects_path)
    df = _normalize_columns(df)
    if CELL_COL not in df.columns:
        raise ValueError(f"Objects 文件缺少列：{CELL_COL}（实际列：{list(df.columns)}）")
    
    if "区域" in df.columns:
        cells_df = df[[CELL_COL, "区域"]].copy()
    else:
        cells_df = df[[CELL_COL]].copy()
        cells_df["区域"] = "未知区域"
        
    cells_df[CELL_COL] = cells_df[CELL_COL].astype(str).str.strip()
    cells_df["区域"] = cells_df["区域"].astype(str).str.strip().replace({"nan": "未知区域", "None": "未知区域", "": "未知区域"})
    
    cells_df = cells_df.replace({"nan": None, "None": None, "": None}).dropna(subset=[CELL_COL])
    cells_df = cells_df.drop_duplicates(subset=[CELL_COL])
    return cells_df


def load_rules(rules_path: Path) -> Tuple[pd.DataFrame, Dict[str, Callable[[pd.Series, pd.Series], pd.Series]]]:
    """
    从 Rules.xlsx 读取规则表。
    规则列名要求包含：监控指标、判断符、阈值
    返回：
      - rules_df：清洗后的 DataFrame
      - op_map：支持的运算符映射
    """
    rules_df = pd.read_excel(rules_path)
    rules_df = _normalize_columns(rules_df)

    required_cols = ["监控指标", "判断符", "阈值"]
    missing = [c for c in required_cols if c not in rules_df.columns]
    if missing:
        raise ValueError(f"Rules 文件缺少列：{missing}（实际列：{list(rules_df.columns)}）")

    # 清洗
    rules_df = rules_df[required_cols].copy()
    rules_df["监控指标"] = rules_df["监控指标"].astype(str).str.strip()
    rules_df["判断符"] = rules_df["判断符"].astype(str).str.strip()
    # 阈值可能是数字/字符串/空；先原样保留，后面按指标列动态转数值

    # 运算符映射（完全动态读取 Rules.xlsx 的判断符字段）
    op_map: Dict[str, Callable[[pd.Series, pd.Series], pd.Series]] = {
        ">": lambda s, t: s > t,
        "<": lambda s, t: s < t,
        ">=": lambda s, t: s >= t,
        "<=": lambda s, t: s <= t,
        "==": lambda s, t: s == t,
        "!=": lambda s, t: s != t,
    }
    return rules_df, op_map


def read_report(path: Path) -> pd.DataFrame:
    """
    读取新进报表（仅 .xlsx/.xls）。
    要求报表包含列：小区名称，以及 Rules.xlsx 中配置的指标列名。
    """
    suffix = path.suffix.lower()
    if suffix in [".xlsx", ".xls"]:
        df = pd.read_excel(path)
    else:
        raise ValueError(f"不支持的报表类型：{path.name}")
    return _normalize_columns(df)


def wait_for_file_ready(path: Path, timeout_s: int = 60, stable_checks: int = 3, interval_s: float = 1.0) -> None:
    """
    Windows 下常见：文件刚落地仍在写入/锁定。这里通过“文件大小稳定”判断可读。
    """
    start = time.time()
    last_size: Optional[int] = None
    stable = 0

    while True:
        if time.time() - start > timeout_s:
            raise TimeoutError(f"等待文件可用超时：{path}")

        if not path.exists():
            stable = 0
            time.sleep(interval_s)
            continue

        try:
            size = path.stat().st_size
        except OSError:
            size = None

        if size is None:
            stable = 0
        else:
            if last_size is not None and size == last_size:
                stable += 1
            else:
                stable = 0
            last_size = size

        if stable >= stable_checks:
            # 再尝试打开一次，防止锁
            try:
                with open(path, "rb"):
                    return
            except OSError:
                stable = 0

        time.sleep(interval_s)


def evaluate_rules(
    report_df: pd.DataFrame,
    cells_df: pd.DataFrame,
    rules_df: pd.DataFrame,
    op_map: Dict[str, Callable[[pd.Series, pd.Series], pd.Series]],
) -> Tuple[pd.DataFrame, List[Dict[str, object]]]:
    """
    动态匹配小区 + 动态规则判定。
    输出：命中的行 + 每条规则的判定结果 + 违规汇总列。
    """
    if CELL_COL not in report_df.columns:
        raise ValueError(f"报表缺少列：{CELL_COL}（实际列：{list(report_df.columns)}）")

    df = report_df.copy()
    df[CELL_COL] = df[CELL_COL].astype(str).str.strip()

    # 只保留需要监控的小区，并加入区域信息
    df = df.merge(cells_df, on=CELL_COL, how="inner")
    if df.empty:
        return df, []

    # 对每条规则生成一列：rule_<指标>_<判断符>_<阈值>
    violated_any = pd.Series(False, index=df.index)
    rule_meta: List[Dict[str, object]] = []

    for i, row in rules_df.iterrows():
        metric = str(row["监控指标"]).strip()
        op = str(row["判断符"]).strip()
        thr_raw = row["阈值"]

        if not metric or metric.lower() == "nan":
            logging.warning("规则第 %s 行监控指标为空，已跳过", i + 2)  # +2 近似考虑表头
            continue
        if op not in op_map:
            logging.warning("规则第 %s 行判断符不支持：%r，已跳过", i + 2, op)
            continue
        if metric not in df.columns:
            logging.warning("报表缺少指标列：%r（规则第 %s 行），已跳过", metric, i + 2)
            continue

        s = pd.to_numeric(df[metric], errors="coerce")
        t = pd.to_numeric(pd.Series([thr_raw] * len(df), index=df.index), errors="coerce")

        col_name = f"rule_{metric}_{op}_{thr_raw}"
        result = op_map[op](s, t)
        # NaN 比较结果为 False；为便于排查可保留 NaN 情况
        result = result.fillna(False)
        df[col_name] = result
        violated_any = violated_any | result
        rule_meta.append({"col": col_name, "metric": metric, "op": op, "threshold": thr_raw})

    df["是否触发告警"] = violated_any
    return df, rule_meta


def save_alerts(df: pd.DataFrame, src_file: Path, output_dir: Path) -> Optional[Path]:
    """
    将触发告警的记录输出到 output 目录。
    """
    if df.empty:
        return None
    alerted = df[df.get("是否触发告警", False) == True].copy()  # noqa: E712
    if alerted.empty:
        return None

    output_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = output_dir / f"alerts_{src_file.stem}_{ts}.xlsx"
    alerted.to_excel(out_path, index=False)
    return out_path


def parse_file_time_range(filename: str) -> str:
    """
    从报表文件名中解析监控时间段。
    支持两种文件名格式：
      - 新格式：两江VIP指标监控_202604090800_202604090900.xlsx -> 2026-04-09 08:00~09:00
      - 旧格式：两江VIP指标监控_2026012716001700.xlsx -> 2026-01-27 16:00~17:00
    解析失败时返回"未知时段"。
    """
    import re
    stem = Path(filename).stem  # 去掉扩展名

    # 新格式：xxx_YYYYMMDDHHMM_YYYYMMDDHHMM
    m = re.search(r'(\d{12})_(\d{12})$', stem)
    if m:
        start_str, end_str = m.group(1), m.group(2)
        try:
            start_dt = datetime.strptime(start_str, "%Y%m%d%H%M")
            end_dt = datetime.strptime(end_str, "%Y%m%d%H%M")
            if start_dt.date() == end_dt.date():
                return f"{start_dt.strftime('%Y-%m-%d')} {start_dt.strftime('%H:%M')}~{end_dt.strftime('%H:%M')}"
            else:
                return f"{start_dt.strftime('%Y-%m-%d %H:%M')} ~ {end_dt.strftime('%Y-%m-%d %H:%M')}"
        except ValueError:
            pass

    # 旧格式：xxx_YYYYMMDDHHMM HHMM（连续14位数字）
    m = re.search(r'(\d{10})(\d{4})$', stem)
    if m:
        date_hour_str, end_hour_str = m.group(1), m.group(2)
        try:
            start_dt = datetime.strptime(date_hour_str, "%Y%m%d%H%M")
            end_hour = int(end_hour_str[:2])
            end_min = int(end_hour_str[2:])
            end_dt = start_dt.replace(hour=end_hour, minute=end_min)
            if end_dt < start_dt:
                end_dt += timedelta(days=1)
            if start_dt.date() == end_dt.date():
                return f"{start_dt.strftime('%Y-%m-%d')} {start_dt.strftime('%H:%M')}~{end_dt.strftime('%H:%M')}"
            else:
                return f"{start_dt.strftime('%Y-%m-%d %H:%M')} ~ {end_dt.strftime('%Y-%m-%d %H:%M')}"
        except (ValueError, IndexError):
            pass

    return "未知时段"


def build_markdown_content(alert_df: pd.DataFrame, rule_meta: List[Dict[str, object]], filename: str, region_name: str = "VIP区域") -> str:
    """
    将告警行格式化为钉钉 Markdown 消息。
    """
    time_range = parse_file_time_range(filename)
    ts_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines: List[str] = [
        f"#### ⚠️ {region_name}指标异常告警",
        "",
        f"> **监控区域**：{region_name}",
        f"> **监控时段**：{time_range}",
        f"> **报表文件**：{filename}",
        f"> **处理时间**：{ts_str}",
        "",
    ]
    alert_count = 0
    for _, row in alert_df.iterrows():
        cell_name = row.get(CELL_COL, "未知小区")
        lines.append(f"**小区：{cell_name}**")
        for meta in rule_meta:
            col = meta["col"]
            if col not in alert_df.columns:
                continue
            if bool(row.get(col, False)) is False:
                continue
            metric = meta["metric"]
            op = meta["op"]
            thr_raw = meta["threshold"]
            cur_val = row.get(metric, "N/A")
            lines.append(f"- {metric} {op} {thr_raw}，当前值：**{cur_val}**")
            alert_count += 1
        lines.append("")  # 分隔不同小区
    lines.append("---")
    lines.append(f"共 **{len(alert_df)}** 个小区触发 **{alert_count}** 条告警")
    return "\n\n".join(lines).strip()


def build_normal_markdown(filename: str, ts: Optional[datetime] = None, region_name: str = "VIP区域") -> str:
    """
    无异常时的钉钉 Markdown 文本。
    """
    if ts is None:
        ts = datetime.now()
    ts_str = ts.strftime("%Y-%m-%d %H:%M:%S")
    time_range = parse_file_time_range(filename)
    lines: List[str] = [
        f"#### ✅ {region_name}指标监控正常",
        "",
        f"> **监控区域**：{region_name}",
        f"> **监控时段**：{time_range}",
        f"> **报表文件**：{filename}",
        f"> **处理时间**：{ts_str}",
        "",
        f"当前 {region_name} 所有指标未发现异常。",
    ]
    return "\n\n".join(lines)


# =========================
# 邻区对切换监控
# =========================

# 切换成功率相关列配置（指标名, 请求次数列, 成功次数列, 成功率列, 显示名）
HANDOVER_METRICS = [
    ("FDD切换出请求总次数(小区对)-异频_重庆", "FDD切换出成功总次数(小区对)-异频_重庆", "FDD切换出成功率-异频_重庆", "异频切换出"),
    ("FDD切换入请求总次数(小区对)-异频_重庆", "FDD切换入成功总次数(小区对)-异频_重庆", "FDD切换入成功率-异频_重庆", "异频切换入"),
    ("FDD切换出请求总次数(小区对)-同频_重庆", "FDD切换出成功总次数(小区对)-同频_重庆", "FDD切换出成功率-同频_重庆", "同频切换出"),
    ("FDD切换入请求总次数(小区对)-同频_重庆", "FDD切换入成功总次数(小区对)-同频_重庆", "FDD切换入成功率-同频_重庆", "同频切换入"),
]


def is_handover_file(filename: str) -> bool:
    """判断文件是否为邻区对切换类型报表。"""
    return "点对点切换" in filename


def parse_handover_time_range(df: pd.DataFrame) -> str:
    """
    从切换数据的'开始时间'/'结束时间'列解析监控时段。
    """
    try:
        start = pd.to_datetime(df["开始时间"]).min()
        end = pd.to_datetime(df["结束时间"]).max()
        if start.date() == end.date():
            return f"{start.strftime('%Y-%m-%d')} {start.strftime('%H:%M')}~{end.strftime('%H:%M')}"
        else:
            return f"{start.strftime('%Y-%m-%d %H:%M')} ~ {end.strftime('%Y-%m-%d %H:%M')}"
    except Exception:
        return "未知时段"


def process_handover_file(
    path: Path,
    objects_path: Path,
    output_dir: Path,
    config: Dict[str, Any],
) -> None:
    """
    处理邻区对切换数据文件：
    1. 读取切换数据
    2. 按 Objects.xlsx 过滤监控小区
    3. 识别异常邻区对（成功率 < 阈值 且 请求次数 >= 最小阈值）
    4. 推送钉钉告警
    """
    logging.info("开始处理切换数据文件：%s", path)
    wait_for_file_ready(path)

    # 读取数据
    ho_df = pd.read_excel(path)
    ho_df = _normalize_columns(ho_df)

    if CELL_COL not in ho_df.columns:
        logging.warning("切换数据缺少列：%s，跳过", CELL_COL)
        return

    ho_df[CELL_COL] = ho_df[CELL_COL].astype(str).str.strip()

    # 使用 Objects.xlsx 过滤小区
    cells_df = load_objects(objects_path)
    filtered = ho_df.merge(cells_df, on=CELL_COL, how="inner")

    # 加载 sheet2 的 eci 映射
    eci_mapping = {}
    try:
        sheet2_df = pd.read_excel(objects_path, sheet_name=1)
        sheet2_df = _normalize_columns(sheet2_df)
        eci_col = next((c for c in sheet2_df.columns if c.lower() == 'eci'), None)
        name_col = next((c for c in sheet2_df.columns if '小区' in c), None)
        if eci_col and name_col:
            sheet2_df = sheet2_df.dropna(subset=[eci_col])
            sheet2_df[eci_col] = pd.to_numeric(sheet2_df[eci_col], errors="coerce")
            sheet2_df = sheet2_df.dropna(subset=[eci_col])
            for _, r in sheet2_df.iterrows():
                eci_mapping[int(r[eci_col])] = str(r[name_col]).strip()
    except Exception as exc:
        logging.warning("读取 Objects.xlsx sheet2 失败（无法加载 eci 映射）：%s", exc)

    if filtered.empty:
        logging.info("切换数据中未匹配到需要监控的小区：%s", path.name)
        if config.get("send_normal_heartbeat", True):
            time_range = parse_handover_time_range(ho_df)
            ts_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            content = f"#### ✅ 邻区对切换监控正常\n\n> **监控时段**：{time_range}\n> **报表文件**：{path.name}\n> **处理时间**：{ts_str}\n\n> **说明**：该报表中未发现 `Objects.xlsx` 中定义的小区。"
            try:
                send_dingtalk_notification("✅ 切换监控正常(无匹配)", content)
            except Exception as exc:
                logging.warning("发送切换无匹配心跳时异常：%s", exc)
        return

    threshold = float(config.get("handover_success_rate_threshold", 0.95))
    min_attempts = int(config.get("handover_min_attempts", 5))
    time_range = parse_handover_time_range(filtered)

    # 收集异常记录
    abnormal_records: List[Dict[str, Any]] = []

    for req_col, succ_col, rate_col, display_name in HANDOVER_METRICS:
        if rate_col not in filtered.columns or req_col not in filtered.columns:
            logging.warning("切换数据缺少列：%s 或 %s，跳过该指标", rate_col, req_col)
            continue

        for _, row in filtered.iterrows():
            req_count = pd.to_numeric(row.get(req_col, 0), errors="coerce")
            succ_count = pd.to_numeric(row.get(succ_col, 0), errors="coerce")
            rate = pd.to_numeric(row.get(rate_col, 1.0), errors="coerce")

            if pd.isna(req_count) or pd.isna(rate):
                continue
            req_count = int(req_count)
            if pd.isna(succ_count):
                succ_count = 0
            else:
                succ_count = int(succ_count)

            if req_count >= min_attempts and rate < threshold:
                access_id = pd.to_numeric(row.get("AccessId", None), errors="coerce")
                peer_cell_id = pd.to_numeric(row.get("PeerCellId", None), errors="coerce")
                
                neighbor_name = str(row.get("LTE邻接关系ID", "未知"))
                
                if pd.notna(access_id) and pd.notna(peer_cell_id):
                    calculated_eci = int(access_id) * 256 + int(peer_cell_id)
                    if calculated_eci in eci_mapping:
                        neighbor_name = eci_mapping[calculated_eci]

                abnormal_records.append({
                    "小区名称": row.get(CELL_COL, "未知"),
                    "区域": row.get("区域", "未知区域"),
                    "LTE邻接关系ID": str(row.get("LTE邻接关系ID", "未知")),
                    "邻区名称": neighbor_name,
                    "切换类型": display_name,
                    "请求次数": req_count,
                    "成功次数": succ_count,
                    "成功率": rate,
                    "成功率百分比": f"{rate * 100:.2f}%",
                })

    if abnormal_records:
        abnormal_df = pd.DataFrame(abnormal_records)
        logging.warning("切换数据发现 %d 条异常邻区对：%s", len(abnormal_df), path.name)

        # 保存异常报告
        out_path = save_handover_alerts(abnormal_df, path, output_dir)
        if out_path:
            logging.info("切换异常报告已保存：%s", out_path)

        # 构建并推送钉钉告警
        content = build_handover_markdown(abnormal_df, path.name, time_range, threshold, min_attempts)
        try:
            send_dingtalk_notification("⚠️ 邻区对切换异常", content)
        except Exception as exc:
            logging.warning("发送切换异常告警失败：%s", exc)
    else:
        logging.info("切换数据未发现异常邻区对：%s", path.name)
        if config.get("send_normal_heartbeat", True):
            content = build_handover_normal_markdown(path.name, time_range, threshold)
            try:
                send_dingtalk_notification("✅ 邻区对切换正常", content)
            except Exception as exc:
                logging.warning("发送切换正常心跳失败：%s", exc)


def save_handover_alerts(abnormal_df: pd.DataFrame, src_file: Path, output_dir: Path) -> Optional[Path]:
    """将切换异常邻区对输出到 output 目录。"""
    if abnormal_df.empty:
        return None
    output_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = output_dir / f"handover_alerts_{src_file.stem}_{ts}.xlsx"
    abnormal_df.to_excel(out_path, index=False)
    return out_path


def build_handover_markdown(
    abnormal_df: pd.DataFrame,
    filename: str,
    time_range: str,
    threshold: float,
    min_attempts: int,
) -> str:
    """将切换异常邻区对格式化为钉钉 Markdown 消息。"""
    ts_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines: List[str] = [
        "#### ⚠️ 邻区对切换异常告警",
        "",
        f"> **监控时段**：{time_range}",
        f"> **报表文件**：{filename}",
        f"> **处理时间**：{ts_str}",
        f"> **判定条件**：成功率 < {threshold*100:.0f}% 且 请求次数 ≥ {min_attempts}",
        "",
    ]

    # 按小区分组
    grouped = abnormal_df.groupby("小区名称")
    cell_count = 0
    alert_count = 0
    for cell_name, group in grouped:
        cell_count += 1
        lines.append(f"**小区：{cell_name}**")
        for _, row in group.iterrows():
            ho_type = row["切换类型"]
            neighbor_name = row.get("邻区名称", row.get("LTE邻接关系ID", "未知"))
            req = row["请求次数"]
            succ = row["成功次数"]
            rate_pct = row["成功率百分比"]
            lines.append(f"- {ho_type} → 邻区={neighbor_name} | 请求{req}次 成功{succ}次 | 成功率：**{rate_pct}**")
            alert_count += 1
        lines.append("")  # 小区间分隔

    lines.append("---")
    lines.append(f"共 **{cell_count}** 个小区触发 **{alert_count}** 条切换异常")
    return "\n\n".join(lines).strip()


def build_handover_normal_markdown(filename: str, time_range: str, threshold: float) -> str:
    """切换无异常时的钉钉 Markdown 文本。"""
    ts_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines: List[str] = [
        "#### ✅ 邻区对切换监控正常",
        "",
        f"> **监控时段**：{time_range}",
        f"> **报表文件**：{filename}",
        f"> **处理时间**：{ts_str}",
        "",
        f"当前所有监控小区的邻区对切换成功率均正常（≥ {threshold*100:.0f}%）。",
    ]
    return "\n\n".join(lines)


def _dingtalk_sign(secret: str) -> Tuple[str, str]:
    """
    钉钉加签：根据密钥计算 HMAC-SHA256 签名。
    返回 (timestamp_ms, url_encoded_sign)。
    """
    import hmac
    import hashlib
    import base64
    import urllib.parse

    timestamp = str(int(round(time.time() * 1000)))
    string_to_sign = f"{timestamp}\n{secret}"
    hmac_code = hmac.new(
        secret.encode("utf-8"),
        string_to_sign.encode("utf-8"),
        digestmod=hashlib.sha256,
    ).digest()
    sign = urllib.parse.quote_plus(base64.b64encode(hmac_code).decode("utf-8"))
    return timestamp, sign


def send_dingtalk_notification(title: str, content: str, webhook_url: Optional[str] = None, secret: Optional[str] = None) -> None:
    """
    发送钉钉机器人 Markdown 消息。失败时仅记录日志，不抛出异常。
    webhook_url 若未传入，则从环境变量 DINGTALK_WEBHOOK 读取。
    secret 若未传入，则从环境变量 DINGTALK_SECRET 读取（加签安全设置）。
    """
    webhook = webhook_url or os.environ.get("DINGTALK_WEBHOOK")
    if not webhook:
        logging.warning("未配置钉钉 webhook（环境变量 DINGTALK_WEBHOOK），已跳过发送。")
        return

    # 加签处理
    sign_secret = secret or os.environ.get("DINGTALK_SECRET", "")
    if sign_secret:
        timestamp, sign = _dingtalk_sign(sign_secret)
        sep = "&" if "?" in webhook else "?"
        webhook = f"{webhook}{sep}timestamp={timestamp}&sign={sign}"

    payload = {
        "msgtype": "markdown",
        "markdown": {
            "title": title,
            "text": content,
        },
    }
    try:
        resp = requests.post(webhook, json=payload, timeout=10)
        if resp.status_code != 200:
            logging.warning("钉钉发送返回非 200 状态：%s，响应：%s", resp.status_code, resp.text)
        else:
            resp_json = resp.json()
            if resp_json.get("errcode", 0) != 0:
                logging.warning("钉钉返回错误：%s", resp_json.get("errmsg", "未知错误"))
            else:
                logging.info("钉钉通知发送成功")
    except requests.RequestException as exc:
        logging.warning("钉钉发送失败：%s", exc)


def process_file(path: Path, objects_path: Path, rules_path: Path, output_dir: Path, config: Dict[str, Any]) -> None:
    logging.info("开始处理文件：%s", path)
    if path.suffix.lower() not in [".xlsx", ".xls"]:
        return

    wait_for_file_ready(path)
    cells = load_objects(objects_path)
    rules_df, op_map = load_rules(rules_path)
    report_df = read_report(path)

    evaluated, rule_meta = evaluate_rules(report_df, cells, rules_df, op_map)
    out = save_alerts(evaluated, path, output_dir)
    alert_df = evaluated[evaluated.get("是否触发告警", False) == True].copy()  # noqa: E712

    if evaluated.empty:
        logging.info("报表中未匹配到需要监控的小区：%s", path.name)
        if config.get("send_normal_heartbeat", True):
            content = build_normal_markdown(path.name, region_name="全部区域")
            # 特殊说明一下没有匹配到小区
            content += "\n\n> **说明**：该报表中未发现 `Objects.xlsx` 中定义的小区。"
            try:
                send_dingtalk_notification("✅ 监控正常(无匹配)", content)
            except Exception as exc:
                logging.warning("发送无匹配心跳时出现异常：%s", exc)
        return

    # 按区域分组处理并推送
    groups = evaluated.groupby("区域")
    
    for region, region_df in groups:
        region_alert_df = region_df[region_df.get("是否触发告警", False) == True].copy()
        
        region_name_display = str(region)
        if region_name_display == "未知区域":
            region_name_display = "VIP区域"
            
        if region_alert_df.empty:
            logging.info("区域 [%s] 未触发告警：%s", region, path.name)
            if config.get("send_normal_heartbeat", True):
                content = build_normal_markdown(path.name, region_name=region_name_display)
                try:
                    send_dingtalk_notification(f"✅ {region_name_display}监控正常", content)
                except Exception as exc:
                    logging.warning("按区域发送钉钉正常心跳时出现异常：%s", exc)
        else:
            logging.warning("区域 [%s] 触发告警：%s", region, path.name)
            content = build_markdown_content(region_alert_df, rule_meta, path.name, region_name_display)
            try:
                send_dingtalk_notification(f"⚠️ {region_name_display}指标异常", content)
            except Exception as exc:
                logging.warning("按区域发送钉钉告警时出现异常：%s", exc)

    if not alert_df.empty:
        logging.warning("已输出包含所有区域的告警文件：%s", out)


def load_processed_set(path: Path) -> Set[str]:
    if not path.exists():
        return set()
    with open(path, "r", encoding="utf-8") as f:
        return set(line.strip() for line in f if line.strip())


def append_processed(path: Path, filename: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "a", encoding="utf-8") as f:
        f.write(filename + "\n")


def connect_sftp(host: str, port: int, username: str, password: str) -> paramiko.SFTPClient:
    transport = paramiko.Transport((host, port))
    transport.connect(username=username, password=password)
    return paramiko.SFTPClient.from_transport(transport)


def list_remote_files(sftp: paramiko.SFTPClient, remote_path: str) -> List[str]:
    try:
        return sftp.listdir(remote_path)
    except FileNotFoundError:
        logging.error("远程路径不存在：%s", remote_path)
        return []
    except Exception as exc:
        logging.warning("列举远程文件失败：%s", exc)
        return []


def download_file(sftp: paramiko.SFTPClient, remote_path: str, filename: str, local_dir: Path) -> Path:
    local_dir.mkdir(parents=True, exist_ok=True)
    local_path = local_dir / filename
    remote_full = posixpath.join(remote_path, filename)
    sftp.get(remote_full, str(local_path))
    return local_path


def process_sftp_task(
    host: str,
    port: int,
    username: str,
    password: str,
    remote_path: str,
    objects_path: Path,
    rules_path: Path,
    output_dir: Path,
    temp_dir: Path,
    processed_file: Path,
    config: Dict[str, Any],
) -> None:
    """
    短连接模式：一次连接 -> 下载新文件 -> 处理 -> 关闭连接
    """
    processed_set = load_processed_set(processed_file)
    logging.info("已记录的处理文件数量：%s", len(processed_set))
    new_files_count = 0

    transport: Optional[paramiko.Transport] = None
    sftp: Optional[paramiko.SFTPClient] = None
    try:
        logging.info("连接 SFTP %s:%s ...", host, port)
        transport = paramiko.Transport((host, port))
        transport.connect(username=username, password=password)
        sftp = paramiko.SFTPClient.from_transport(transport)
        logging.info("SFTP 连接成功")

        files = list_remote_files(sftp, remote_path)
        logging.info("远程目录文件数：%s", len(files))

        for fname in files:
            suffix = Path(fname).suffix.lower()
            if suffix not in [".xlsx", ".xls"]:
                continue
            if fname in processed_set:
                continue

            try:
                logging.info("发现新文件，开始下载：%s", fname)
                local_path = download_file(sftp, remote_path, fname, temp_dir)
                # 根据文件类型分流处理
                if is_handover_file(fname) and config.get("handover_enabled", True):
                    logging.info("识别为切换数据文件，使用切换处理逻辑：%s", fname)
                    process_handover_file(local_path, objects_path, output_dir, config)
                else:
                    process_file(local_path, objects_path, rules_path, output_dir, config)
                append_processed(processed_file, fname)
                processed_set.add(fname)
                new_files_count += 1
                logging.info("文件处理完成并记录：%s", fname)
            except Exception as exc:
                logging.exception("处理文件 %s 时异常：%s", fname, exc)
    except Exception as exc:
        logging.warning("SFTP 连接或处理异常：%s", exc)
    finally:
        if sftp:
            try:
                sftp.close()
                logging.info("SFTP 连接已关闭")
            except Exception:
                pass
        if transport:
            try:
                transport.close()
            except Exception:
                pass
        
        # 如果本轮扫描没有处理新文件，且开启了无异常推送，发送心跳
        if new_files_count == 0 and config.get("send_normal_heartbeat", True):
            ts_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            content = f"#### ✅ 监控系统运行中\n\n> **扫描结果**：未发现新报表文件\n> **扫描时间**：{ts_str}\n\n当前 SFTP 目录暂无待处理文件。"
            try:
                send_dingtalk_notification("✅ 监控运行中", content)
            except Exception as exc:
                logging.warning("发送运行心跳失败：%s", exc)


def cleanup_old_files(dirs: List[Path], days: int) -> None:
    """
    清理指定目录下超过 days 天的文件。
    """
    if days <= 0:
        return
    
    cutoff_time = time.time() - (days * 86400)
    count = 0
    
    for d in dirs:
        if not d.exists():
            continue
        for p in d.iterdir():
            if p.is_file():
                try:
                    if p.stat().st_mtime < cutoff_time:
                        p.unlink()
                        count += 1
                        logging.info("已清理过期文件：%s", p.name)
                except Exception as e:
                    logging.warning("清理文件失败 %s: %s", p, e)

def scan_loop(
    config: Dict[str, Any],
    objects_path: Path,
    rules_path: Path,
    output_dir: Path,
    temp_dir: Path,
    processed_file: Path,
) -> None:
    """
    后台扫描线程主循环：根据 stop/force 事件控制节奏。
    """
    global last_scan_time

    while not scan_stop_event.is_set():
        # 每次循环重新计算间隔，以便配置修改后生效
        current_interval = int(config.get("sleep_minutes", 5) * 60)
        
        # 执行过期文件清理
        try:
            r_days = int(config.get("retention_days", 7))
            cleanup_old_files([temp_dir, output_dir], r_days)
        except Exception as e:
            logging.warning("自动清理文件时发生异常：%s", e)
        
        logging.info("开始本轮 SFTP 扫描任务...")
        process_sftp_task(
            host=str(config["sftp_host"]),
            port=int(config["sftp_port"]),
            username=str(config["sftp_user"]),
            password=str(config["sftp_password"]),
            remote_path=str(config["sftp_remote_path"]),
            objects_path=objects_path,
            rules_path=rules_path,
            output_dir=output_dir,
            temp_dir=temp_dir,
            processed_file=processed_file,
            config=config,
        )
        last_scan_time = datetime.now()
        
        # 计算下一次等待时间
        mode = config.get("schedule_mode", "interval")
        wait_seconds = 300  # default
        
        if mode == "hourly":
            target_minute = int(config.get("schedule_minute", 0))
            now = datetime.now()
            # 计算本小时的目标时间
            target_time = now.replace(minute=target_minute, second=0, microsecond=0)
            # 如果目标时间已过，则推迟到下一小时
            if target_time <= now:
                target_time += timedelta(hours=1)
            wait_seconds = int((target_time - now).total_seconds())
            logging.info("调度模式：整点模式（每小时 %d 分）。下一次扫描时间：%s", target_minute, target_time.strftime("%H:%M:%S"))
        elif mode == "fixed_times":
            raw_times = str(config.get("schedule_fixed_times", "")).replace("，", ",").split(",")
            valid_times = []
            for t_str in raw_times:
                try:
                    dt = datetime.strptime(t_str.strip(), "%H:%M")
                    valid_times.append(dt.time())
                except ValueError:
                    pass
            
            if not valid_times:
                logging.warning("固定时间点模式未配置有效时间，默认休眠 1 小时")
                wait_seconds = 3600
            else:
                valid_times.sort()
                now = datetime.now()
                target_time = None
                
                # 寻找今天还没过的时间点
                for t in valid_times:
                    candidate = now.replace(hour=t.hour, minute=t.minute, second=0, microsecond=0)
                    if candidate > now:
                        target_time = candidate
                        break
                
                # 如果今天的时间点都过了，取明天的第一个时间点
                if target_time is None:
                    t = valid_times[0]
                    target_time = now.replace(hour=t.hour, minute=t.minute, second=0, microsecond=0) + timedelta(days=1)
                
                wait_seconds = int((target_time - now).total_seconds())
                time_str = ",".join(t.strftime("%H:%M") for t in valid_times)
                logging.info("调度模式：固定时间点 (%s)。下一次扫描时间：%s", time_str, target_time.strftime("%Y-%m-%d %H:%M:%S"))
        else:
            wait_seconds = int(config.get("sleep_minutes", 5) * 60)
            logging.info("调度模式：间隔模式。下一次扫描将在 %d 秒后进行。", wait_seconds)

        waited = 0
        while waited < wait_seconds and not scan_stop_event.is_set():
            if scan_force_event.is_set():
                logging.info("收到手动扫描请求，立即开始新一轮扫描。")
                scan_force_event.clear()
                break
            time.sleep(1)
            waited += 1


class TextHandler(logging.Handler):
    """
    自定义日志处理器，将日志输出到 Tkinter 的 ScrolledText 控件。
    使用 after 方法确保线程安全。
    """
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.see(tk.END)
            self.text_widget.configure(state='disabled')
        # 调度到主线程执行
        self.text_widget.after(0, append)


class MonitorApp:
    def __init__(self, root, args):
        self.root = root
        self.args = args
        self.root.title("指标监控工具")
        self.root.geometry("700x650")
        
        # 拦截关闭事件，改为最小化
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

        # 加载配置
        self.config = self.load_config_gui()
        
        # 初始化钉钉环境变量
        if self.config.get("dingtalk_webhook"):
            os.environ["DINGTALK_WEBHOOK"] = self.config["dingtalk_webhook"]
        if self.config.get("dingtalk_secret"):
            os.environ["DINGTALK_SECRET"] = self.config["dingtalk_secret"]
        
        # 界面布局
        self.create_widgets()
        
        # 配置日志重定向
        self.setup_gui_logger()
        
        # 启动托盘图标（在独立线程）
        self.tray_thread = threading.Thread(target=self.setup_tray, daemon=True)
        self.tray_thread.start()

        # 启动监控线程
        self.start_monitor_thread()

    def load_config_gui(self) -> Dict[str, Any]:
        if CONFIG_PATH.exists():
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                logging.error(f"读取配置失败: {e}")
        return get_default_config()

    def save_config_gui(self):
        # 从界面获取值更新 config 字典
        try:
            self.config["sftp_host"] = self.var_host.get().strip()
            self.config["sftp_port"] = int(self.var_port.get())
            self.config["sftp_user"] = self.var_user.get().strip()
            self.config["sftp_password"] = self.var_pass.get()
            self.config["sftp_remote_path"] = self.var_path.get().strip()
            self.config["sleep_minutes"] = float(self.var_sleep.get())
            self.config["dingtalk_webhook"] = self.var_webhook.get().strip()
            self.config["dingtalk_secret"] = self.var_secret.get().strip()
            self.config["send_normal_heartbeat"] = self.var_send_normal.get()
            self.config["schedule_mode"] = self.var_schedule_mode.get()
            self.config["schedule_minute"] = int(self.var_schedule_minute.get())
            self.config["schedule_fixed_times"] = self.var_schedule_fixed.get().strip()
            self.config["retention_days"] = int(self.var_retention.get())
            self.config["handover_enabled"] = self.var_ho_enabled.get()
            self.config["handover_success_rate_threshold"] = float(self.var_ho_threshold.get())
            self.config["handover_min_attempts"] = int(self.var_ho_min_attempts.get())
            
            # 校验一下
            validate_config(self.config)
            
            # 写入文件
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            
            # 更新环境变量
            if self.config["dingtalk_webhook"]:
                os.environ["DINGTALK_WEBHOOK"] = self.config["dingtalk_webhook"]
            if self.config.get("dingtalk_secret"):
                os.environ["DINGTALK_SECRET"] = self.config["dingtalk_secret"]
                
            logging.info("配置已保存，下一次扫描周期生效。")
            messagebox.showinfo("成功", "配置已保存！")
        except ValueError as e:
            messagebox.showerror("配置错误", str(e))
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")

    def create_widgets(self):
        # 选项卡控件
        tab_control = ttk.Notebook(self.root)
        
        # Tab 1: 运行日志
        tab_logs = ttk.Frame(tab_control)
        tab_control.add(tab_logs, text='运行日志')
        
        self.log_text = scrolledtext.ScrolledText(tab_logs, state='disabled', font=('Consolas', 9))
        self.log_text.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Tab 2: 参数配置
        tab_settings = ttk.Frame(tab_control)
        tab_control.add(tab_settings, text='参数配置')
        
        # 配置表单
        form_frame = ttk.Frame(tab_settings, padding="20")
        form_frame.pack(fill='both', expand=True)
        
        def add_row(parent, label, var, row, is_password=False):
            ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=5)
            entry = ttk.Entry(parent, textvariable=var, width=40, show="*" if is_password else None)
            entry.grid(row=row, column=1, sticky=tk.W, padx=10, pady=5)

        self.var_host = tk.StringVar(value=self.config.get("sftp_host", ""))
        self.var_port = tk.StringVar(value=str(self.config.get("sftp_port", 22)))
        self.var_user = tk.StringVar(value=self.config.get("sftp_user", ""))
        self.var_pass = tk.StringVar(value=self.config.get("sftp_password", ""))
        self.var_path = tk.StringVar(value=self.config.get("sftp_remote_path", ""))
        self.var_webhook = tk.StringVar(value=self.config.get("dingtalk_webhook", ""))
        self.var_secret = tk.StringVar(value=self.config.get("dingtalk_secret", ""))
        self.var_send_normal = tk.BooleanVar(value=self.config.get("send_normal_heartbeat", True))
        
        # 调度相关变量
        self.var_sleep = tk.StringVar(value=str(self.config.get("sleep_minutes", 5)))
        self.var_schedule_mode = tk.StringVar(value=self.config.get("schedule_mode", "interval"))
        self.var_schedule_minute = tk.StringVar(value=str(self.config.get("schedule_minute", 0)))
        self.var_schedule_fixed = tk.StringVar(value=str(self.config.get("schedule_fixed_times", "08:00,12:00,18:00")))
        self.var_retention = tk.StringVar(value=str(self.config.get("retention_days", 7)))
        
        # 切换监控相关变量
        self.var_ho_enabled = tk.BooleanVar(value=self.config.get("handover_enabled", True))
        self.var_ho_threshold = tk.StringVar(value=str(self.config.get("handover_success_rate_threshold", 0.95)))
        self.var_ho_min_attempts = tk.StringVar(value=str(self.config.get("handover_min_attempts", 5)))

        add_row(form_frame, "SFTP 地址:", self.var_host, 0)
        add_row(form_frame, "SFTP 端口:", self.var_port, 1)
        add_row(form_frame, "SFTP 用户:", self.var_user, 2)
        add_row(form_frame, "SFTP 密码:", self.var_pass, 3, is_password=True)
        add_row(form_frame, "远程路径:", self.var_path, 4)
        
        # 调度模式选择区域
        ttk.Label(form_frame, text="调度模式:").grid(row=5, column=0, sticky=tk.W, pady=5)
        mode_frame = ttk.Frame(form_frame)
        mode_frame.grid(row=5, column=1, sticky=tk.W, padx=10)
        
        rb1 = ttk.Radiobutton(mode_frame, text="间隔模式 (每隔X分钟)", variable=self.var_schedule_mode, value="interval")
        rb1.pack(side=tk.LEFT, padx=5)
        rb2 = ttk.Radiobutton(mode_frame, text="整点模式 (每小时第X分)", variable=self.var_schedule_mode, value="hourly")
        rb2.pack(side=tk.LEFT, padx=5)
        rb3 = ttk.Radiobutton(mode_frame, text="固定时间点 (HH:MM)", variable=self.var_schedule_mode, value="fixed_times")
        rb3.pack(side=tk.LEFT, padx=5)

        # 调度参数输入
        # 为了界面简洁，这里放两行，根据模式不同，用户填写对应行即可（或者都填也不影响，只取生效的）
        # 更友好的方式是动态显示，但简单起见，我们列出两个输入框
        
        ttk.Label(form_frame, text="[间隔模式] 休眠分钟:").grid(row=6, column=0, sticky=tk.W, pady=5)
        ttk.Entry(form_frame, textvariable=self.var_sleep, width=10).grid(row=6, column=1, sticky=tk.W, padx=10)
        
        ttk.Label(form_frame, text="[整点模式] 第几分钟(0-59):").grid(row=7, column=0, sticky=tk.W, pady=5)
        ttk.Entry(form_frame, textvariable=self.var_schedule_minute, width=10).grid(row=7, column=1, sticky=tk.W, padx=10)
        
        ttk.Label(form_frame, text="[固定时间] (逗号分隔):").grid(row=8, column=0, sticky=tk.W, pady=5)
        ttk.Entry(form_frame, textvariable=self.var_schedule_fixed, width=40).grid(row=8, column=1, sticky=tk.W, padx=10)

        add_row(form_frame, "文件保留天数:", self.var_retention, 9)
        add_row(form_frame, "钉钉Webhook:", self.var_webhook, 10)
        add_row(form_frame, "钉钉加签密钥:", self.var_secret, 11, is_password=True)
        
        ttk.Label(form_frame, text="无异常推送:").grid(row=12, column=0, sticky=tk.W, pady=5)
        ttk.Checkbutton(form_frame, text="开启无异常时的监控正常通知", variable=self.var_send_normal).grid(row=12, column=1, sticky=tk.W, padx=10, pady=5)

        # 分隔线 + 切换监控配置区域
        ttk.Separator(form_frame, orient='horizontal').grid(row=13, column=0, columnspan=2, sticky='ew', pady=10)
        ttk.Label(form_frame, text="── 邻区对切换监控 ──", font=('', 9, 'bold')).grid(row=14, column=0, columnspan=2, sticky=tk.W, pady=2)

        ttk.Label(form_frame, text="切换监控:").grid(row=15, column=0, sticky=tk.W, pady=5)
        ttk.Checkbutton(form_frame, text="启用邻区对切换异常监控", variable=self.var_ho_enabled).grid(row=15, column=1, sticky=tk.W, padx=10, pady=5)

        add_row(form_frame, "切换成功率阈值(0~1):", self.var_ho_threshold, 16)
        add_row(form_frame, "最低请求次数:", self.var_ho_min_attempts, 17)

        btn_frame = ttk.Frame(tab_settings)
        btn_frame.pack(fill='x', pady=10, padx=20)
        ttk.Button(btn_frame, text="保存配置", command=self.save_config_gui).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="立即扫描", command=lambda: scan_force_event.set()).pack(side=tk.LEFT, padx=5)

        tab_control.pack(expand=True, fill='both')

    def setup_gui_logger(self):
        handler = TextHandler(self.log_text)
        formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
        handler.setFormatter(formatter)
        logging.getLogger().addHandler(handler)
        logging.getLogger().setLevel(logging.INFO)

    def start_monitor_thread(self):
        objects_path = Path(self.args.objects)
        rules_path = Path(self.args.rules)
        output_dir = Path(self.args.output)
        temp_dir = Path(self.args.temp)
        processed_file = Path(self.args.processed)

        # 确保目录存在
        output_dir.mkdir(parents=True, exist_ok=True)
        temp_dir.mkdir(parents=True, exist_ok=True)

        worker = threading.Thread(
            target=scan_loop,
            args=(self.config, objects_path, rules_path, output_dir, temp_dir, processed_file),
            daemon=True,
        )
        worker.start()

    def hide_window(self):
        self.root.withdraw()

    def show_window(self):
        self.root.deiconify()
        self.root.lift()

    def quit_app(self):
        scan_stop_event.set()
        if hasattr(self, 'icon'):
            self.icon.stop()
        self.root.quit()

    def setup_tray(self):
        image = Image.new("RGB", (64, 64), (0, 128, 255))
        draw = ImageDraw.Draw(image)
        draw.rectangle((16, 16, 48, 48), fill=(255, 255, 255))
        
        menu = pystray.Menu(
            MenuItem("显示主界面", lambda icon, item: self.root.after(0, self.show_window)),
            MenuItem("立即扫描", lambda icon, item: scan_force_event.set()),
            MenuItem("退出程序", lambda icon, item: self.root.after(0, self.quit_app))
        )
        self.icon = pystray.Icon("VIPMonitor", image, "VIP区域指标监控", menu)
        self.icon.run()

def main() -> None:
    parser = argparse.ArgumentParser(description="网优报表自动监控告警（SFTP 轮询，动态 Rules/Objects）")
    parser.add_argument("--objects", default=str(DEFAULT_OBJECTS_PATH), help="Objects.xlsx 路径（列：小区名称）")
    parser.add_argument("--rules", default=str(DEFAULT_RULES_PATH), help="Rules.xlsx 路径（列：监控指标/判断符/阈值）")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT_DIR), help="告警输出目录（默认 ./output）")
    parser.add_argument("--temp", default=str(DEFAULT_TEMP_DIR), help="下载临时目录（默认 ./temp_data）")
    parser.add_argument("--processed", default=str(DEFAULT_PROCESSED_FILE), help="已处理文件记录（默认 ./processed_files.txt）")
    parser.add_argument("--log-level", default="INFO", help="日志级别：DEBUG/INFO/WARNING/ERROR")
    args = parser.parse_args()
    
    root = tk.Tk()
    app = MonitorApp(root, args)
    root.mainloop()


if __name__ == "__main__":
    main()
