#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
考勤自动审核脚本 v1.1
====================
功能：自动扫描本地批次文件，并发调用豆包2.0 Pro API进行考勤审核，自动汇总结果

核心特性：
    - 并发处理：支持多线程并发处理多个批次文件
    - 自动重试：API调用失败自动重试，提高成功率
    - 断点续传：记录处理状态，中断后可继续处理
    - 格式校验：自动验证AI返回结果是否为标准CSV格式
    - 结果汇总：合并所有批次结果，生成Excel和TXT格式汇总文件

依赖库：
    - pandas：数据处理和CSV/Excel文件读写
    - openpyxl：Excel文件生成
    - requests：HTTP请求库，直接调用豆包API
    - concurrent.futures：Python标准库，多线程并发

安装命令：
    pip install pandas openpyxl requests
"""

import os
import sys
import re
import time
import json
import logging
import signal
import requests
import pandas as pd
from io import StringIO
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

# 全局变量，用于优雅中断
_shutdown_requested = False


def _signal_handler(signum, frame):
    """处理Ctrl+C中断信号"""
    global _shutdown_requested
    _shutdown_requested = True
    print("\n\n⚠️ 收到中断信号，正在安全退出...")

# ============================================================
# ========== 核心配置项 请修改此处 ==========
# ============================================================

# 【文件配置】
RULE_FILE_NAME = "考勤规则V3.7（5月考勤规则）.TXT"  # 规则文件名，需与脚本放在同一目录🔴🔴🔴
BATCH_FOLDER = "考勤批次数据"        # 批次文件夹名，存放所有待审核的批次TXT文件

# 【并发配置】
CONCURRENT_BATCHES = 5               # 并发处理数量，建议固定为5
RETRY_TIMES = 2                      # API调用失败重试次数，默认2次

# 【API配置】
# 🔴🔴🔴 【高亮！必须修改】以下是唯一需要手动填写的内容 🔴🔴🔴
API_KEY = os.environ.get("DOUBAO_API_KEY", "")  # 从环境变量读取豆包API密钥
# 🔴🔴🔴 上方是唯一需要手动填写的配置项 🔴🔴🔴

API_BASE_URL = "https://ark.cn-beijing.volces.com/api/v3"  # 豆包API地址（火山引擎Ark服务）
MODEL_NAME = "doubao-seed-2-0-pro-260215"  # 豆包2.0 Pro模型（注意：是2-0不是2.0）
TEMPERATURE = 0.3  # 模型随机性参数，默认0.3（低随机性+保留语义理解），范围0-1，数值越高输出越随机

# 【请求配置】
API_TIMEOUT = 180  # API请求超时时间（秒），默认3分钟
MAX_TOKENS = 16000  # API返回最大token数，根据批次数据量调整

# ============================================================
# ========== 核心配置项 结束 ==========
# ============================================================

# 固定表头（AI必须返回此格式的CSV）
CSV_HEADER = "序号,考勤异常汇总,非考勤内容"

# 系统提示词（约束AI仅输出CSV格式）
SYSTEM_PROMPT = (
    "你是一个严格的考勤审核员，必须100%按照用户提供的考勤规则执行校验，"
    "仅输出标准CSV格式内容，禁止输出任何其他内容，"
    "包括解释、思考、开场白、结束语"
)


class AttendanceAuditor:
    """考勤审核器主类
    
    负责整个考勤审核流程的调度和执行，包括：
    - 初始化和配置加载
    - 批次文件扫描
    - 并发API调用
    - 结果验证和保存
    - 最终结果汇总
    """

    def __init__(self):
        """初始化审核器，加载配置和必要资源"""
        # 获取脚本所在目录
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 构建各文件/目录路径
        self.rule_file_path = os.path.join(self.script_dir, RULE_FILE_NAME)
        self.batch_folder_path = os.path.join(self.script_dir, BATCH_FOLDER)
        self.result_backup_folder = os.path.join(self.script_dir, "批次结果备份")
        self.log_folder = os.path.join(self.script_dir, "审核日志")
        
        # 运行时变量
        self.rule_content = ""           # 考勤规则全文内容
        self.log_file = None             # 当前日志文件路径
        self.processed_batches = set()   # 已处理成功的批次名称集合
        
        # 初始化各模块
        self._cleanup_old_results()      # 清理上次运行的结果
        self._init_logging()             # 初始化日志系统（清理后创建）
        self._create_directories()       # 创建必要目录
        self._load_rule_file()           # 加载规则文件
        self._load_processed_records()   # 加载已处理批次记录
        self._init_api_session()         # 初始化HTTP会话

    def _cleanup_old_results(self):
        """清理上次运行的结果数据
        
        清理内容：
        - 批次结果备份文件夹中的所有文件
        - 审核日志文件夹中的所有文件
        - 考勤总审核结果.xlsx
        - 考勤总审核结果.txt
        
        保留内容：
        - 运行说明.txt
        - 考勤规则文件
        - 考勤批次数据文件夹及其内容
        - 脚本文件本身
        """
        cleaned_count = 0
        
        # 清理批次结果备份文件夹
        if os.path.exists(self.result_backup_folder):
            for filename in os.listdir(self.result_backup_folder):
                file_path = os.path.join(self.result_backup_folder, filename)
                try:
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                        cleaned_count += 1
                except Exception as e:
                    print(f"⚠️ 清理文件失败: {file_path}, 错误: {e}")
        
        # 清理审核日志文件夹
        if os.path.exists(self.log_folder):
            for filename in os.listdir(self.log_folder):
                file_path = os.path.join(self.log_folder, filename)
                try:
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                        cleaned_count += 1
                except Exception as e:
                    print(f"⚠️ 清理文件失败: {file_path}, 错误: {e}")
        
        # 清理汇总结果文件
        for result_file in ["考勤总审核结果.xlsx", "考勤总审核结果.txt"]:
            result_path = os.path.join(self.script_dir, result_file)
            if os.path.exists(result_path):
                try:
                    os.remove(result_path)
                    cleaned_count += 1
                except Exception as e:
                    print(f"⚠️ 清理文件失败: {result_path}, 错误: {e}")
        
        # 重置已处理批次记录
        self.processed_batches = set()
        
        print(f"🧹 已清理 {cleaned_count} 个上次运行的结果文件")

    def _init_logging(self):
        """初始化日志系统
        
        创建日志文件并配置日志记录器，同时输出到文件和控制台。
        日志文件名格式：审核日志_YYYYMMDD_HHMMSS.txt
        """
        # 确保日志目录存在
        os.makedirs(self.log_folder, exist_ok=True)
        
        # 生成日志文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = os.path.join(self.log_folder, f"审核日志_{timestamp}.txt")
        
        # 配置日志记录器
        self.logger = logging.getLogger("AttendanceAuditor")
        self.logger.setLevel(logging.INFO)
        
        # 清除已有处理器（避免重复添加）
        self.logger.handlers.clear()
        
        # 文件处理器：输出到日志文件
        file_handler = logging.FileHandler(self.log_file, encoding="utf-8")
        file_handler.setLevel(logging.INFO)
        file_format = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        file_handler.setFormatter(file_format)
        self.logger.addHandler(file_handler)
        
        # 控制台处理器：输出到终端
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        console_format = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        console_handler.setFormatter(console_format)
        self.logger.addHandler(console_handler)
        
        # 输出启动信息
        self.logger.info("=" * 60)
        self.logger.info("考勤自动审核脚本 v1.1 启动")
        self.logger.info(f"日志文件: {self.log_file}")
        self.logger.info("=" * 60)

    def _create_directories(self):
        """创建必要的目录（如果不存在）
        
        创建以下目录：
        - 批次结果备份：存放每个批次的原始审核结果
        - 审核日志：存放运行日志和处理状态记录
        """
        dirs_to_create = [
            self.result_backup_folder,
            self.log_folder
        ]
        for dir_path in dirs_to_create:
            os.makedirs(dir_path, exist_ok=True)
            self.logger.info(f"目录已创建/已存在: {dir_path}")

    def _load_rule_file(self):
        """加载考勤规则文件
        
        读取规则文件内容，供后续发送给AI作为审核依据。
        如果文件不存在或读取失败，直接退出程序。
        """
        if not os.path.exists(self.rule_file_path):
            self.logger.error(f"❌ 规则文件不存在: {self.rule_file_path}")
            self.logger.error("请将考勤规则文件放在与脚本相同的目录下，然后重新运行脚本。")
            sys.exit(1)
        
        try:
            with open(self.rule_file_path, "r", encoding="utf-8") as f:
                self.rule_content = f.read()
            self.logger.info(f"✅ 规则文件加载成功: {RULE_FILE_NAME}")
            self.logger.info(f"规则内容长度: {len(self.rule_content)} 字符")
        except Exception as e:
            self.logger.error(f"❌ 读取规则文件失败: {e}")
            sys.exit(1)

    def _init_api_session(self):
        """初始化HTTP会话
        
        创建可复用的requests会话实例，支持连接复用。
        """
        try:
            self.api_session = requests.Session()
            self.api_session.headers.update({
                "Authorization": f"Bearer {API_KEY.strip()}",
                "Content-Type": "application/json"
            })
            self.logger.info("✅ HTTP会话初始化成功")
        except Exception as e:
            self.logger.error(f"❌ HTTP会话初始化失败: {e}")
            sys.exit(1)

    def _load_processed_records(self):
        """加载已处理成功的批次记录
        
        从JSON文件中读取之前已成功处理的批次列表，
        实现断点续传功能，避免重复处理。
        """
        status_file = os.path.join(self.log_folder, "处理状态记录.json")
        if os.path.exists(status_file):
            try:
                with open(status_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.processed_batches = set(data.get("processed_batches", []))
                self.logger.info(f"已加载处理状态记录，已处理批次: {len(self.processed_batches)} 个")
            except Exception as e:
                self.logger.warning(f"读取处理状态记录失败: {e}，将从头开始处理")
                self.processed_batches = set()
        else:
            self.logger.info("未找到处理状态记录，将从头开始处理所有批次")

    def _save_processed_record(self, batch_name, status, process_time, error_count=0, error_msg=""):
        """保存批次处理状态记录
        
        参数:
            batch_name (str): 批次名称
            status (str): 处理状态，"success" 或 "failed"
            process_time (float): 处理耗时（秒）
            error_count (int): 异常数量
            error_msg (str): 错误信息
        """
        status_file = os.path.join(self.log_folder, "处理状态记录.json")
        
        # 读取现有记录
        data = {"processed_batches": [], "batch_details": {}}
        if os.path.exists(status_file):
            try:
                with open(status_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                pass
        
        # 更新记录：成功则加入已处理列表
        if status == "success":
            if batch_name not in data["processed_batches"]:
                data["processed_batches"].append(batch_name)
            self.processed_batches.add(batch_name)
        
        # 记录批次详情
        data["batch_details"][batch_name] = {
            "status": status,
            "process_time": round(process_time, 2),
            "error_count": error_count,
            "error_msg": error_msg,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        # 保存记录
        with open(status_file, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def scan_batch_files(self):
        """扫描批次文件夹，获取所有待处理的TXT文件
        
        返回:
            list: 待处理批次列表，每个元素为 (batch_name, file_path) 元组
        """
        if not os.path.exists(self.batch_folder_path):
            self.logger.error(f"❌ 批次文件夹不存在: {self.batch_folder_path}")
            self.logger.error("请在脚本目录下创建「考勤批次数据」文件夹，并将批次文件放入其中。")
            sys.exit(1)
        
        batch_files = []
        for filename in os.listdir(self.batch_folder_path):
            if filename.endswith(".txt"):
                file_path = os.path.join(self.batch_folder_path, filename)
                batch_name = filename.replace(".txt", "")
                
                # 跳过已处理成功的批次
                if batch_name not in self.processed_batches:
                    batch_files.append((batch_name, file_path))
                else:
                    self.logger.info(f"⏭️ 跳过已处理批次: {batch_name}")
        
        # 按批次编号排序（提取数字部分进行排序）
        def sort_key(x):
            match = re.search(r'(\d+)', x[0])
            return int(match.group(1)) if match else 0
        
        batch_files.sort(key=sort_key)
        
        self.logger.info(f"📋 扫描到 {len(batch_files)} 个待处理批次文件")
        for name, path in batch_files:
            self.logger.info(f"  - {name}")
        
        return batch_files

    def read_batch_file(self, file_path):
        """读取批次文件内容
        
        参数:
            file_path (str): 批次文件路径
        
        返回:
            str: 文件内容，读取失败返回None
        """
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read()
            self.logger.info(f"✅ 批次文件读取成功: {os.path.basename(file_path)}")
            return content
        except Exception as e:
            self.logger.error(f"❌ 读取批次文件失败: {file_path}, 错误: {e}")
            return None

    def call_doubao_api(self, batch_content, batch_name, retry_count=0):
        """调用豆包API进行考勤审核
        
        参数:
            batch_content (str): 批次文件内容
            batch_name (str): 批次名称
            retry_count (int): 当前重试次数
        
        返回:
            str: AI返回的内容，调用失败返回None
        """
        # 构建用户输入：规则 + 批次数据
        user_input = (
            f"【考勤规则V3.0全文】\n{self.rule_content}\n\n"
            f"【当前批次考勤数据】\n{batch_content}\n\n"
            f"请按照上述规则对当前批次数据进行考勤审核，输出标准CSV格式结果。"
        )
        
        try:
            self.logger.info(
                f"📡 正在调用豆包API处理批次: {batch_name} "
                f"(尝试 {retry_count + 1}/{RETRY_TIMES + 1})"
            )
            
            # 构建请求数据
            payload = {
                "model": MODEL_NAME,
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_input}
                ],
                "temperature": TEMPERATURE,
                "max_tokens": MAX_TOKENS
            }
            
            # 调用API
            api_url = f"{API_BASE_URL}/chat/completions"
            response = self.api_session.post(
                api_url,
                json=payload,
                timeout=API_TIMEOUT
            )
            response.raise_for_status()
            
            # 解析响应
            result = response.json()
            
            # 提取AI回复内容
            if "choices" in result and len(result["choices"]) > 0:
                content = result["choices"][0]["message"]["content"]
                self.logger.info(f"✅ API调用成功，批次: {batch_name}")
                return content
            else:
                self.logger.error(f"❌ API返回格式异常: {batch_name}")
                return None
        
        except requests.exceptions.RequestException as e:
            self.logger.error(f"❌ API调用失败: {batch_name}, 错误: {e}")
            return None
        except Exception as e:
            self.logger.error(f"❌ API调用失败: {batch_name}, 错误: {e}")
            return None

    def validate_csv_format(self, content):
        """验证返回内容是否为标准CSV格式
        
        参数:
            content (str): AI返回的内容
        
        返回:
            tuple: (is_valid: bool, result: str)
                   is_valid为True时，result为清理后的CSV内容
                   is_valid为False时，result为错误原因
        """
        if not content:
            return False, "内容为空"
        
        # 清理内容：移除可能的markdown代码块标记
        cleaned_content = content.strip()
        if cleaned_content.startswith("```"):
            # 移除markdown代码块
            lines = cleaned_content.split("\n")
            # 找到第一个包含表头的行
            start_idx = 0
            for i, line in enumerate(lines):
                if "序号" in line and "考勤异常汇总" in line:
                    start_idx = i
                    break
            # 找到最后一个数据行
            end_idx = len(lines)
            for i in range(len(lines) - 1, -1, -1):
                if lines[i].strip() and not lines[i].startswith("```"):
                    end_idx = i + 1
                    break
            cleaned_content = "\n".join(lines[start_idx:end_idx])
        
        # 检查是否包含固定表头
        if CSV_HEADER not in cleaned_content:
            return False, f"缺少标准表头，期望: {CSV_HEADER}"
        
        # 尝试解析CSV
        try:
            df = pd.read_csv(StringIO(cleaned_content))
            required_columns = ["序号", "考勤异常汇总", "非考勤内容"]
            if not all(col in df.columns for col in required_columns):
                return False, f"列名不匹配，期望: {required_columns}, 实际: {list(df.columns)}"
            return True, cleaned_content
        except Exception as e:
            return False, f"CSV解析失败: {e}"

    def process_single_batch(self, batch_name, file_path):
        """处理单个批次文件
        
        完整处理流程：读取文件 → 调用API → 验证格式 → 保存结果
        
        参数:
            batch_name (str): 批次名称
            file_path (str): 批次文件路径
        
        返回:
            DataFrame: 解析后的CSV数据，处理失败返回None
        """
        self.logger.info(f"🔄 开始处理批次: {batch_name}")
        start_time = time.time()
        
        # 步骤1: 读取批次文件
        batch_content = self.read_batch_file(file_path)
        if not batch_content:
            self._save_processed_record(batch_name, "failed", 0, error_msg="文件读取失败")
            return None
        
        # 步骤2: 调用API（带重试机制）
        api_result = None
        for retry in range(RETRY_TIMES + 1):
            api_result = self.call_doubao_api(batch_content, batch_name, retry)
            
            if api_result:
                # 验证CSV格式
                is_valid, result = self.validate_csv_format(api_result)
                if is_valid:
                    api_result = result
                    break
                else:
                    self.logger.warning(f"⚠️ CSV格式验证失败: {batch_name}, 原因: {result}")
                    if retry < RETRY_TIMES:
                        self.logger.info(f"🔄 准备重试 (第 {retry + 2} 次)")
                        time.sleep(2)  # 重试前等待2秒
                    else:
                        self.logger.error(f"❌ CSV格式验证失败，已达最大重试次数: {batch_name}")
                        api_result = None
            else:
                if retry < RETRY_TIMES:
                    self.logger.info(f"🔄 API调用失败，准备重试 (第 {retry + 2} 次)")
                    time.sleep(3)  # 重试前等待3秒
                else:
                    self.logger.error(f"❌ API调用失败，已达最大重试次数: {batch_name}")
        
        # 步骤3: 处理结果
        if not api_result:
            process_time = time.time() - start_time
            self._save_processed_record(batch_name, "failed", process_time, error_msg="API调用或格式验证失败")
            return None
        
        # 步骤4: 保存结果到备份文件夹
        result_filename = f"{batch_name}_结果.csv"
        result_path = os.path.join(self.result_backup_folder, result_filename)
        try:
            with open(result_path, "w", encoding="utf-8") as f:
                f.write(api_result)
            self.logger.info(f"💾 批次结果已保存: {result_filename}")
        except Exception as e:
            self.logger.error(f"❌ 保存批次结果失败: {result_path}, 错误: {e}")
        
        # 步骤5: 记录处理状态
        process_time = time.time() - start_time
        self._save_processed_record(batch_name, "success", process_time)
        
        # 步骤6: 解析CSV结果
        try:
            df = pd.read_csv(StringIO(api_result))
            record_count = len(df)
            self.logger.info(
                f"✅ 批次处理完成: {batch_name}, "
                f"耗时: {process_time:.2f}秒, "
                f"记录数: {record_count}"
            )
            return df
        except Exception as e:
            self.logger.error(f"❌ 解析CSV结果失败: {batch_name}, 错误: {e}")
            return None

    def merge_and_save_results(self, all_results):
        """合并所有批次结果并保存
        
        参数:
            all_results (list): 所有批次的DataFrame列表
        """
        self.logger.info("=" * 60)
        self.logger.info("📊 开始汇总所有批次结果")
        
        if not all_results:
            self.logger.warning("⚠️ 没有有效的处理结果，无法生成汇总文件")
            return
        
        # 合并所有DataFrame
        merged_df = pd.concat(all_results, ignore_index=True)
        
        # 按序号排序（提取序号中的数字部分进行排序）
        def sort_key(x):
            try:
                match = re.search(r'(\d+)', str(x))
                return int(match.group(1)) if match else 0
            except Exception:
                return 0
        
        merged_df = merged_df.sort_values(
            by="序号", 
            key=lambda x: x.apply(sort_key)
        ).reset_index(drop=True)
        
        # 保存为Excel文件
        excel_path = os.path.join(self.script_dir, "考勤总审核结果.xlsx")
        try:
            merged_df.to_excel(excel_path, index=False, engine="openpyxl")
            self.logger.info(f"✅ Excel汇总结果已保存: {excel_path}")
        except Exception as e:
            self.logger.error(f"❌ 保存Excel汇总结果失败: {e}")
        
        # 保存为TXT文件（CSV格式，UTF-8 BOM编码）
        txt_path = os.path.join(self.script_dir, "考勤总审核结果.txt")
        try:
            merged_df.to_csv(txt_path, index=False, encoding="utf-8-sig")
            self.logger.info(f"✅ TXT汇总结果已保存: {txt_path}")
        except Exception as e:
            self.logger.error(f"❌ 保存TXT汇总结果失败: {e}")
        
        self.logger.info(f"📊 汇总结果共 {len(merged_df)} 条记录")
        self.logger.info("=" * 60)

    def run(self):
        """主运行流程
        
        执行步骤：
        1. 扫描批次文件
        2. 并发处理批次
        3. 合并结果
        4. 输出最终统计
        """
        global _shutdown_requested
        start_time = time.time()
        
        self.logger.info("🚀 开始执行考勤自动审核流程")
        
        # 步骤1: 扫描批次文件
        batch_files = self.scan_batch_files()
        if not batch_files:
            self.logger.info("✅ 所有批次文件已处理完成，无需重复处理")
            # 仍然尝试合并已有结果
            self._merge_existing_results()
            return
        
        # 步骤2: 并发处理批次
        all_results = []
        total_batches = len(batch_files)
        processed_count = 0
        
        # 按并发数量分组处理
        for i in range(0, total_batches, CONCURRENT_BATCHES):
            # 检查是否收到中断信号
            if _shutdown_requested:
                self.logger.warning("⚠️ 用户中断，已处理的批次结果已保存")
                break
            
            batch_group = batch_files[i:i + CONCURRENT_BATCHES]
            group_num = i // CONCURRENT_BATCHES + 1
            self.logger.info(f"📦 处理批次组 {group_num}: 共 {len(batch_group)} 个批次")
            
            # 使用线程池并发处理
            with ThreadPoolExecutor(max_workers=CONCURRENT_BATCHES) as executor:
                # 提交任务
                future_to_batch = {
                    executor.submit(self.process_single_batch, name, path): name
                    for name, path in batch_group
                }
                
                # 等待完成
                for future in as_completed(future_to_batch):
                    if _shutdown_requested:
                        self.logger.warning("⚠️ 用户中断，取消剩余任务")
                        executor.shutdown(wait=False, cancel_futures=True)
                        break
                    
                    batch_name = future_to_batch[future]
                    processed_count += 1
                    try:
                        result = future.result()
                        if result is not None:
                            all_results.append(result)
                        self.logger.info(
                            f"📈 进度: {processed_count}/{total_batches} 批次已完成"
                        )
                    except Exception as e:
                        self.logger.error(f"❌ 批次处理异常: {batch_name}, 错误: {e}")
        
        # 步骤3: 合并结果（即使中断也合并已处理的结果）
        if all_results:
            self.merge_and_save_results(all_results)
        else:
            self.logger.warning("⚠️ 没有有效的处理结果，无法生成汇总文件")
        
        # 步骤4: 输出最终统计
        total_time = time.time() - start_time
        self.logger.info("=" * 60)
        if _shutdown_requested:
            self.logger.warning("⚠️ 考勤自动审核流程被中断")
        else:
            self.logger.info("🎉 考勤自动审核流程完成")
        self.logger.info(f"📊 总处理批次: {total_batches}")
        self.logger.info(f"✅ 成功处理: {len(all_results)}")
        self.logger.info(f"❌ 失败批次: {total_batches - len(all_results)}")
        self.logger.info(f"⏱️ 总耗时: {total_time:.2f}秒 ({total_time/60:.1f}分钟)")
        self.logger.info(f"📁 结果备份目录: {self.result_backup_folder}")
        self.logger.info(f"📁 审核日志目录: {self.log_folder}")
        self.logger.info("=" * 60)

    def _merge_existing_results(self):
        """合并已有的结果文件
        
        用于所有批次已处理完成的情况，或需要重新生成汇总文件时。
        """
        self.logger.info("📂 尝试合并已有的批次结果文件")
        
        all_results = []
        result_files = []
        
        # 查找所有结果文件
        for filename in os.listdir(self.result_backup_folder):
            if filename.endswith("_结果.csv"):
                result_files.append(os.path.join(self.result_backup_folder, filename))
        
        if not result_files:
            self.logger.warning("⚠️ 未找到任何批次结果文件")
            return
        
        # 加载所有结果
        for file_path in result_files:
            try:
                df = pd.read_csv(file_path)
                all_results.append(df)
                self.logger.info(f"✅ 已加载结果文件: {os.path.basename(file_path)}")
            except Exception as e:
                self.logger.warning(f"⚠️ 加载结果文件失败: {file_path}, 错误: {e}")
        
        # 合并结果
        if all_results:
            self.merge_and_save_results(all_results)


def main():
    """主函数
    
    检查API密钥配置，创建审核器实例并运行。
    """
    # 注册中断信号处理
    signal.signal(signal.SIGINT, _signal_handler)
    signal.signal(signal.SIGTERM, _signal_handler)
    
    print("=" * 60)
    print("考勤自动审核脚本 v1.1")
    print("=" * 60)
    
    # 检查API密钥是否已配置
    if API_KEY == "请在此处粘贴你的豆包API密钥" or not API_KEY or len(API_KEY) < 10:
        print("❌ 错误: 请先配置API密钥！")
        print("请打开脚本，找到【核心配置项】区域，将API_KEY变量设置为你的豆包API密钥。")
        print("获取API密钥方式: 访问 https://console.volcengine.com/ 注册并创建API密钥")
        sys.exit(1)
    
    # 创建审核器实例并运行
    auditor = AttendanceAuditor()
    auditor.run()


if __name__ == "__main__":
    main()