import pandas as pd
import pdfplumber
import os
from pathlib import Path
import glob
import re
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import queue
import time

class PDFToExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF转Excel批量转换工具")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 创建队列用于线程间通信
        self.log_queue = queue.Queue()
        
        # 初始化变量
        self.pdf_folder = tk.StringVar()
        self.excel_folder = tk.StringVar()
        self.is_running = False
        self.success_count = 0
        self.failed_count = 0
        self.total_files = 0
        self.processed_files = 0
        
        # 表头匹配模式
        self.column_pattern = re.compile(r'^\([A-H]\)$')
        
        self.setup_ui()
        self.process_queue()
    
    def setup_ui(self):
        """设置用户界面"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="PDF转Excel批量转换工具", 
                                font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # PDF文件夹选择
        ttk.Label(main_frame, text="PDF文件夹:").grid(row=1, column=0, sticky=tk.W, pady=5)
        pdf_entry = ttk.Entry(main_frame, textvariable=self.pdf_folder, width=50)
        pdf_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(0, 5))
        ttk.Button(main_frame, text="浏览", command=self.browse_pdf_folder).grid(row=1, column=2, pady=5)
        
        # Excel文件夹选择
        ttk.Label(main_frame, text="Excel输出文件夹:").grid(row=2, column=0, sticky=tk.W, pady=5)
        excel_entry = ttk.Entry(main_frame, textvariable=self.excel_folder, width=50)
        excel_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=(0, 5))
        ttk.Button(main_frame, text="浏览", command=self.browse_excel_folder).grid(row=2, column=2, pady=5)
        
        # 选项框架
        options_frame = ttk.LabelFrame(main_frame, text="转换选项", padding="5")
        options_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        ttk.Label(options_frame, text="起始页码:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.start_page = tk.StringVar(value="2")
        ttk.Entry(options_frame, textvariable=self.start_page, width=10).grid(row=0, column=1, sticky=tk.W, pady=5, padx=(5, 0))
        
        ttk.Label(options_frame, text="线程数:").grid(row=0, column=2, sticky=tk.W, pady=5, padx=(20, 0))
        self.thread_count = tk.StringVar(value="1")
        ttk.Spinbox(options_frame, from_=1, to=10, textvariable=self.thread_count, width=10).grid(row=0, column=3, sticky=tk.W, pady=5, padx=(5, 0))
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        self.start_button = ttk.Button(button_frame, text="开始转换", command=self.start_conversion)
        self.start_button.grid(row=0, column=0, padx=(0, 5))
        
        self.stop_button = ttk.Button(button_frame, text="停止转换", command=self.stop_conversion, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=1, padx=(5, 0))
        
        # 进度条
        ttk.Label(main_frame, text="转换进度:").grid(row=5, column=0, sticky=tk.W, pady=(10, 5))
        self.progress_bar = ttk.Progressbar(main_frame, mode='determinate')
        self.progress_bar.grid(row=5, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 5))
        
        # 进度标签
        self.progress_label = ttk.Label(main_frame, text="准备就绪")
        self.progress_label.grid(row=6, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))
        
        # 日志框
        log_frame = ttk.LabelFrame(main_frame, text="转换日志", padding="5")
        log_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 状态栏
        self.status_var = tk.StringVar(value="准备就绪")
        ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).grid(
            row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
    
    def browse_pdf_folder(self):
        folder = filedialog.askdirectory(title="选择PDF文件夹")
        if folder:
            self.pdf_folder.set(folder)
    
    def browse_excel_folder(self):
        folder = filedialog.askdirectory(title="选择Excel输出文件夹")
        if folder:
            self.excel_folder.set(folder)
    
    def log_message(self, message):
        """添加消息到日志"""
        self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_progress(self, current, total, message):
        """更新进度条和标签"""
        if total > 0:
            self.progress_bar['value'] = (current / total) * 100
        self.progress_label.config(text=message)
        self.status_var.set(f"已处理: {current}/{total} 成功: {self.success_count} 失败: {self.failed_count}")
        self.root.update_idletasks()
    
    def process_queue(self):
        """处理队列中的消息"""
        try:
            while True:
                message = self.log_queue.get_nowait()
                if message == "COMPLETED":
                    self.conversion_completed()
                else:
                    self.log_message(message)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.process_queue)
    
    def set_column_widths(self, writer, sheet_name, df):
        """设置Excel列宽"""
        worksheet = writer.sheets[sheet_name]
        for idx, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).str.len().max(), len(str(col))) + 2
            column_letter = chr(65 + idx) if idx < 26 else f"A{chr(65 + idx - 26)}"
            worksheet.column_dimensions[column_letter].width = min(max_len, 50)
    
    def is_valid_date(self, text):
        """验证日期格式"""
        if text is None:
            return False
        date_patterns = [
            r'\d{1,2}/\d{1,2}/\d{4}', r'\d{1,2}-\d{1,2}-\d{4}',
            r'\d{4}/\d{1,2}/\d{1,2}', r'\d{4}-\d{1,2}-\d{1,2}',
            r'\d{4}年\d{1,2}月\d{1,2}日', r'\d{1,2}\.\d{1,2}\.\d{4}',
        ]
        return any(re.search(pattern, str(text)) for pattern in date_patterns)
    
    def is_valid_amount(self, text):
        """验证金额格式"""
        if text is None:
            return False
        amount_patterns = [
            r'\$[\d,]+', r'USD\s*[\d,]+', r'HKD\s*[\d,]+',
            r'HK\$[\d,]+', r'[\d,]+\.?\d*\s*(美元|港币|元)',
        ]
        return any(re.search(pattern, str(text)) for pattern in amount_patterns)
    
    def is_header_row(self, row):
        """判断是否为表头行 - 改进版本：检查(A)-(H)列标识"""
        if not row:
            return False
        
        # 检查行中是否有符合(A)-(H)模式的单元格
        valid_columns = 0
        for cell in row:
            if cell is not None and self.column_pattern.match(str(cell).strip()):
                valid_columns += 1
                
        # 如果有至少3个有效列标识，则认为这是有效的表头行
        return valid_columns >= 3
    
    def contains_house_number(self, row):
        """检查行是否包含屋號信息"""
        if not row:
            return False
        
        house_patterns = [
            r'屋號',
            r'House number',
            r'屋名',
            r'Name of the house'
        ]
        
        row_text = ' '.join([str(cell) if cell is not None else '' for cell in row])
        return any(pattern in row_text for pattern in house_patterns)
    
    def is_valid_data_row(self, row):
        """严格验证数据行：必须同时包含日期和金额"""
        if not any(cell for cell in row):
            return False
        row_text = ' '.join([str(cell) if cell is not None else '' for cell in row if cell])
        return self.is_valid_date(row_text) and self.is_valid_amount(row_text)
    
    def has_valid_header_structure(self, table):
        """检查表格是否有有效的表头结构"""
        if not table or len(table) < 3:
            return False
            
        # 检查前三行是否有表头
        for i in range(min(3, len(table))):
            if self.is_header_row(table[i]):
                return True
        return False
    
    def extract_tables_with_lines_strategy(self, page, page_num):
        """使用线条检测策略提取表格 - 加入表头识别，如果没有表头跳过该表格"""
        strategy_settings = {
            "vertical_strategy": "lines", 
            "horizontal_strategy": "lines",
            "explicit_vertical_lines": page.curves + page.edges,
            "explicit_horizontal_lines": page.curves + page.edges,
            "join_tolerance": 15
        }
        
        tables = page.extract_tables(strategy_settings)
        data_rows = []
        skipped_tables = 0
        
        if tables:
            for table_idx, table in enumerate(tables):
                if table and len(table) > 0:
                    # 检查是否有有效的表头结构
                    has_header = self.has_valid_header_structure(table)
                    
                    if has_header:
                        # 如果有表头，找到表头行并跳过三行
                        data_start_index = 0
                        for row_idx, row in enumerate(table):
                            if self.is_header_row(row):
                                data_start_index = min(row_idx + 3, len(table))  # 跳过表头行+额外两行
                                break
                        
                        if data_start_index == 0:
                            data_start_index = 3  # 默认跳过前三行
                        
                        # 处理表头之后的数据行
                        for row in table[data_start_index:]:
                            if any(cell is not None for cell in row):
                                # 检查是否包含屋號信息，如果是则跳过
                                if self.contains_house_number(row):
                                    continue
                                    
                                cleaned_row = [str(cell) if cell is not None else '' for cell in row]
                                processed_row = self.improved_column_separation(cleaned_row)
                                # 添加识别方法标记
                                processed_row.append("线条检测策略(有表头)")
                                data_rows.append(processed_row)
                    else:
                        # 如果没有表头，跳过整个表格
                        skipped_tables += 1
        
        return data_rows
    
    def improved_column_separation(self, row):
        """改进的数据分列处理"""
        # 首先确保所有元素都是字符串
        row = [str(cell) if cell is not None else '' for cell in row]
        
        # 如果已经是11列，直接返回
        if len(row) == 11:
            return row
        
        # 如果列数不足，尝试智能分列
        if len(row) < 11:
            # 尝试从文本中提取关键信息
            row_text = ' '.join(row)
            separated_data = self.separate_columns_by_patterns(row_text)
            
            if len(separated_data) >= 11:
                return separated_data[:11]
            elif len(separated_data) > len(row):
                # 补充空列到11列
                separated_data.extend([''] * (11 - len(separated_data)))
                return separated_data
        
        # 如果以上方法都不行，保持原样并补充空列
        if len(row) > 11:
            return row[:11]
        elif len(row) < 11:
            row.extend([''] * (11 - len(row)))
        
        return row
    
    def separate_columns_by_patterns(self, text):
        """根据模式分列数据"""
        # 定义列模式
        patterns = [
            # 日期模式 (A, B, C 列)
            r'(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{4})',
            # 金额模式 (H 列)
            r'(\$[\d,]+\.?\d*|USD\s*[\d,]+\.?\d*|HKD\s*[\d,]+\.?\d*)',
            # 建筑信息模式 (D, E, F 列)
            r'(Tower\s*\d+[A-Z]?)',
            # 楼层和单位
            r'(\d+)\s*([A-Z])',
        ]
        
        result = []
        remaining_text = text
        
        # 提取日期信息
        date_matches = re.findall(r'\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{4}', remaining_text)
        for date in date_matches[:3]:  # 最多取3个日期
            result.append(date)
            remaining_text = remaining_text.replace(date, '', 1)
        
        # 如果日期不足3个，补充空值
        while len(result) < 3:
            result.append('')
        
        # 提取建筑名称
        building_match = re.search(r'Tower\s*\d+[A-Z]?', remaining_text, re.IGNORECASE)
        if building_match:
            result.append(building_match.group(0))
            remaining_text = remaining_text.replace(building_match.group(0), '', 1)
        else:
            result.append('')
        
        # 提取楼层和单位
        unit_match = re.search(r'(\d+)\s*([A-Z])', remaining_text)
        if unit_match:
            result.append(unit_match.group(1))  # 楼层
            result.append(unit_match.group(2))  # 单位
            remaining_text = remaining_text.replace(unit_match.group(0), '', 1)
        else:
            result.extend(['', ''])  # 补充两个空值
        
        # 车位信息 (G列)
        result.append('')  # 默认为空
        
        # 提取金额 (H列)
        amount_match = re.search(r'(\$[\d,]+\.?\d*|USD\s*[\d,]+\.?\d*|HKD\s*[\d,]+\.?\d*)', remaining_text)
        if amount_match:
            result.append(amount_match.group(0))
            remaining_text = remaining_text.replace(amount_match.group(0), '', 1)
        else:
            result.append('')
        
        # 剩余文本放入I、J列
        remaining_text = remaining_text.strip()
        if remaining_text:
            # 尝试分割剩余文本
            parts = re.split(r'\s{2,}', remaining_text)
            if len(parts) >= 2:
                result.append(parts[0])  # I列
                result.append(' '.join(parts[1:]))  # J列
            else:
                result.append(remaining_text)  # I列
                result.append('')  # J列
        else:
            result.extend(['', ''])  # I、J列
        
        # K列
        result.append('')
        
        return result
    
    def extract_text_data(self, page, page_num):
        """使用文本策略提取数据 - 跳过前三行，要求第一列是日期且有金额"""
        text = page.extract_text()
        if not text:
            return []
        
        text_data = []
        lines = text.split('\n')
        
        # 跳过前三行（假设是表头）
        data_start_index = 3
        if data_start_index >= len(lines):
            return []
        
        # 处理数据行（跳过前三行）
        for line in lines[data_start_index:]:
            if not line.strip():
                continue
                
            # 检查是否包含屋號信息，如果是则跳过
            if '屋號' in line or 'House number' in line or '屋名' in line or 'Name of the house' in line:
                continue
                
            # 使用更精确的分割方法
            cells = re.split(r'\s{2,}', line)
            if cells and any(cells):
                # 检查金额列
                amount_count = sum(1 for cell in cells if self.is_valid_amount(cell))
                
                # 检查第一列是否为日期
                first_col_is_date = len(cells) > 0 and self.is_valid_date(cells[0])
                
                # 要求至少有一列金额且第一列是日期
                if amount_count >= 1 and first_col_is_date:
                    processed_row = self.improved_column_separation(cells)
                    # 添加识别方法标记
                    processed_row.append("文本提取策略")
                    text_data.append(processed_row)
        
        return text_data
    
    def extract_tables_with_table_detection(self, page, page_num):
        """使用表格检测策略提取数据 - 修正逻辑：有表头跳过三行，无表头且第一列不是日期则跳过整个表格"""
        tables_data = []
        tables = page.extract_tables()
        
        if tables:
            for table in tables:
                if table and len(table) > 0:
                    # 检查是否有有效的表头结构
                    has_header = self.has_valid_header_structure(table)
                    
                    if has_header:
                        # 如果有表头，找到表头行并跳过三行
                        data_start_index = 0
                        for row_idx, row in enumerate(table):
                            if self.is_header_row(row):
                                data_start_index = min(row_idx + 3, len(table))  # 跳过表头行+额外两行
                                break
                        
                        if data_start_index == 0:
                            data_start_index = 3  # 默认跳过前三行
                        
                        # 处理表头之后的数据行
                        for row in table[data_start_index:]:
                            if any(cell is not None for cell in row):
                                # 检查是否包含屋號信息，如果是则跳过
                                if self.contains_house_number(row):
                                    continue
                                    
                                cleaned_row = [str(cell) if cell is not None else '' for cell in row]
                                processed_row = self.improved_column_separation(cleaned_row)
                                # 添加识别方法标记
                                processed_row.append("表格检测策略(有表头)")
                                tables_data.append(processed_row)
                    else:
                        # 如果没有表头，检查表格中是否有第一列是日期的行
                        has_date_row = False
                        for row in table:
                            if any(cell is not None for cell in row) and len(row) > 0:
                                if self.is_valid_date(row[0]):
                                    has_date_row = True
                                    break
                        
                        # 如果没有日期行，跳过整个表格
                        if not has_date_row:
                            continue
                        
                        # 如果有日期行，只处理第一列是日期的行
                        for row in table:
                            if any(cell is not None for cell in row) and len(row) > 0:
                                # 检查是否包含屋號信息，如果是则跳过
                                if self.contains_house_number(row):
                                    continue
                                
                                # 检查第一列是否为日期
                                if self.is_valid_date(row[0]):
                                    cleaned_row = [str(cell) if cell is not None else '' for cell in row]
                                    processed_row = self.improved_column_separation(cleaned_row)
                                    # 添加识别方法标记
                                    processed_row.append("表格检测策略(无表头)")
                                    tables_data.append(processed_row)
        
        return tables_data
    
    def extract_tables_from_pdf(self, pdf_path, start_page=2, max_empty_pages=3):
        """从PDF中提取表格数据 - 改进版本"""
        all_tables_data = []
        empty_page_count = 0
        pages_with_tables_but_no_data = []  # 记录有表格但无数据的页码
        
        column_names = [
            '(A)临时买卖合约日期', '(B)买卖合约日期', '(C)终止买卖合约日期',
            '(D)大厦名称', '(E)楼层', '(F)单位', '(G)车位信息', '(H)成交金额',
            '(I)售价修改细节及日期', '(J)支付条款', '(K)买方是否为卖方相关人士',
            '识别方法'  # 新增列：识别方法
        ]
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                self.log_queue.put(f"处理文件: {os.path.basename(pdf_path)}, 总页数: {total_pages}")
                
                for page_num in range(start_page-1, total_pages):
                    if not self.is_running:
                        break
                        
                    current_page = page_num + 1
                    page = pdf.pages[page_num]
                    page_has_tables = False
                    page_has_data = False
                    page_data = []
                    
                    # 策略1: 表格检测（修正逻辑）
                    tables_data1 = self.extract_tables_with_table_detection(page, current_page)
                    if tables_data1:
                        page_has_tables = True
                        page_has_data = True
                        page_data.extend(tables_data1)
                    
                    # 策略2: 线条检测（加入表头识别）
                    if not page_data:
                        tables_data2 = self.extract_tables_with_lines_strategy(page, current_page)
                        if tables_data2:
                            page_has_tables = True
                            page_has_data = True
                            page_data.extend(tables_data2)
                    
                    # 策略3: 文本提取（跳过前三行，只要求金额）
                    if not page_data:
                        text_data = self.extract_text_data(page, current_page)
                        if text_data:
                            page_has_tables = True
                            page_has_data = True
                            page_data.extend(text_data)
                    
                    # 检查页面是否有表格结构但无数据
                    if page_has_tables and not page_data:
                        pages_with_tables_but_no_data.append(current_page)
                    
                    if page_data:
                        all_tables_data.extend(page_data)
                        empty_page_count = 0
                    else:
                        empty_page_count += 1
                    
                    if empty_page_count >= max_empty_pages:
                        self.log_queue.put(f"连续 {max_empty_pages} 页无数据，停止处理")
                        break
                
                # 打印有表头但无数据的页码
                if pages_with_tables_but_no_data:
                    self.log_queue.put(f"以下页码有表格但未提取到数据: {pages_with_tables_but_no_data}")
                
                # 统计各识别方法的数据量
                strategy_stats = {}
                for row in all_tables_data:
                    if len(row) >= 12:  # 确保有识别方法列
                        strategy = row[11] if len(row) > 11 else "未知方法"
                        strategy_stats[strategy] = strategy_stats.get(strategy, 0) + 1
                
                self.log_queue.put("=== 识别方法统计 ===")
                for strategy, count in strategy_stats.items():
                    self.log_queue.put(f"{strategy}: {count} 行")
                
                total_extracted = len(all_tables_data)
                self.log_queue.put(f"文件 {os.path.basename(pdf_path)} 总共提取了 {total_extracted} 行数据")
            
            # 数据清理
            cleaned_data = [row for row in all_tables_data if any(cell for cell in row[:11])]  # 只检查前11列是否为空
            self.log_queue.put(f"清理后有效数据: {len(cleaned_data)} 行")
            return cleaned_data, column_names, None
            
        except Exception as e:
            return None, None, str(e)
    
    def process_single_pdf(self, pdf_file, excel_folder, start_page=2):
        """处理单个PDF文件"""
        if not self.is_running:
            return "已停止", False
            
        try:
            tables_data, column_names, error = self.extract_tables_from_pdf(pdf_file, start_page)
            
            if error:
                return f"✗ 处理文件 {os.path.basename(pdf_file)} 时出错: {error}", False
            
            if not tables_data:
                return f"警告: {os.path.basename(pdf_file)} 中未找到有效数据", False
            
            # 创建并保存DataFrame
            df = pd.DataFrame(tables_data, columns=column_names)
            df = df.replace('', pd.NA).dropna(how='all').fillna('')
            
            pdf_name = os.path.splitext(os.path.basename(pdf_file))[0]
            excel_file = os.path.join(excel_folder, f"{pdf_name}.xlsx")
            
            with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='交易资料', index=False)
                self.set_column_widths(writer, '交易资料', df)
            
            return f"✓ 成功转换: {os.path.basename(pdf_file)} (提取了 {len(df)} 行数据)", True
            
        except Exception as e:
            return f"✗ 处理文件 {os.path.basename(pdf_file)} 时出错: {str(e)}", False
    
    def start_conversion(self):
        """开始转换"""
        if not self.pdf_folder.get() or not self.excel_folder.get():
            messagebox.showerror("错误", "请选择PDF文件夹和Excel输出文件夹")
            return
        
        self.is_running = True
        self.success_count = 0
        self.failed_count = 0
        self.total_files = 0
        self.processed_files = 0
        
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.progress_bar['value'] = 0
        self.progress_label.config(text="开始处理...")
        self.log_text.delete(1.0, tk.END)
        
        threading.Thread(target=self.batch_convert_pdf_to_excel, daemon=True).start()
    
    def stop_conversion(self):
        """停止转换"""
        self.is_running = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.progress_label.config(text="转换已停止")
        self.status_var.set("转换已停止")
    
    def conversion_completed(self):
        """转换完成"""
        self.is_running = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.progress_bar['value'] = 100
        
        completion_text = f"转换完成! 成功: {self.success_count}, 失败: {self.failed_count}"
        self.progress_label.config(text=completion_text)
        self.status_var.set(completion_text)
        messagebox.showinfo("完成", completion_text)
    
    def batch_convert_pdf_to_excel(self):
        """批量转换PDF文件到Excel"""
        pdf_folder = self.pdf_folder.get()
        excel_folder = self.excel_folder.get()
        
        if not os.path.exists(pdf_folder):
            self.log_queue.put("错误: PDF文件夹不存在")
            return
        
        Path(excel_folder).mkdir(parents=True, exist_ok=True)
        pdf_files = glob.glob(os.path.join(pdf_folder, "*.pdf"))
        
        if not pdf_files:
            self.log_queue.put("错误: 在PDF文件夹中未找到PDF文件")
            return
        
        self.total_files = len(pdf_files)
        self.processed_files = 0
        self.success_count = 0
        self.failed_count = 0
        max_workers = int(self.thread_count.get())
        
        self.log_queue.put(f"找到 {self.total_files} 个PDF文件，使用 {max_workers} 个线程处理")
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_file = {
                executor.submit(self.process_single_pdf, pdf_file, excel_folder, int(self.start_page.get())): pdf_file 
                for pdf_file in pdf_files
            }
            
            for future in as_completed(future_to_file):
                if not self.is_running:
                    break
                    
                pdf_file = future_to_file[future]
                try:
                    result, success = future.result()
                    self.log_queue.put(result)
                    self.processed_files += 1
                    
                    if success:
                        self.success_count += 1
                    else:
                        self.failed_count += 1
                    
                    self.update_progress(self.processed_files, self.total_files,
                                       f"处理中: {self.processed_files}/{self.total_files}")
                except Exception as e:
                    self.log_queue.put(f"✗ 处理文件 {os.path.basename(pdf_file)} 时发生异常: {str(e)}")
                    self.processed_files += 1
                    self.failed_count += 1
        
        self.log_queue.put("COMPLETED")

def main():
    root = tk.Tk()
    app = PDFToExcelConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()