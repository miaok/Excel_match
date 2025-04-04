# -*- coding: utf-8 -*-
"""
数据处理模块 - 负责Excel文件的读取和数据操作
"""

import os
import pandas as pd
from qfluentwidgets import InfoBar, InfoBarPosition
from PySide6.QtWidgets import QApplication


class DataProcessor:
    """数据处理类，负责Excel文件的读取和数据操作"""
    
    def __init__(self):
        self.sheets = {}
        self.excel_file = None
    
    def load_excel_file(self, file_path, parent=None):
        """加载Excel文件
        
        Args:
            file_path: Excel文件路径
            parent: 父窗口，用于显示消息
            
        Returns:
            tuple: (成功标志, 工作表名称列表或错误消息)
        """
        if not file_path:
            return False, "未选择文件"
            
        try:
            # 清空之前的数据
            self.sheets = {}
            self.excel_file = file_path
            
            # 使用pandas读取Excel文件，设置错误处理和类型检测
            try:
                # 优化: 先获取所有工作表名称
                excel = pd.ExcelFile(file_path)
                sheet_names = excel.sheet_names
                
                if not sheet_names:
                    raise ValueError("Excel文件中没有工作表")
                
                # 显示加载进度
                if parent:
                    InfoBar.info(
                        title="正在加载",
                        content=f"发现 {len(sheet_names)} 个工作表，开始读取数据...",
                        parent=parent,
                        position=InfoBarPosition.TOP,
                        duration=2000
                    )
                    QApplication.processEvents()  # 更新UI
                
                # 逐个读取工作表
                for sheet_name in sheet_names:
                    try:
                        # 尝试读取工作表，设置更多参数以提高兼容性
                        df = pd.read_excel(
                            file_path, 
                            sheet_name=sheet_name,
                            engine='openpyxl',  # 使用openpyxl引擎提高兼容性
                            na_values=['NA', 'N/A', ''],  # 处理多种空值表示
                            keep_default_na=True
                        )
                        
                        # 执行基本数据清洗
                        df = df.replace({pd.NA: None})  # 统一空值表示
                        
                        # 添加到sheets字典
                        self.sheets[sheet_name] = df
                    except Exception as sheet_error:
                        # 如果单个工作表加载失败，记录错误但继续处理其他工作表
                        if parent:
                            InfoBar.warning(
                                title="工作表加载警告",
                                content=f"工作表 '{sheet_name}' 加载失败: {str(sheet_error)}",
                                parent=parent,
                                position=InfoBarPosition.TOP,
                                duration=3000
                            )
                        continue
                
                # 检查是否成功加载了任何工作表
                if not self.sheets:
                    raise ValueError("所有工作表加载失败")
                
            except ImportError:
                # 如果openpyxl不可用，回退到xlrd
                try:
                    self.sheets = pd.read_excel(file_path, sheet_name=None, engine='xlrd')
                except Exception as e:
                    raise ValueError(f"Excel文件读取失败: {str(e)}")
            except Exception as e:
                raise ValueError(f"Excel文件读取失败: {str(e)}")
            
            # 更新界面显示工作表
            sheet_names = list(self.sheets.keys())
            
            # 显示成功消息
            if parent:
                InfoBar.success(
                    title="成功",
                    content=f"已加载Excel文件: {os.path.basename(file_path)} ({len(sheet_names)} 个工作表)",
                    parent=parent,
                    position=InfoBarPosition.TOP,
                    duration=3000
                )
                
            return True, sheet_names

        except Exception as e:
            # 提供更友好的错误提示
            error_message = f"加载Excel文件时出错: {str(e)}"
            
            if "No such file" in str(e):
                error_message = "找不到指定的Excel文件，请检查文件路径是否正确"
            elif "openpyxl" in str(e) and "not installed" in str(e):
                error_message = "缺少openpyxl库，请安装后再试: pip install openpyxl"
            elif "xlrd" in str(e) and "not installed" in str(e):
                error_message = "缺少xlrd库，请安装后再试: pip install xlrd"
            elif "Unsupported format" in str(e) or "Invalid file format" in str(e):
                error_message = "不支持的Excel文件格式，请确保文件为有效的.xlsx或.xls格式"
            elif "Permission denied" in str(e):
                error_message = "无法访问Excel文件，请检查文件是否被其他程序占用或是否有访问权限"
            
            return False, error_message
    
    def get_sheet_data(self, sheet_name):
        """获取指定工作表的数据
        
        Args:
            sheet_name: 工作表名称
            
        Returns:
            DataFrame: 工作表数据
        """
        if sheet_name in self.sheets:
            return self.sheets[sheet_name]
        return None
    
    def get_all_sheets(self):
        """获取所有工作表
        
        Returns:
            dict: 所有工作表数据
        """
        return self.sheets
    
    def get_sheet_names(self):
        """获取所有工作表名称
        
        Returns:
            list: 工作表名称列表
        """
        return list(self.sheets.keys())
    
    def find_common_columns(self, dataframes):
        """查找多个DataFrame中的共同列
        
        Args:
            dataframes: DataFrame列表
            
        Returns:
            list: 共同列名称列表
        """
        if not dataframes:
            return []
            
        # 获取第一个DataFrame的列
        common_cols = set(dataframes[0].columns)
        
        # 与其他DataFrame的列取交集
        for df in dataframes[1:]:
            common_cols = common_cols.intersection(set(df.columns))
            
        return list(common_cols)
    
    def align_dataframe_columns(self, dataframes):
        """对齐多个DataFrame的列，确保它们有相同的列结构
        
        Args:
            dataframes: DataFrame列表
            
        Returns:
            list: 对齐后的DataFrame列表
        """
        if not dataframes:
            return []
            
        # 获取所有列名的并集
        all_columns = set()
        for df in dataframes:
            all_columns.update(df.columns)
        
        # 确保所有DataFrame都有这些列
        aligned_dfs = []
        for df in dataframes:
            # 找出当前DataFrame缺少的列
            missing_columns = all_columns - set(df.columns)
            
            # 创建一个新的DataFrame，包含所有列
            aligned_df = df.copy()
            
            # 添加缺少的列，并填充为None
            for col in missing_columns:
                aligned_df[col] = None
                
            # 添加到结果列表
            aligned_dfs.append(aligned_df)
            
        return aligned_dfs