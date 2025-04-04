# -*- coding: utf-8 -*-
"""
查询处理模块 - 负责处理查询条件和执行查询操作
"""

import pandas as pd
import numpy as np
from qfluentwidgets import InfoBar, InfoBarPosition, MessageBox


class QueryProcessor:
    """查询处理类，负责处理查询条件和执行查询操作"""
    
    def __init__(self, data_processor=None):
        self.data_processor = data_processor
        self.merge_how = 'outer'  # 默认合并方式为外连接
    
    def set_data_processor(self, data_processor):
        """设置数据处理器
        
        Args:
            data_processor: 数据处理器实例
        """
        self.data_processor = data_processor
    
    def set_merge_how(self, merge_how):
        """设置合并方式
        
        Args:
            merge_how: 合并方式，可选值为'outer'、'inner'、'left'
        """
        self.merge_how = merge_how
    
    def execute_stack_mode(self, selected_sheet_names, query_fields, match_fields, parent=None):
        """执行垂直堆叠模式，适用于工作表有相似结构的情况
        
        Args:
            selected_sheet_names: 选中的工作表名称列表
            query_fields: 查询字段列表
            match_fields: 显示字段列表
            parent: 父窗口，用于显示消息
            
        Returns:
            DataFrame: 查询结果
        """
        if not self.data_processor:
            return None
            
        # 存储所有工作表数据的列表，用于垂直堆叠
        all_dfs = []
        
        # 处理每个选择的工作表
        for sheet_name in selected_sheet_names:
            if not sheet_name or sheet_name not in self.data_processor.get_sheet_names():
                continue  # 跳过无效的工作表
                
            # 获取当前工作表数据
            current_df = self.data_processor.get_sheet_data(sheet_name).copy()
            
            # 跳过空数据
            if current_df.empty:
                continue
                
            # 应用查询条件（每个工作表使用相同的查询条件）
            filtered_df = self.apply_query_conditions(current_df, query_fields)
            
            # 跳过筛选后为空的数据
            if filtered_df.empty:
                continue
                
            # 添加工作表名称列，方便识别数据来源
            # 使用.loc来避免SettingWithCopyWarning
            filtered_df = filtered_df.copy()  # 创建副本以避免警告
            filtered_df.loc[:, '数据来源'] = sheet_name
            
            # 添加到结果列表
            all_dfs.append(filtered_df)
        
        # 如果没有有效数据，返回空结果
        if not all_dfs:
            if parent:
                InfoBar.warning(
                    title="查询结果为空",
                    content="没有找到符合条件的数据",
                    parent=parent,
                    position=InfoBarPosition.TOP,
                    duration=3000
                )
            return None
        
        # 垂直堆叠所有工作表数据
        try:
            # 尝试直接堆叠
            result_df = pd.concat(all_dfs, ignore_index=True)
        except Exception as e:
            # 如果直接堆叠失败，尝试对齐列后再堆叠
            try:
                aligned_dfs = self.data_processor.align_dataframe_columns(all_dfs)
                result_df = pd.concat(aligned_dfs, ignore_index=True)
            except Exception as e2:
                if parent:
                    MessageBox("堆叠错误", f"合并数据时出错: {str(e2)}", parent).exec()
                return None
        
        # 应用显示字段筛选
        if match_fields:
            result_df = self.apply_display_columns(result_df, match_fields)
        
        return result_df
    
    def execute_merge_mode(self, selected_sheet_names, query_fields, match_fields, parent=None):
        """执行合并模式，适用于工作表之间有关联关系的情况
        
        Args:
            selected_sheet_names: 选中的工作表名称列表
            query_fields: 查询字段列表
            match_fields: 显示字段列表
            parent: 父窗口，用于显示消息
            
        Returns:
            DataFrame: 查询结果
        """
        if not self.data_processor or len(selected_sheet_names) < 2:
            # 如果只有一个工作表，使用堆叠模式
            if len(selected_sheet_names) == 1:
                return self.execute_stack_mode(selected_sheet_names, query_fields, match_fields, parent)
            return None
        
        # 获取所有选中工作表的数据
        sheet_dfs = {}
        for sheet_name in selected_sheet_names:
            if sheet_name in self.data_processor.get_sheet_names():
                sheet_dfs[sheet_name] = self.data_processor.get_sheet_data(sheet_name).copy()
        
        if not sheet_dfs:
            return None
        
        # 查找所有工作表的共同列，作为可能的合并键
        common_columns = self.data_processor.find_common_columns(list(sheet_dfs.values()))
        
        # 如果没有共同列，无法执行合并，回退到堆叠模式
        if not common_columns:
            if parent:
                InfoBar.warning(
                    title="无法合并",
                    content="所选工作表没有共同列，无法执行合并操作，已切换为堆叠模式",
                    parent=parent,
                    position=InfoBarPosition.TOP,
                    duration=3000
                )
            return self.execute_stack_mode(selected_sheet_names, query_fields, match_fields, parent)
        
        # 获取合并键（这里假设已经通过UI选择了合并键）
        merge_key = common_columns[0]  # 默认使用第一个共同列作为合并键
        
        # 检查是否有工作表特定的查询条件
        sheets_with_conditions = {}
        filtered_dfs = {}
        
        # 先应用每个工作表特定的查询条件
        for sheet_name, df in sheet_dfs.items():
            # 获取该工作表特定的查询条件
            sheet_query_fields = self.get_sheet_specific_query_fields(sheet_name, query_fields)
            
            # 如果有查询条件，应用它们
            if sheet_query_fields:
                sheets_with_conditions[sheet_name] = sheet_query_fields
                filtered_df = self.apply_query_conditions(df, sheet_query_fields)
                
                # 如果筛选后不为空，添加到结果
                if not filtered_df.empty:
                    filtered_dfs[sheet_name] = filtered_df
        
        # 如果所有工作表都有特定的查询条件，使用已筛选的数据进行合并
        if len(sheets_with_conditions) == len(sheet_dfs) and filtered_dfs:
            merged_df = self.merge_filtered_sheets(filtered_dfs, sheet_dfs, sheets_with_conditions, merge_key)
        else:
            # 否则，先合并所有工作表，然后应用全局查询条件
            merged_df = self.merge_all_sheets(sheet_dfs, merge_key)
            
            # 应用全局查询条件
            all_query_fields = self.get_all_query_fields(query_fields)
            if all_query_fields:
                merged_df = self.apply_final_filtering(merged_df, all_query_fields)
        
        # 如果合并结果为空，返回None
        if merged_df is None or merged_df.empty:
            if parent:
                InfoBar.warning(
                    title="查询结果为空",
                    content="没有找到符合条件的数据",
                    parent=parent,
                    position=InfoBarPosition.TOP,
                    duration=3000
                )
            return None
        
        # 应用显示字段筛选
        if match_fields:
            merged_df = self.apply_display_columns(merged_df, match_fields)
        
        return merged_df
    
    def merge_all_sheets(self, sheet_dfs, merge_key):
        """合并所有工作表
        
        Args:
            sheet_dfs: 工作表数据字典
            merge_key: 合并键
            
        Returns:
            DataFrame: 合并后的数据
        """
        if not sheet_dfs:
            return None
            
        # 获取工作表名称列表
        sheet_names = list(sheet_dfs.keys())
        
        # 从第一个工作表开始
        merged_df = sheet_dfs[sheet_names[0]].copy()
        merged_df['数据来源'] = sheet_names[0]  # 添加来源列
        
        # 依次合并其他工作表
        for i in range(1, len(sheet_names)):
            sheet_name = sheet_names[i]
            right_df = sheet_dfs[sheet_name].copy()
            right_df['数据来源'] = sheet_name  # 添加来源列
            
            # 执行合并
            merged_df = pd.merge(
                merged_df, right_df, 
                on=merge_key,  # 使用指定的合并键
                how=self.merge_how,  # 使用指定的合并方式
                suffixes=(f'_{sheet_names[0]}', f'_{sheet_name}')  # 添加后缀以区分同名列
            )
        
        return merged_df
    
    def merge_filtered_sheets(self, filtered_dfs, sheet_dfs, sheets_with_conditions, merge_key):
        """合并已筛选的工作表
        
        Args:
            filtered_dfs: 已筛选的工作表数据字典
            sheet_dfs: 原始工作表数据字典
            sheets_with_conditions: 有查询条件的工作表字典
            merge_key: 合并键
            
        Returns:
            DataFrame: 合并后的数据
        """
        if not filtered_dfs:
            return None
            
        # 获取工作表名称列表
        sheet_names = list(filtered_dfs.keys())
        
        # 从第一个工作表开始
        merged_df = filtered_dfs[sheet_names[0]].copy()
        merged_df['数据来源'] = sheet_names[0]  # 添加来源列
        
        # 依次合并其他工作表
        for i in range(1, len(sheet_names)):
            sheet_name = sheet_names[i]
            right_df = filtered_dfs[sheet_name].copy()
            right_df['数据来源'] = sheet_name  # 添加来源列
            
            # 执行合并
            merged_df = pd.merge(
                merged_df, right_df, 
                on=merge_key,  # 使用指定的合并键
                how=self.merge_how,  # 使用指定的合并方式
                suffixes=(f'_{sheet_names[0]}', f'_{sheet_name}')  # 添加后缀以区分同名列
            )
        
        return merged_df
    
    def apply_final_filtering(self, merged_df, all_query_fields):
        """应用最终的全局筛选条件
        
        Args:
            merged_df: 合并后的数据
            all_query_fields: 全局查询条件
            
        Returns:
            DataFrame: 筛选后的数据
        """
        if merged_df is None or merged_df.empty or not all_query_fields:
            return merged_df
            
        # 创建一个全为True的初始掩码
        final_mask = pd.Series(True, index=merged_df.index)
        
        # 应用每个查询条件
        for field_tuple in all_query_fields:
            if len(field_tuple) < 3:
                continue
                
            column = field_tuple[0]
            operator = field_tuple[1]
            value = field_tuple[2]
            logic = field_tuple[3] if len(field_tuple) > 3 else "AND"
            
            # 跳过空值
            if not column or not operator or not value:
                continue
                
            # 检查列是否存在
            target_column = None
            for col in merged_df.columns:
                if col == column or col.startswith(f"{column}_"):
                    target_column = col
                    break
                    
            if target_column is None:
                continue
                
            # 应用条件
            condition_mask = self.apply_single_condition(merged_df, target_column, operator, value)
            
            # 根据逻辑运算符合并条件
            if logic == "AND":
                final_mask = final_mask & condition_mask
            elif logic == "OR":
                final_mask = final_mask | condition_mask
        
        # 应用最终掩码
        return merged_df[final_mask]
    
    def get_all_query_fields(self, query_fields):
        """获取所有查询字段
        
        Args:
            query_fields: 查询字段列表
            
        Returns:
            list: 所有查询字段列表
        """
        return query_fields
    
    def get_sheet_specific_query_fields(self, sheet_name, query_fields):
        """获取特定工作表的查询字段
        
        Args:
            sheet_name: 工作表名称
            query_fields: 查询字段列表
            
        Returns:
            list: 特定工作表的查询字段列表
        """
        sheet_specific_fields = []
        
        for field_tuple in query_fields:
            if len(field_tuple) < 3:
                continue
                
            column = field_tuple[0]
            operator = field_tuple[1]
            value = field_tuple[2]
            
            # 跳过空值
            if not column or not operator or not value:
                continue
                
            # 检查是否是特定工作表的查询条件
            if ':' in column:
                parts = column.split(':', 1)
                if parts[0] == sheet_name:
                    # 提取实际的列名
                    actual_column = parts[1]
                    sheet_specific_fields.append((actual_column, operator, value))
            else:
                # 通用查询条件，适用于所有工作表
                sheet_specific_fields.append((column, operator, value))
        
        return sheet_specific_fields
    
    def apply_query_conditions(self, df, query_fields):
        """应用查询条件
        
        Args:
            df: 数据框
            query_fields: 查询字段列表
            
        Returns:
            DataFrame: 筛选后的数据
        """
        if df is None or df.empty or not query_fields:
            return df
            
        # 创建一个全为True的初始掩码
        mask = pd.Series(True, index=df.index)
        
        # 记录每个列的条件，用于检测逻辑矛盾
        column_conditions = {}
        
        # 应用每个查询条件
        for field_tuple in query_fields:
            if len(field_tuple) < 3:
                continue
                
            column = field_tuple[0]
            operator = field_tuple[1]
            value = field_tuple[2]
            logic = field_tuple[3] if len(field_tuple) > 3 else "AND"
            
            # 跳过空值
            if not column or not operator or not value:
                continue
                
            # 检查列是否存在
            if column not in df.columns:
                continue
                
            # 记录条件
            if column not in column_conditions:
                column_conditions[column] = []
            column_conditions[column].append((operator, value, logic))
            
            # 应用条件
            condition_mask = self.apply_single_condition(df, column, operator, value)
            
            # 根据逻辑运算符合并条件
            if logic == "AND":
                mask = mask & condition_mask
            elif logic == "OR":
                mask = mask | condition_mask
        
        # 检查逻辑矛盾
        conflict_columns = self.check_logical_contradictions(column_conditions)
        if conflict_columns:
            # 有逻辑矛盾，但仍然返回结果
            pass
        
        # 应用最终掩码
        return df[mask]
    
    def apply_single_condition(self, df, column, operator, value):
        """应用单个查询条件
        
        Args:
            df: 数据框
            column: 列名
            operator: 操作符
            value: 值
            
        Returns:
            Series: 条件掩码
        """
        if df is None or df.empty or column not in df.columns:
            return pd.Series(False, index=df.index)
            
        # 获取列数据
        col_data = df[column]
        
        # 创建掩码
        mask = pd.Series(False, index=df.index)
        
        try:
            # 根据操作符应用不同的条件
            if operator == "=":
                # 等于
                mask = col_data.astype(str) == str(value)
            elif operator == "!=":
                # 不等于
                mask = col_data.astype(str) != str(value)
            elif operator == "<":
                # 小于
                # 尝试转换为数值类型
                try:
                    numeric_col = pd.to_numeric(col_data, errors='coerce')
                    numeric_val = float(value)
                    mask = numeric_col < numeric_val
                except:
                    # 如果转换失败，使用字符串比较
                    mask = col_data.astype(str) < str(value)
            elif operator == "<=":
                # 小于等于
                try:
                    numeric_col = pd.to_numeric(col_data, errors='coerce')
                    numeric_val = float(value)
                    mask = numeric_col <= numeric_val
                except:
                    mask = col_data.astype(str) <= str(value)
            elif operator == ">":
                # 大于
                try:
                    numeric_col = pd.to_numeric(col_data, errors='coerce')
                    numeric_val = float(value)
                    mask = numeric_col > numeric_val
                except:
                    mask = col_data.astype(str) > str(value)
            elif operator == ">=":
                # 大于等于
                try:
                    numeric_col = pd.to_numeric(col_data, errors='coerce')
                    numeric_val = float(value)
                    mask = numeric_col >= numeric_val
                except:
                    mask = col_data.astype(str) >= str(value)
            elif operator == "包含":
                # 包含
                mask = col_data.astype(str).str.contains(str(value), na=False)
            elif operator == "不包含":
                # 不包含
                mask = ~col_data.astype(str).str.contains(str(value), na=False)
            elif operator == "开头是":
                # 开头是
                mask = col_data.astype(str).str.startswith(str(value), na=False)
            elif operator == "结尾是":
                # 结尾是
                mask = col_data.astype(str).str.endswith(str(value), na=False)
            elif operator == "为空":
                # 为空
                mask = col_data.isna() | (col_data.astype(str) == "")
            elif operator == "不为空":
                # 不为空
                mask = ~(col_data.isna() | (col_data.astype(str) == ""))
            else:
                # 默认为等于
                mask = col_data.astype(str) == str(value)
        except Exception as e:
            # 如果出错，返回全False的掩码
            mask = pd.Series(False, index=df.index)
        
        # 处理NaN值
        mask = mask.fillna(False)
        
        return mask
    
    def check_logical_contradictions(self, column_conditions):
        """检查逻辑矛盾
        
        Args:
            column_conditions: 列条件字典
            
        Returns:
            list: 有逻辑矛盾的列列表
        """
        conflict_columns = []
        
        for column, conditions in column_conditions.items():
            # 如果一个列有多个条件
            if len(conditions) > 1:
                # 检查是否有矛盾的条件
                has_conflict = False
                
                # 检查特定的矛盾情况
                for i in range(len(conditions)):
                    op1, val1, logic1 = conditions[i]
                    
                    for j in range(i+1, len(conditions)):
                        op2, val2, logic2 = conditions[j]
                        
                        # 检查常见的矛盾情况
                        if logic1 == "AND" and logic2 == "AND":
                            # 例如: x = 1 AND x = 2
                            if op1 == "=" and op2 == "=" and val1 != val2:
                                has_conflict = True
                                break
                            # 例如: x = 1 AND x != 1
                            elif (op1 == "=" and op2 == "!=" and val1 == val2) or \
                                 (op1 == "!=" and op2 == "=" and val1 == val2):
                                has_conflict = True
                                break
                            # 例如: x < 5 AND x > 10
                            elif (op1 == "<" and op2 == ">" and float(val1) <= float(val2)) or \
                                 (op1 == ">" and op2 == "<" and float(val1) >= float(val2)):
                                has_conflict = True
                                break
                            # 例如: x <= 5 AND x >= 10
                            elif (op1 == "<=" and op2 == ">=" and float(val1) < float(val2)) or \
                                 (op1 == ">=" and op2 == "<=" and float(val1) > float(val2)):
                                has_conflict = True
                                break
                            # 例如: x 为空 AND x 不为空
                            elif (op1 == "为空" and op2 == "不为空") or \
                                 (op1 == "不为空" and op2 == "为空"):
                                has_conflict = True
                                break
                    
                    if has_conflict:
                        break
                
                if has_conflict:
                    conflict_columns.append(column)
        
        return conflict_columns
    
    def apply_display_columns(self, df, match_fields):
        """应用显示字段筛选
        
        Args:
            df: 数据框
            match_fields: 显示字段列表
            
        Returns:
            DataFrame: 筛选后的数据
        """
        if df is None or df.empty or not match_fields:
            return df
            
        # 获取要显示的列
        display_columns = []
        column_renames = {}
        
        for field_tuple in match_fields:
            if len(field_tuple) < 1:
                continue
                
            column = field_tuple[0]
            custom_title = field_tuple[1] if len(field_tuple) > 1 else None
            
            # 跳过空值
            if not column:
                continue
                
            # 检查列是否存在
            matching_columns = []
            for col in df.columns:
                if col == column or col.startswith(f"{column}_"):
                    matching_columns.append(col)
            
            # 添加匹配的列
            for col in matching_columns:
                display_columns.append(col)
                
                # 如果有自定义标题，添加到重命名字典
                if custom_title and custom_title.strip():
                    column_renames[col] = custom_title
        
        # 始终包含数据来源列
        if '数据来源' in df.columns and '数据来源' not in display_columns:
            display_columns.append('数据来源')
        
        # 如果没有指定显示列，显示所有列
        if not display_columns:
            return df
            
        # 筛选列
        result_df = df[display_columns].copy()
        
        # 重命名列
        if column_renames:
            result_df = result_df.rename(columns=column_renames)
        
        return result_df