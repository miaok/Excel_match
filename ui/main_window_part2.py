# -*- coding: utf-8 -*-
"""
主窗口模块的补充部分 - 包含ExcelMatchWindow类的其余方法
"""

# 这个文件包含了主窗口类的剩余方法，需要与main_window.py一起使用
# 在main_window.py中导入这些方法并添加到ExcelMatchWindow类中

def _addMatchField(self):
    """添加显示字段"""
    # 获取所有可用的列
    all_columns = self._getAllMatchColumns()
    
    # 创建显示字段组件
    matchField = MatchFieldWidget(self.matchFieldsContainer, all_columns)
    matchField.removeRequested.connect(self._removeMatchField)
    
    # 添加到布局
    self.matchFieldsLayout.addWidget(matchField)
    
    # 保存到显示字段列表
    self.match_fields.append(matchField)
    
    # 更新执行按钮状态
    self._updateExecuteButtonState()

def _updateExecuteButtonState(self):
    """更新执行按钮状态"""
    # 检查是否有选中的工作表
    has_selected_sheets = False
    for button in self.selected_sheets:
        if button.isChecked():
            has_selected_sheets = True
            break
    
    # 检查是否有查询字段
    has_query_fields = len(self.query_fields) > 0
    
    # 检查是否有显示字段
    has_match_fields = len(self.match_fields) > 0
    
    # 更新按钮状态
    self.executeQueryButton.setEnabled(has_selected_sheets and has_query_fields and has_match_fields)
    self.addQueryButton.setEnabled(self.data_processor.get_sheet_names())
    self.addMatchButton.setEnabled(self.data_processor.get_sheet_names())

def _removeMatchField(self, widget):
    """删除显示字段"""
    # 从列表中移除
    if widget in self.match_fields:
        self.match_fields.remove(widget)
    
    # 从布局中移除
    if widget.parentWidget():
        widget.deleteLater()
    
    # 重新排列布局
    self._reflowMatchFieldsLayout()
    
    # 更新执行按钮状态
    self._updateExecuteButtonState()

def _reflowMatchFieldsLayout(self):
    """重新排列显示字段布局"""
    # 使用FlowLayout的重新布局功能
    self.matchFieldsLayout.update()

def _reflowQueryFieldsLayout(self):
    """重新排列查询字段布局"""
    # 使用QVBoxLayout的重新布局功能
    self.queryFieldsLayout.update()

def _removeQueryField(self, widget):
    """删除查询字段"""
    # 从列表中移除
    if widget in self.query_fields:
        self.query_fields.remove(widget)
    
    # 从布局中移除
    if widget.parentWidget():
        widget.deleteLater()
    
    # 重新排列布局
    self._reflowQueryFieldsLayout()
    
    # 更新执行按钮状态
    self._updateExecuteButtonState()

def _getAllQueryColumns(self):
    """获取所有可用的查询列"""
    if not self.data_processor:
        return []
        
    # 获取所有工作表
    sheets = self.data_processor.get_all_sheets()
    if not sheets:
        return []
    
    # 获取处理模式
    processing_mode = self.processingModeCombo.currentText()
    
    # 获取所有列
    all_columns = set()
    
    # 获取选中的工作表
    selected_sheet_names = []
    for button in self.selected_sheets:
        if button.isChecked():
            selected_sheet_names.append(button.text())
    
    # 如果是合并模式，获取共同列
    if processing_mode == "合并" and len(selected_sheet_names) >= 2:
        # 获取选中工作表的数据
        selected_dfs = [sheets[name] for name in selected_sheet_names if name in sheets]
        
        # 获取共同列
        common_columns = self.data_processor.find_common_columns(selected_dfs)
        
        # 添加共同列
        all_columns.update(common_columns)
        
        # 添加工作表特定的列
        for sheet_name in selected_sheet_names:
            if sheet_name in sheets:
                sheet_df = sheets[sheet_name]
                for col in sheet_df.columns:
                    if col not in common_columns:
                        all_columns.add(f"{sheet_name}:{col}")
    else:
        # 堆叠模式，获取所有列
        for sheet_name in selected_sheet_names:
            if sheet_name in sheets:
                sheet_df = sheets[sheet_name]
                all_columns.update(sheet_df.columns)
    
    return sorted(list(all_columns))

def _getAllMatchColumns(self):
    """获取所有可用的显示列"""
    if not self.data_processor:
        return []
        
    # 获取所有工作表
    sheets = self.data_processor.get_all_sheets()
    if not sheets:
        return []
    
    # 获取处理模式
    processing_mode = self.processingModeCombo.currentText()
    
    # 获取所有列
    all_columns = set()
    
    # 获取选中的工作表
    selected_sheet_names = []
    for button in self.selected_sheets:
        if button.isChecked():
            selected_sheet_names.append(button.text())
    
    # 如果是合并模式，获取共同列和所有列
    if processing_mode == "合并" and len(selected_sheet_names) >= 2:
        # 获取选中工作表的数据
        selected_dfs = [sheets[name] for name in selected_sheet_names if name in sheets]
        
        # 获取共同列
        common_columns = self.data_processor.find_common_columns(selected_dfs)
        
        # 添加共同列
        all_columns.update(common_columns)
        
        # 添加所有列
        for sheet_name in selected_sheet_names:
            if sheet_name in sheets:
                sheet_df = sheets[sheet_name]
                for col in sheet_df.columns:
                    if col not in common_columns:
                        all_columns.add(col)
    else:
        # 堆叠模式，获取所有列
        for sheet_name in selected_sheet_names:
            if sheet_name in sheets:
                sheet_df = sheets[sheet_name]
                all_columns.update(sheet_df.columns)
    
    # 添加数据来源列
    all_columns.add("数据来源")
    
    return sorted(list(all_columns))

def _reflowSheetSelectionLayout(self):
    """重新排列工作表选择布局"""
    # 使用FlowLayout的重新布局功能
    self.sheetSelectionLayout.update()

def _showModeInfo(self):
    """显示模式信息"""
    MessageBox(
        "数据处理模式说明",
        "堆叠模式: 将不同工作表的数据垂直堆叠在一起，适用于工作表结构相似的情况。\n\n"
        "合并模式: 通过共同字段将不同工作表的数据关联合并，适用于工作表之间有关联关系的情况。",
        self
    ).exec()

def _onProcessingModeChanged(self, index):
    """处理模式变化时的处理"""
    # 更新查询字段和显示字段的可选项
    # 清空现有的查询字段和显示字段
    self._clearQueryFields()
    self._clearMatchFields()
    
    # 添加新的查询字段和显示字段
    self._addQueryField()
    self._addMatchField()
    
    # 更新执行按钮状态
    self._updateExecuteButtonState()

def onResize(self, event):
    """窗口大小变化时的处理"""
    # 调整左侧面板的大小
    if self.leftWidget and self.rightWidget:
        # 获取当前窗口大小
        window_width = self.width()
        window_height = self.height()
        
        # 调整分割器大小
        left_width = min(600, window_width * 0.4)  # 左侧面板最大宽度为600或窗口宽度的40%
        self.splitter.setSizes([left_width, window_width - left_width])
        
        # 调整左侧面板内部组件的大小
        self._adjustLeftPanelSizes(window_height)
    
    # 调用父类的resizeEvent
    super().resizeEvent(event)

def _adjustLeftPanelSizes(self, available_height):
    """调整左侧面板内部组件的大小"""
    # 计算可用高度（减去顶部文件选择区域和底部按钮区域）
    content_height = available_height - 100  # 估计值，根据实际情况调整
    
    # 分配高度给三个主要区域
    sheet_height = content_height * 0.25  # 25%给工作表选择区域
    query_height = content_height * 0.35  # 35%给查询条件区域
    match_height = content_height * 0.35  # 35%给显示字段区域
    
    # 设置最小高度
    sheet_height = max(180, sheet_height)
    query_height = max(200, query_height)
    match_height = max(200, match_height)
    
    # 应用高度
    for button in self.selected_sheets:
        button.parentWidget().setMinimumHeight(int(sheet_height))
    
    self.queryFieldsContainer.setMinimumHeight(int(query_height))
    self.matchFieldsContainer.setMinimumHeight(int(match_height))

def displayResults(self, df):
    """显示查询结果"""
    if df is None or df.empty:
        self.clearResultTable()
        return
    
    # 保存结果数据
    self.result_data = df
    
    # 清空表格
    self.resultTable.clear()
    
    # 设置表格列数和标题
    columns = df.columns.tolist()
    self.resultTable.setColumnCount(len(columns))
    self.resultTable.setHorizontalHeaderLabels(columns)
    
    # 设置表格行数
    row_count = len(df)
    self.resultTable.setRowCount(row_count)
    
    # 填充数据
    for row in range(row_count):
        for col in range(len(columns)):
            # 获取单元格值
            value = df.iloc[row, col]
            
            # 处理None和NaN值
            if value is None or (hasattr(value, 'isna') and value.isna()):
                value = ""
            else:
                value = str(value)
            
            # 创建表格项
            item = QTableWidgetItem(value)
            
            # 设置表格项
            self.resultTable.setItem(row, col, item)
    
    # 调整列宽
    self.resultTable.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
    self.resultTable.resizeColumnsToContents()
    
    # 更新结果标签
    self.resultLabel.setText(f"查询结果 ({row_count} 行)")
    
    # 显示成功消息
    InfoBar.success(
        title="查询成功",
        content=f"找到 {row_count} 条符合条件的数据",
        parent=self,
        position=InfoBarPosition.TOP,
        duration=3000
    )