# -*- coding: utf-8 -*-
"""
主窗口模块 - 定义Excel多条件多sheet查询工具的主窗口
"""

import os
from PySide6.QtWidgets import (QApplication, QFileDialog, QHeaderView, QWidget,
                               QTableWidgetItem, QVBoxLayout, QHBoxLayout,
                               QSplitter, QLabel, QMessageBox, QDialog)
from PySide6.QtCore import Qt

from qfluentwidgets import (FluentWindow, NavigationItemPosition, FluentIcon,
                            SubtitleLabel, PrimaryPushButton, ComboBox, TableWidget,
                            MessageBox, InfoBar, InfoBarPosition, ToolButton,
                            LineEdit, SmoothScrollArea, FlowLayout, Dialog)

from ui.components import QueryFieldWidget, MatchFieldWidget, MergeKeyDialog


class ExcelMatchWindow(FluentWindow):
    """Excel多条件多sheet查询工具主窗口"""

    def __init__(self, data_processor, query_processor):
        super().__init__()
        self.setWindowTitle("Excel多条件多sheet查询")
        self.resize(1600, 800)
        self.setMinimumSize(1200, 700)  # 设置窗口最小尺寸

        # 数据处理和查询处理
        self.data_processor = data_processor
        self.query_processor = query_processor
        
        # 数据存储
        self.excel_file = None
        self.selected_sheets = []
        self.query_fields = []
        self.match_fields = []
        self.result_data = None
        
        # 界面响应式布局
        self.splitter = None
        self.leftWidget = None
        self.rightWidget = None

        # 初始化UI
        self._initUI()
        self._connectSignalToSlot()
        
        # 窗口大小变化时重新调整布局
        self.resizeEvent = self.onResize

    def _initUI(self):
        """初始化UI"""
        # 添加导航项
        self.homeInterface = QWidget(self)
        self.homeInterface.setObjectName("homeInterface")

        # 添加子界面
        self.addSubInterface(self.homeInterface, FluentIcon.HOME, "主页", position=NavigationItemPosition.TOP)
        
        # 设置主页布局
        self._initHomeInterface()

    def _initHomeInterface(self):
        """初始化主页界面，使用响应式布局"""
        # 创建主布局
        mainLayout = QVBoxLayout(self.homeInterface)
        mainLayout.setContentsMargins(10, 5, 10, 5)  # 减小上下边距
        mainLayout.setSpacing(3)  # 减小间距
        
        # 文件选择区域
        fileAreaLayout = QHBoxLayout()
        fileAreaLayout.setContentsMargins(5, 5, 5, 5)  # 减小边距
        fileLabel = SubtitleLabel("Excel文件")
        fileLabel.setContentsMargins(5, 5, 5, 5)  # 减小边距
        
        self.filePathEdit = LineEdit()
        self.filePathEdit.setReadOnly(True)
        self.filePathEdit.setPlaceholderText("请选择Excel文件")
        
        self.selectFileButton = PrimaryPushButton("选择文件")
        
        fileAreaLayout.addWidget(fileLabel)
        fileAreaLayout.addWidget(self.filePathEdit, 1)  # 1表示可伸缩
        fileAreaLayout.addWidget(self.selectFileButton)
        mainLayout.addLayout(fileAreaLayout)
        
        # 创建分割器，左侧为查询配置，右侧为结果显示
        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.setChildrenCollapsible(False)
        
        # 左侧查询配置区域
        self.leftWidget = QWidget()
        self.leftWidget.setMinimumWidth(450)  # 设置左侧区域最小宽度
        leftLayout = QVBoxLayout(self.leftWidget)
        leftLayout.setContentsMargins(0, 0, 5, 0)
        leftLayout.setSpacing(2)  # 保持统一的间距
        
        # 创建一个垂直滚动区域，包含所有左侧组件
        leftScrollArea = SmoothScrollArea()
        leftScrollArea.setWidgetResizable(True)
        leftScrollArea.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        leftScrollArea.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        leftScrollArea.setMinimumHeight(600)  # 设置滚动区域最小高度
        
        leftScrollContent = QWidget()
        leftScrollLayout = QVBoxLayout(leftScrollContent)
        leftScrollLayout.setContentsMargins(5, 5, 5, 5)
        leftScrollLayout.setSpacing(10)  # 区域之间的间距
        
        # ========== 1. 工作表选择区域 ==========
        sheetSelectionSection = QWidget()
        sheetSelectionSection.setMinimumHeight(180)  # 设置最小高度
        sheetSelectionLayout = QVBoxLayout(sheetSelectionSection)
        sheetSelectionLayout.setContentsMargins(0, 0, 0, 0)
        sheetSelectionLayout.setSpacing(5)
        
        # 工作表选择标题和按钮区域
        sheetTitleLayout = QHBoxLayout()
        sheetTitleLayout.setContentsMargins(5, 5, 5, 5)
        sheetSelectionLabel = SubtitleLabel("设置查询工作表")
        sheetSelectionLabel.setContentsMargins(5, 5, 5, 5)
        sheetTitleLayout.addWidget(sheetSelectionLabel)
        sheetTitleLayout.addStretch(1)
        sheetSelectionLayout.addLayout(sheetTitleLayout)
        
        # 工作表选择内容区域
        self.sheetSelectionContainer = QWidget()
        self.sheetSelectionContainer.setStyleSheet("""
            QWidget {
                background-color: #f8f8f8;
                border: 1px solid #ddd;
                border-radius: 4px;
            }
        """)
        self.sheetSelectionLayout = FlowLayout(self.sheetSelectionContainer)
        self.sheetSelectionLayout.setContentsMargins(5, 10, 5, 10)
        self.sheetSelectionLayout.setHorizontalSpacing(6)
        self.sheetSelectionLayout.setVerticalSpacing(2)
        self.sheetSelectionLayout.setAlignment(Qt.AlignTop)  # 内容靠上对齐
        
        sheetSelectionLayout.addWidget(self.sheetSelectionContainer, 1)  # 1表示可伸缩
        
        # 添加数据处理模式选择
        modeSelectionLayout = QHBoxLayout()
        modeSelectionLayout.setContentsMargins(5, 5, 5, 5)
        modeSelectionLabel = QLabel("数据处理模式:")
        self.processingModeCombo = ComboBox()
        self.processingModeCombo.addItems(["堆叠", "合并"])
        self.processingModeCombo.setToolTip("堆叠: 适用于工作表有相似结构\n合并: 适用于工作表之间有关联关系")
        self.processingModeCombo.setMinimumWidth(120)
        
        # 当处理模式变化时，更新查询和显示字段的可选项
        self.processingModeCombo.currentIndexChanged.connect(self._onProcessingModeChanged)
        
        # 模式说明
        modeInfoButton = ToolButton(FluentIcon.HELP)
        modeInfoButton.setToolTip("堆叠模式: 将不同工作表数据垂直组合\n合并模式: 通过共同字段将不同工作表数据关联合并")
        modeInfoButton.clicked.connect(self._showModeInfo)
        
        modeSelectionLayout.addWidget(modeSelectionLabel)
        modeSelectionLayout.addWidget(self.processingModeCombo)
        modeSelectionLayout.addWidget(modeInfoButton)
        modeSelectionLayout.addStretch(1)
        sheetSelectionLayout.addLayout(modeSelectionLayout)
        
        leftScrollLayout.addWidget(sheetSelectionSection, 1)  # 1表示可伸缩
        
        # ========== 2. 查询条件区域 ==========
        queryConditionSection = QWidget()
        queryConditionSection.setMinimumHeight(200)  # 设置最小高度
        queryConditionLayout = QVBoxLayout(queryConditionSection)
        queryConditionLayout.setContentsMargins(0, 0, 0, 0)
        queryConditionLayout.setSpacing(5)
        
        # 查询条件标题和按钮
        queryTitleLayout = QHBoxLayout()
        queryTitleLayout.setContentsMargins(5, 5, 5, 5)
        queryConditionLabel = SubtitleLabel("设置查询条件")
        queryConditionLabel.setContentsMargins(5, 5, 5, 5)
        queryTitleLayout.addWidget(queryConditionLabel)
        self.addQueryButton = ToolButton(FluentIcon.ADD)
        self.addQueryButton.setToolTip("添加查询条件列标题")
        self.addQueryButton.setEnabled(False)
        queryTitleLayout.addWidget(self.addQueryButton)
        queryTitleLayout.addStretch(1)
        queryConditionLayout.addLayout(queryTitleLayout)
        
        # 查询条件内容区域
        self.queryFieldsContainer = QWidget()
        self.queryFieldsContainer.setStyleSheet("""
            QWidget {
                background-color: #f9f9f9;
                border: 1px solid #ddd;
                border-radius: 4px;
            }
        """)
        self.queryFieldsContainer.setMinimumHeight(150)  # 设置最小高度
        self.queryFieldsLayout = QVBoxLayout(self.queryFieldsContainer)
        self.queryFieldsLayout.setContentsMargins(5, 5, 5, 5)
        self.queryFieldsLayout.setSpacing(3)
        self.queryFieldsLayout.setAlignment(Qt.AlignTop)  # 内容靠上对齐
        
        queryConditionLayout.addWidget(self.queryFieldsContainer, 1)  # 1表示可伸缩
        
        leftScrollLayout.addWidget(queryConditionSection, 1)  # 1表示可伸缩
        
        # ========== 3. 显示字段区域 ==========
        displayFieldsSection = QWidget()
        displayFieldsSection.setMinimumHeight(200)  # 设置最小高度
        displayFieldsLayout = QVBoxLayout(displayFieldsSection)
        displayFieldsLayout.setContentsMargins(0, 0, 0, 0)
        displayFieldsLayout.setSpacing(5)
        
        # 显示字段标题和按钮
        displayTitleLayout = QHBoxLayout()
        displayTitleLayout.setContentsMargins(5, 5, 5, 5)
        displayFieldsLabel = SubtitleLabel("设置结果列标题")
        displayFieldsLabel.setContentsMargins(5, 5, 5, 5)
        displayTitleLayout.addWidget(displayFieldsLabel)
        self.addMatchButton = ToolButton(FluentIcon.ADD)
        self.addMatchButton.setToolTip("添加结果要显示的列标题")
        self.addMatchButton.setEnabled(False)
        displayTitleLayout.addWidget(self.addMatchButton)
        displayTitleLayout.addStretch(1)
        displayFieldsLayout.addLayout(displayTitleLayout)
        
        # 显示字段内容区域
        self.matchFieldsContainer = QWidget()
        self.matchFieldsContainer.setStyleSheet("""
            QWidget {
                background-color: #f8f8f8;
                border: 1px solid #ddd;
                border-radius: 4px;
            }
        """)
        self.matchFieldsContainer.setMinimumHeight(150)  # 设置最小高度
        self.matchFieldsLayout = FlowLayout(self.matchFieldsContainer)
        self.matchFieldsLayout.setContentsMargins(5, 5, 5, 5)
        self.matchFieldsLayout.setHorizontalSpacing(6)
        self.matchFieldsLayout.setVerticalSpacing(2)
        self.matchFieldsLayout.setAlignment(Qt.AlignTop)  # 内容靠上对齐
        
        displayFieldsLayout.addWidget(self.matchFieldsContainer, 1)  # 1表示可伸缩
        
        leftScrollLayout.addWidget(displayFieldsSection, 1)  # 1表示可伸缩
        
        # 执行查询按钮
        executeLayout = QHBoxLayout()
        executeLayout.setContentsMargins(5, 10, 5, 5)
        self.executeQueryButton = PrimaryPushButton("开始查询")
        self.executeQueryButton.setIcon(FluentIcon.SEARCH)
        self.executeQueryButton.setEnabled(False)
        executeLayout.addWidget(self.executeQueryButton)
        leftScrollLayout.addLayout(executeLayout)
        
        # 设置滚动区域内容
        leftScrollArea.setWidget(leftScrollContent)
        leftLayout.addWidget(leftScrollArea)
        
        # 右侧结果区域
        self.rightWidget = QWidget()
        rightLayout = QVBoxLayout(self.rightWidget)
        rightLayout.setContentsMargins(5, 0, 0, 0)
        
        # 结果标题
        resultTitleLayout = QHBoxLayout()
        self.resultLabel = SubtitleLabel("查询结果")
        resultTitleLayout.addWidget(self.resultLabel)
        resultTitleLayout.addStretch(1)
        rightLayout.addLayout(resultTitleLayout)
        
        # 结果表格
        self.resultTable = TableWidget()
        self.resultTable.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.resultTable.setSortingEnabled(True)
        self.resultTable.setAlternatingRowColors(True)
        rightLayout.addWidget(self.resultTable, 1)  # 1表示可伸缩
        
        # 添加左右两侧部件到分割器
        self.splitter.addWidget(self.leftWidget)
        self.splitter.addWidget(self.rightWidget)
        self.splitter.setSizes([600, 800])  # 设置初始大小比例
        
        # 将分割器添加到主布局
        mainLayout.addWidget(self.splitter, 1)  # 1表示可伸缩
        
        # 初始化数据
        self.query_fields = []  # 查询字段列表
        self.match_fields = []  # 显示字段列表

    def _connectSignalToSlot(self):
        """连接信号和槽"""
        self.selectFileButton.clicked.connect(self.selectExcelFile)
        self.addQueryButton.clicked.connect(self._addQueryField)
        self.addMatchButton.clicked.connect(self._addMatchField)
        self.executeQueryButton.clicked.connect(self.executeMultiSheetQuery)

    def selectExcelFile(self):
        """选择Excel文件"""
        filePath, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", "Excel Files (*.xlsx *.xls)")

        if not filePath:
            return

        # 显示加载中状态
        self.filePathEdit.setText("正在加载...")
        QApplication.processEvents()  # 确保UI更新

        # 清空之前的数据
        self.clearResultTable()
        
        # 清空已选择的工作表
        self._clearSheetSelections()
        
        # 清空查询字段和显示字段
        self._clearAllFields()

        # 使用数据处理器加载Excel文件
        success, result = self.data_processor.load_excel_file(filePath, self)
        
        if success:
            # 加载成功，result是工作表名称列表
            sheet_names = result
            
            # 添加所有工作表按钮
            if sheet_names:
                # 创建所有工作表的TogglePushButton
                for sheet_name in sheet_names:
                    self._addSheetToggleButton(sheet_name)
                
                # 自动添加一个查询条件和一个显示字段
                self._addQueryField()
                self._addMatchField()
            
            # 更新字段按钮状态
            self._updateExecuteButtonState()

            # 更新文件路径显示
            self.filePathEdit.setText(filePath)
            self.excel_file = filePath
        else:
            # 加载失败，result是错误消息
            self.filePathEdit.setText("")
            MessageBox("错误", result, self).exec()

    def _addSheetToggleButton(self, sheet_name):
        """添加工作表切换按钮"""
        if not sheet_name:
            return
            
        # 创建TogglePushButton
        toggleButton = TogglePushButton(sheet_name)
        toggleButton.setCheckable(True)
        toggleButton.setChecked(True)  # 默认选中
        toggleButton.toggled.connect(lambda checked: self._onSheetToggled(sheet_name, checked))
        
        # 设置按钮样式 - 使按钮更紧凑
        toggleButton.setMinimumWidth(80)
        toggleButton.setMaximumWidth(150)
        toggleButton.setMinimumHeight(28)
        toggleButton.setMaximumHeight(28)
        
        # 添加到布局
        self.sheetSelectionLayout.addWidget(toggleButton)
        
        # 保存到已选择的工作表集合
        self.selected_sheets.append(toggleButton)
        
        # 添加后立即更新布局
        self._reflowSheetSelectionLayout()
        
        # 更新执行按钮状态
        self._updateExecuteButtonState()
    
    def _onSheetToggled(self, sheet_name, checked):
        """工作表选择状态改变时的处理"""
        # 更新执行按钮状态
        self._updateExecuteButtonState()
        
    def _clearSheetSelections(self):
        """清空所有工作表选择"""
        # 清空已选择的工作表
        for button in self.selected_sheets:
            if button.parentWidget():
                button.deleteLater()
        self.selected_sheets = []
        
        # 重新排列布局
        self._reflowSheetSelectionLayout()

    def _clearAllFields(self):
        """清空所有字段（查询字段和匹配字段）"""
        self._clearQueryFields()
        self._clearMatchFields()

    def _clearQueryFields(self):
        """清空所有查询字段"""
        # 清空查询字段
        for widget in self.query_fields:
            if widget.parentWidget():
                widget.deleteLater()
        self.query_fields = []
        
    def _clearMatchFields(self):
        """清空所有显示字段"""
        # 清空显示字段
        for widget in self.match_fields:
            if widget.parentWidget():
                widget.deleteLater()
        self.match_fields = []

    def executeMultiSheetQuery(self):
        """执行多工作表查询，可选择合并或堆叠不同工作表的数据"""
        try:
            # 检查是否有选中的工作表
            selected_sheet_names = []
            for button in self.selected_sheets:
                if button.isChecked():
                    selected_sheet_names.append(button.text())
            
            if not selected_sheet_names:
                MessageBox(
                    "无法执行查询", 
                    "请先选择至少一个工作表", 
                    self
                ).exec()
                return
                
            # 工作表间关系处理方式
            # 从下拉框获取当前选择的处理模式
            processing_mode = self.processingModeCombo.currentText()
            
            # 获取查询条件
            query_conditions = []
            for widget in self.query_fields:
                query_conditions.append(widget.getQueryCondition())
                
            # 获取显示字段
            match_fields = []
            for widget in self.match_fields:
                match_fields.append(widget.getMatchField())
            
            # 检查是否有查询条件
            has_query_conditions = False
            for condition in query_conditions:
                if condition[2].strip():  # 检查值是否非空
                    has_query_conditions = True
                    break
                    
            if not has_query_conditions:
                # 告诉用户没有设置查询条件，但仍会执行
                InfoBar.info(
                    title="查询提示",
                    content="未设置查询条件，将返回所有数据",
                    parent=self,
                    position=InfoBarPosition.TOP,
                    duration=3000
                )
            
            # 执行对应模式的查询
            if processing_mode == "堆叠":
                # 垂直堆叠模式 - 适用于工作表有相似结构的情况
                result_df = self.query_processor.execute_stack_mode(
                    selected_sheet_names, query_conditions, match_fields, self)
            elif processing_mode == "合并" and len(selected_sheet_names) >= 2:
                # 合并模式 - 适用于不同工作表之间有关联关系的情况
                # 获取所有选中工作表的数据
                sheet_dfs = {}
                for sheet_name in selected_sheet_names:
                    sheet_dfs[sheet_name] = self.data_processor.get_sheet_data(sheet_name)
                
                # 查找所有工作表的共同列，作为可能的合并键
                common_columns = self.data_processor.find_common_columns(list(sheet_dfs.values()))
                
                # 如果没有共同列，无法执行合并，回退到堆叠模式
                if not common_columns:
                    InfoBar.warning(
                        title="无法合并",
                        content="所选工作表没有共同列，无法执行合并操作，已切换为堆叠模式",
                        parent=self,
                        position=InfoBarPosition.TOP,
                        duration=3000
                    )
                    result_df = self.query_processor.execute_stack_mode(
                        selected_sheet_names, query_conditions, match_fields, self)
                else:
                    # 显示合并键选择对话框
                    merge_key = self._showMergeKeySelectionDialog(common_columns)
                    if merge_key:
                        # 设置合并方式
                        self.query_processor.set_merge_how(self.merge_how)
                        # 执行合并查询
                        result_df = self.query_processor.execute_merge_mode(
                            selected_sheet_names, query_conditions, match_fields, self)
                    else:
                        # 用户取消了合并键选择，回退到堆叠模式
                        InfoBar.info(
                            title="模式调整",
                            content="未选择合并键，已自动切换为堆叠模式",
                            parent=self,
                            position=InfoBarPosition.TOP,
                            duration=3000
                        )
                        result_df = self.query_processor.execute_stack_mode(
                            selected_sheet_names, query_conditions, match_fields, self)
            else:
                # 如果是合并模式但只选择了一个工作表，提示用户并使用堆叠模式
                if processing_mode == "合并" and len(selected_sheet_names) == 1:
                    InfoBar.info(
                        title="模式调整",
                        content="合并模式需要至少两个工作表，已自动切换为堆叠模式",
                        parent=self,
                        position=InfoBarPosition.TOP,
                        duration=3000
                    )
                # 执行堆叠模式
                result_df = self.query_processor.execute_stack_mode(
                    selected_sheet_names, query_conditions, match_fields, self)
            
            # 显示结果
            if result_df is not None and not result_df.empty:
                self.displayResults(result_df)
            else:
                self.clearResultTable()
                InfoBar.warning(
                    title="查询结果为空",
                    content="没有找到符合条件的数据",
                    parent=self,
                    position=InfoBarPosition.TOP,
                    duration=3000
                )
            
        except KeyError as e:
            MessageBox("查询错误", f"列名错误: {str(e)}", self).exec()
            self.clearResultTable()
        except ValueError as e:
            MessageBox("查询错误", f"值错误: {str(e)}", self).exec()
            self.clearResultTable()
        except Exception as e:
            MessageBox("错误", f"执行查询时发生意外错误: {str(e)}", self).exec()
            self.clearResultTable()
    
    def _showMergeKeySelectionDialog(self, common_columns):
        """显示合并键选择对话框"""
        if not common_columns:
            return None
            
        # 创建对话框
        dialog = Dialog("选择合并键", self)
        dialog.setMinimumWidth(400)
        
        # 创建合并键选择组件
        mergeKeyWidget = MergeKeyDialog(dialog, common_columns)
        dialog.setContentWidget(mergeKeyWidget)
        
        # 连接确认按钮
        mergeKeyWidget.confirmButton.clicked.connect(dialog.accept)
        
        # 显示对话框
        if dialog.exec():
            # 获取选择的合并键和合并方式
            self.merge_how = mergeKeyWidget.getMergeHow()
            return mergeKeyWidget.getSelectedKey()
        
        return None
    
    def clearResultTable(self):
        """清空结果表格"""
        self.resultTable.clear()
        self.resultTable.setRowCount(0)
        self.resultTable.setColumnCount(0)
        self.result_data = None
    
    def _addQueryField(self):
        """添加查询字段"""
        # 获取所有可用的列
        all_columns = self._getAllQueryColumns()
        
        # 创建查询字段组件
        queryField = QueryFieldWidget(self.queryFieldsContainer, all_columns)
        queryField.removeRequested.connect(self._removeQueryField)
        
        # 添加到布局
        self.queryFieldsLayout.addWidget(queryField)
        
        # 保存到查询字段列表
        self.query_fields.append(queryField)
        
        # 更新执行按钮状态
        self._updateExecuteButtonState()
    
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