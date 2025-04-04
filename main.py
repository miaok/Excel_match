# --- START OF FILE main.py ---

import sys
import os
import pandas as pd
from PySide6.QtWidgets import (QApplication, QFileDialog, QHeaderView, QWidget,
                               QTableWidgetItem, QVBoxLayout, QHBoxLayout, QGridLayout,
                               QSplitter, QScrollArea, QFrame, QMainWindow, QPushButton,
                               QComboBox, QTableWidget, QLabel, QLineEdit, QToolButton,
                               QMessageBox, QGroupBox, QFormLayout)
from PySide6.QtCore import Qt, Signal, QSize, QRect, QMargins
from PySide6.QtGui import QIcon, QFont

from qfluentwidgets import (Dialog, FluentWindow, NavigationItemPosition, SplashScreen,
                            FluentIcon, setTheme, Theme, SubtitleLabel, PushButton,
                            ComboBox, TableWidget, MessageBox, InfoBar, InfoBarPosition,
                            ToolButton, FluentStyleSheet, LineEdit, SmoothScrollArea, FlowLayout, 
                            Flyout, PrimaryPushButton, PushButton, TogglePushButton)

class ExcelMatchWindow(FluentWindow):
    """Excel多条件多sheet查询工具主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel多条件多sheet查询")
        self.resize(1600, 800)
        self.setMinimumSize(1200, 700)  # 设置窗口最小尺寸

        # 数据存储
        self.excel_file = None
        self.sheets = {}
        self.selected_sheets = []
        self.query_fields = []
        self.match_fields = []
        self.result_data = None
        self.merge_how = 'outer'  # 默认合并方式为外连接
        
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
        self.query_fields = []  # 查询字段列表，元组 (ComboBox, LineEdit)
        self.match_fields = []  # 显示字段列表，元组 (ComboBox, LineEdit)，LineEdit用于自定义标题

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

        try:
            self.filePathEdit.setText("正在加载...")
            QApplication.processEvents()  # 确保UI更新

            # 清空之前的数据
            self.sheets = {}
            self.clearResultTable()
            
            # 清空已选择的工作表
            self._clearSheetSelections()
            
            # 清空查询字段和显示字段
            self._clearAllFields()

            # 使用pandas读取Excel文件，设置错误处理和类型检测
            try:
                # 优化: 先获取所有工作表名称
                excel = pd.ExcelFile(filePath)
                sheet_names = excel.sheet_names
                
                if not sheet_names:
                    raise ValueError("Excel文件中没有工作表")
                
                # 显示加载进度
                InfoBar.info(
                    title="正在加载",
                    content=f"发现 {len(sheet_names)} 个工作表，开始读取数据...",
                    parent=self,
                    position=InfoBarPosition.TOP,
                    duration=2000
                )
                QApplication.processEvents()  # 更新UI
                
                # 逐个读取工作表
                for sheet_name in sheet_names:
                    try:
                        # 尝试读取工作表，设置更多参数以提高兼容性
                        df = pd.read_excel(
                            filePath, 
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
                        InfoBar.warning(
                            title="工作表加载警告",
                            content=f"工作表 '{sheet_name}' 加载失败: {str(sheet_error)}",
                            parent=self,
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
                    self.sheets = pd.read_excel(filePath, sheet_name=None, engine='xlrd')
                except Exception as e:
                    raise ValueError(f"Excel文件读取失败: {str(e)}")
            except Exception as e:
                raise ValueError(f"Excel文件读取失败: {str(e)}")
            
            # 更新界面显示工作表
            sheet_names = list(self.sheets.keys())
            
            # 添加所有工作表按钮
            if sheet_names:
                # 创建所有工作表的TogglePushButton
                for sheet_name in sheet_names:
                    self._addSheetToggleButton(sheet_name)
                
                # 自动添加一个查询条件和一个显示字段
                self._addQueryField()
                self._addMatchField()
            else:
                # 这种情况不应该发生，因为前面已经检查过
                raise ValueError("没有找到有效的工作表")
            
            # 更新字段按钮状态
            self._updateExecuteButtonState()

            # 更新文件路径显示
            self.filePathEdit.setText(filePath)

            # 显示成功消息
            InfoBar.success(
                title="成功",
                content=f"已加载Excel文件: {os.path.basename(filePath)} ({len(sheet_names)} 个工作表)",
                parent=self,
                position=InfoBarPosition.TOP,
                duration=3000
            )

        except Exception as e:
            # 清空文件路径
            self.filePathEdit.setText("")
            
            # 显示详细的错误信息
            error_message = f"加载Excel文件时出错: {str(e)}"
            
            # 提供更友好的错误提示
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
            
            # 显示错误对话框
            MessageBox("错误", error_message, self).exec()
            
            # 打印异常堆栈跟踪，方便调试
            import traceback
            traceback.print_exc()

    def _addSheetToggleButton(self, sheet_name):
        """添加工作表切换按钮"""
        if not self.sheets or not sheet_name:
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
        for field_tuple in self.query_fields:
            # 字段元组现在可能是(列选择框, 操作符选择框, 值输入框, 逻辑选择框)
            if len(field_tuple) > 0 and field_tuple[0].parentWidget():
                field_tuple[0].parentWidget().deleteLater()
        self.query_fields = []
        
    def _clearMatchFields(self):
        """清空所有显示字段"""
        # 清空显示字段
        for field_tuple in self.match_fields:
            # 字段元组现在可能是(列选择框, 自定义标题框)
            if len(field_tuple) > 0 and field_tuple[0].parentWidget():
                field_tuple[0].parentWidget().deleteLater()
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
            
            # 检查是否有查询条件
            has_query_conditions = False
            for field_tuple in self.query_fields:
                if len(field_tuple) >= 3 and field_tuple[2].text().strip():
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
                self._executeStackMode(selected_sheet_names)
            elif processing_mode == "合并" and len(selected_sheet_names) >= 2:
                # 合并模式 - 适用于不同工作表之间有关联关系的情况
                self._executeMergeMode(selected_sheet_names)
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
                self._executeStackMode(selected_sheet_names)
            
        except KeyError as e:
            MessageBox("查询错误", f"列名错误: {str(e)}", self).exec()
            self.clearResultTable()
        except ValueError as e:
            MessageBox("查询错误", f"值错误: {str(e)}", self).exec()
            self.clearResultTable()
        except Exception as e:
            MessageBox("错误", f"执行查询时发生意外错误: {str(e)}", self).exec()
            self.clearResultTable()
            
    def _executeStackMode(self, selected_sheet_names):
        """执行垂直堆叠模式，适用于工作表有相似结构的情况"""
        # 存储所有工作表数据的列表，用于垂直堆叠
        all_dfs = []
        
        # 处理每个选择的工作表
        for sheet_name in selected_sheet_names:
            if not sheet_name or sheet_name not in self.sheets:
                continue  # 跳过无效的工作表
                
            # 获取当前工作表数据
            current_df = self.sheets[sheet_name].copy()
            
            # 跳过空数据
            if current_df.empty:
                continue
                
            # 应用查询条件（每个工作表使用相同的查询条件）
            filtered_df = self._applyQueryConditions(current_df, self.query_fields)
            
            # 跳过筛选后为空的数据
            if filtered_df.empty:
                continue
                
            # 添加工作表名称列，方便识别数据来源
            # 使用.loc来避免SettingWithCopyWarning
            filtered_df = filtered_df.copy()  # 创建副本以避免警告
            filtered_df.loc[:, '数据来源'] = sheet_name
            
            # 将筛选后的数据添加到列表
            all_dfs.append(filtered_df)
        
        # 如果没有有效数据，显示提示
        if not all_dfs:
            # 使用MessageBox替代InfoBar
            MessageBox(
                "查询结果", 
                "未找到匹配记录，请检查查询条件或选择其他工作表。", 
                self
            ).exec()
            self.clearResultTable()
            return
            
        # 垂直堆叠所有数据（类似VSTACK功能）
        try:
            # 使用列对齐方法确保所有DataFrame具有相同的列结构
            aligned_dfs = self._alignDataFrameColumns(all_dfs)
            
            # 垂直堆叠对齐后的DataFrame
            stacked_df = pd.concat(aligned_dfs, ignore_index=True)
        except Exception as e:
            raise ValueError(f"无法垂直堆叠数据: {str(e)}")
            
        # 筛选显示列
        self._processAndDisplayResults(stacked_df)
            
    def _executeMergeMode(self, selected_sheet_names):
        """执行合并模式，通过关联列合并不同的工作表"""
        if len(selected_sheet_names) < 2:
            # 如果只有一个工作表，则转为堆叠模式处理
            self._executeStackMode(selected_sheet_names)
            return
            
        try:
            # 获取所有选中的工作表数据
            sheet_dfs = {}
            for sheet_name in selected_sheet_names:
                if sheet_name in self.sheets and not self.sheets[sheet_name].empty:
                    # 获取工作表数据副本
                    sheet_dfs[sheet_name] = self.sheets[sheet_name].copy()
            
            if not sheet_dfs:
                # 使用MessageBox替代InfoBar
                MessageBox(
                    "查询结果", 
                    "未找到有效工作表数据，请检查所选工作表。", 
                    self
                ).exec()
                self.clearResultTable()
                return
            
            # 保存所有列信息，用于后续更新查询和显示字段
            self.all_merge_columns = {}
            for sheet_name, df in sheet_dfs.items():
                for col in df.columns:
                    # 构造带工作表名的完整列名，例如"工作表1.列名"
                    full_col_name = f"{sheet_name}.{col}"
                    self.all_merge_columns[full_col_name] = (sheet_name, col)
                
            # 找出工作表间的共同列，可能用于关联
            common_columns = self._findCommonColumns(list(sheet_dfs.values()))
            
            if not common_columns:
                # 如果没有共同列，提示用户并回退到堆叠模式
                InfoBar.warning(
                    title="无法执行合并",
                    content="所选工作表没有共同列可供合并，将使用堆叠模式",
                    parent=self,
                    position=InfoBarPosition.TOP,
                    duration=3000
                )
                self._executeStackMode(selected_sheet_names)
                return
                
            # 如果有多个共同列，让用户选择合并键
            merge_key = None
            if len(common_columns) > 1:
                # 使用对话框让用户选择合并键
                merge_key = self._showMergeKeySelectionDialog(common_columns)
                if not merge_key:  # 用户取消选择
                    InfoBar.warning(
                        title="合并取消",
                        content="未选择合并键，将使用堆叠模式",
                        parent=self,
                        position=InfoBarPosition.TOP,
                        duration=3000
                    )
                    self._executeStackMode(selected_sheet_names)
                    return
            else:
                # 只有一个共同列，直接使用
                merge_key = common_columns[0]
                InfoBar.info(
                    title="合并信息",
                    content=f"使用唯一共同列 '{merge_key}' 作为合并键",
                    parent=self,
                    position=InfoBarPosition.TOP,
                    duration=3000
                )
                
            # 获取所有查询条件
            all_query_fields = self._getAllQueryFields()
            
            # 如果没有查询条件，则按正常方式处理
            if not all_query_fields:
                InfoBar.info(
                    title="合并情况",
                    content="没有设置查询条件，将合并所有数据",
                    parent=self,
                    position=InfoBarPosition.TOP,
                    duration=3000
                )
                merged_df = self._mergeAllSheets(sheet_dfs, merge_key)
                if merged_df is None or merged_df.empty:
                    # 使用MessageBox替代InfoBar
                    MessageBox(
                        "查询结果", 
                        "合并后无有效数据，请检查工作表数据或合并键。", 
                        self
                    ).exec()
                    self.clearResultTable()
                    return
                    
                merged_df['数据来源'] = '合并数据'
                self._processAndDisplayResults(merged_df)
                return
            
            # 创建所有工作表的查询过滤结果
            filtered_dfs = {}
            sheets_with_conditions = set()
            all_condition_errors = []  # 收集所有条件错误信息
            
            # 获取每个工作表对应的查询条件
            for sheet_name, df in sheet_dfs.items():
                sheet_query_fields = self._getSheetSpecificQueryFields(sheet_name)
                
                # 如果有对应的查询条件，记录并执行查询
                if sheet_query_fields:
                    sheets_with_conditions.add(sheet_name)
                    try:
                        # 检查是否有条件不满足
                        pre_filtered_df = df.copy()
                        for field in sheet_query_fields:
                            if len(field) >= 3 and field[2].text().strip():
                                column = field[0].currentText()
                                operator = field[1].currentText()
                                value = field[2].text().strip()
                                
                                # 应用单个条件检查
                                temp_mask = self._applySingleCondition(pre_filtered_df, column, operator, value)
                                if not temp_mask.any():
                                    all_condition_errors.append(f"工作表 '{sheet_name}' 的条件 '{column} {operator} {value}' 没有匹配数据")
                        
                        # 如果没有错误，才应用完整的查询条件
                        if not all_condition_errors:
                            filtered_df = self._applyQueryConditions(df, sheet_query_fields)
                            
                            # 如果过滤后有数据，添加标识并保存
                            if not filtered_df.empty:
                                filtered_df = filtered_df.copy()
                                filtered_df.loc[:, f'{sheet_name}_数据来源'] = True
                                filtered_dfs[sheet_name] = filtered_df
                    except Exception as e:
                        all_condition_errors.append(f"工作表 '{sheet_name}' 查询出错: {str(e)}")
            
            # 如果有任何条件错误，立即停止并显示
            if all_condition_errors:
                message = "查询条件错误:\n\n"
                for idx, err in enumerate(all_condition_errors, 1):
                    message += f"{idx}. {err}\n"
                message += "\n请检查查询条件。"
                
                MessageBox("查询结果", message, self).exec()
                self.clearResultTable()
                return
            
            # 检查是否有任何工作表有查询条件
            if not sheets_with_conditions:
                # 没有查询条件，直接合并所有工作表
                merged_df = self._mergeAllSheets(sheet_dfs, merge_key)
            else:
                # 如果有查询条件，需要首先将有条件的工作表进行过滤
                # 查看是否有任何满足条件的数据
                if not filtered_dfs:
                    # 使用MessageBox替代InfoBar
                    MessageBox(
                        "查询结果", 
                        "没有满足条件的数据，请检查查询条件或选择其他工作表。", 
                        self
                    ).exec()
                    self.clearResultTable()
                    return
                
                # 开始合并满足条件的工作表
                if len(filtered_dfs) == 1:
                    # 如果只有一个工作表有满足条件的数据，直接使用该数据
                    merged_df = list(filtered_dfs.values())[0]
                else:
                    # 多个工作表需要合并
                    merged_df = self._mergeFilteredSheets(filtered_dfs, sheet_dfs, sheets_with_conditions, merge_key)
            
            # 如果没有成功合并数据，结束处理
            if merged_df is None or merged_df.empty:
                # 使用MessageBox替代InfoBar
                MessageBox(
                    "查询结果", 
                    "合并后无满足条件的数据，请检查查询条件或选择其他工作表。", 
                    self
                ).exec()
                self.clearResultTable()
                return
                
            # 创建统一的数据来源列
            merged_df['数据来源'] = '合并数据'
            
            # 进行最终的条件过滤，确保结果只包含满足所有条件的记录
            if all_query_fields:
                try:
                    # 将所有查询条件应用到合并后的数据
                    final_filtered_df = self._applyFinalFiltering(merged_df, all_query_fields)
                    
                    # 如果过滤后无数据，显示信息并返回
                    if final_filtered_df.empty:
                        # 使用MessageBox替代InfoBar
                        MessageBox(
                            "查询结果", 
                            "合并后应用所有条件筛选，无匹配记录。请检查查询条件是否过于严格。", 
                            self
                        ).exec()
                        self.clearResultTable()
                        return
                    
                    # 使用最终过滤后的数据
                    merged_df = final_filtered_df
                except Exception as e:
                    MessageBox(
                        "最终过滤错误", 
                        f"应用最终查询条件时出错: {str(e)}\n请检查查询条件。", 
                        self
                    ).exec()
                    self.clearResultTable()
                    return
            
            # 筛选显示列
            self._processAndDisplayResults(merged_df)
            
        except Exception as e:
            # 在合并模式出错时显示错误信息并清空结果
            MessageBox(
                "合并查询错误", 
                f"合并查询出错: {str(e)}\n请检查查询条件和工作表数据。", 
                self
            ).exec()
            self.clearResultTable()

    def _applySingleCondition(self, df, column, operator, value):
        """应用单个查询条件并返回掩码"""
        import warnings
        warnings.filterwarnings("ignore", category=UserWarning, module="pandas.core.tools.datetimes")
        
        date_formats = [
            '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', 
            '%Y/%m/%d', '%Y-%m-%d %H:%M:%S',
            '%d-%m-%Y', '%m-%d-%Y'
        ]
        
        # 如果列不存在，返回全False掩码
        if column not in df.columns:
            return pd.Series([False] * len(df))
            
        # 检测日期列
        is_datetime_column = False
        date_col = None
        
        try:
            if pd.api.types.is_datetime64_any_dtype(df[column].dtype):
                is_datetime_column = True
                date_col = df[column]
            else:
                # 尝试转换第一个非空值来检测是否可能是日期字符串
                sample = df[column].dropna().iloc[0] if not df[column].dropna().empty else None
                if sample and isinstance(sample, str):
                    try:
                        # 尝试使用常见日期格式进行解析
                        for date_format in date_formats:
                            try:
                                date_col = pd.to_datetime(df[column], format=date_format, errors='coerce')
                                # 如果大部分值不是NaT，说明找到了正确的格式
                                if date_col.notna().sum() > 0.5 * len(date_col):
                                    is_datetime_column = True
                                    break
                            except:
                                continue
                                
                        # 如果所有尝试失败，最后尝试自动推断
                        if not is_datetime_column:
                            with warnings.catch_warnings():
                                warnings.simplefilter("ignore")
                                date_col = pd.to_datetime(df[column], errors='coerce')
                                # 如果转换成功（不全是NaT），则视为日期列
                                if not date_col.isna().all() and date_col.notna().sum() > 0.5 * len(date_col):
                                    is_datetime_column = True
                    except:
                        pass
        except:
            pass
        
        # 根据操作符和字段类型构建查询条件
        if operator == "包含":
            if is_datetime_column:
                # 对日期列进行字符串包含查询
                str_col = df[column].astype(str)
                return str_col.str.contains(value, case=False, na=False)
            elif pd.api.types.is_numeric_dtype(df[column]):
                # 数值列转字符串后包含查询
                return df[column].astype(str).str.contains(value, case=False, na=False)
            else:
                # 字符串列直接包含查询
                return df[column].astype(str).str.contains(value, case=False, na=False)
        
        elif operator == "不包含":
            if is_datetime_column:
                str_col = df[column].astype(str)
                return ~str_col.str.contains(value, case=False, na=False)
            elif pd.api.types.is_numeric_dtype(df[column]):
                return ~df[column].astype(str).str.contains(value, case=False, na=False)
            else:
                return ~df[column].astype(str).str.contains(value, case=False, na=False)
        
        elif operator == "等于":
            if is_datetime_column:
                try:
                    # 尝试使用已知的日期格式解析查询值
                    query_date = None
                    for date_format in date_formats:
                        try:
                            query_date = pd.to_datetime(value, format=date_format)
                            break
                        except:
                            continue
                            
                    # 如果格式化失败，尝试自动推断
                    if query_date is None:
                        query_date = pd.to_datetime(value)
                        
                    return date_col.dt.date == query_date.date()
                except:
                    return df[column].astype(str) == value
            elif pd.api.types.is_numeric_dtype(df[column]):
                try:
                    return df[column] == float(value)
                except ValueError:
                    return df[column].astype(str) == value
            else:
                return df[column].astype(str) == value
        
        # ... [为其他操作符添加相似的逻辑]
        # 基本操作符的逻辑
        elif operator == "大于":
            if is_datetime_column:
                try:
                    for date_format in date_formats:
                        try:
                            query_date = pd.to_datetime(value, format=date_format)
                            break
                        except:
                            continue
                    if 'query_date' not in locals():
                        query_date = pd.to_datetime(value)
                    return date_col > query_date
                except:
                    return df[column].astype(str) > value
            else:
                try:
                    return df[column] > float(value)
                except:
                    return df[column].astype(str) > value
                    
        elif operator == "小于":
            if is_datetime_column:
                try:
                    for date_format in date_formats:
                        try:
                            query_date = pd.to_datetime(value, format=date_format)
                            break
                        except:
                            continue
                    if 'query_date' not in locals():
                        query_date = pd.to_datetime(value)
                    return date_col < query_date
                except:
                    return df[column].astype(str) < value
            else:
                try:
                    return df[column] < float(value)
                except:
                    return df[column].astype(str) < value
                    
        # 默认返回全False
        return pd.Series([False] * len(df))

    def _mergeAllSheets(self, sheet_dfs, merge_key):
        """合并所有工作表，不考虑查询条件"""
        if not sheet_dfs:
            return None
            
        merged_df = None
        sheet_names = list(sheet_dfs.keys())
        
        for i, sheet_name in enumerate(sheet_names):
            if i == 0:
                merged_df = sheet_dfs[sheet_name].copy()
                continue
                
            try:
                merged_df = pd.merge(
                    merged_df,
                    sheet_dfs[sheet_name],
                    on=merge_key,
                    how=self.merge_how,
                    suffixes=(f'_{sheet_names[0]}', f'_{sheet_name}')
                )
            except Exception as e:
                InfoBar.warning(
                    title="合并错误",
                    content=f"合并工作表 '{sheet_name}' 时出错: {str(e)}",
                    parent=self,
                    position=InfoBarPosition.TOP,
                    duration=3000
                )
        
        return merged_df
    
    def _mergeFilteredSheets(self, filtered_dfs, sheet_dfs, sheets_with_conditions, merge_key):
        """合并经过过滤的工作表数据"""
        if not filtered_dfs:
            return None
            
        # 根据合并策略决定如何处理
        how = self.merge_how.lower()
        
        # 获取第一个工作表作为基础
        first_sheet = list(filtered_dfs.keys())[0]
        result_df = filtered_dfs[first_sheet]
        
        # 合并其余工作表
        for sheet_name, df in list(filtered_dfs.items())[1:]:
            try:
                # 确保两个DataFrame都有合并键
                if merge_key in result_df.columns and merge_key in df.columns:
                    # 应用合并
                    result_df = pd.merge(
                        result_df, 
                        df,
                        on=merge_key,
                        how=how,
                        suffixes=('', f'_{sheet_name}')
                    )
                else:
                    # 如果合并键不存在，记录错误并跳过
                    MessageBox(
                        "合并错误", 
                        f"合并键 '{merge_key}' 在工作表 '{sheet_name}' 中不存在，跳过此工作表。", 
                        self
                    ).exec()
            except Exception as e:
                MessageBox(
                    "合并错误", 
                    f"合并工作表 '{sheet_name}' 时出错: {str(e)}", 
                    self
                ).exec()
        
        # 对于非inner join，需要处理未设置查询条件的工作表
        if how in ['outer', 'left'] and len(sheets_with_conditions) < len(sheet_dfs):
            # 找出未设置条件的工作表
            sheets_without_conditions = set(sheet_dfs.keys()) - sheets_with_conditions
            
            # 合并未设置条件的工作表（如有必要）
            if sheets_without_conditions:
                for sheet_name in sheets_without_conditions:
                    if sheet_name in sheet_dfs:
                        df = sheet_dfs[sheet_name]
                        try:
                            # 确保两个DataFrame都有合并键
                            if merge_key in result_df.columns and merge_key in df.columns:
                                # 应用合并
                                result_df = pd.merge(
                                    result_df, 
                                    df,
                                    on=merge_key,
                                    how=how,
                                    suffixes=('', f'_{sheet_name}')
                                )
                            else:
                                # 如果合并键不存在，记录错误并跳过
                                MessageBox(
                                    "合并错误", 
                                    f"合并键 '{merge_key}' 在工作表 '{sheet_name}' 中不存在，跳过此工作表。", 
                                    self
                                ).exec()
                        except Exception as e:
                            MessageBox(
                                "合并错误", 
                                f"合并工作表 '{sheet_name}' 时出错: {str(e)}", 
                                self
                            ).exec()
        
        return result_df

    def _applyFinalFiltering(self, merged_df, all_query_fields):
        """对合并后的数据应用最终查询条件，确保只返回满足所有条件的数据"""
        if not all_query_fields or merged_df.empty:
            return merged_df
            
        # 收集错误信息
        error_messages = []
        
        # 创建一个全True的掩码，初始选中所有行
        mask = pd.Series([True] * len(merged_df))
        
        # 遍历每个查询字段并应用条件
        for field in all_query_fields:
            if len(field) >= 3 and field[2].text().strip():
                full_column = field[0].currentText()
                operator = field[1].currentText()
                value = field[2].text().strip()
                
                # 处理带工作表前缀的列名
                if '.' in full_column:
                    sheet_name, column = full_column.split('.', 1)
                    
                    # 查看目标列是否存在于合并后的数据中
                    if column in merged_df.columns:
                        target_column = column
                    elif full_column in merged_df.columns:
                        target_column = full_column
                    else:
                        # 列不存在，添加错误信息
                        error_messages.append(f"列 '{full_column}' 在合并数据中不存在")
                        continue
                else:
                    # 直接使用列名
                    if full_column in merged_df.columns:
                        target_column = full_column
                    else:
                        # 列不存在，添加错误信息
                        error_messages.append(f"列 '{full_column}' 在合并数据中不存在")
                        continue
                
                # 应用单个条件
                condition_mask = self._applySingleCondition(merged_df, target_column, operator, value)
                
                # 如果条件无匹配数据，添加错误信息
                if not condition_mask.any():
                    error_messages.append(f"条件 '{target_column} {operator} {value}' 在合并数据中没有匹配记录")
                
                # 结合当前条件掩码
                mask = mask & condition_mask
        
        # 如果有错误信息，显示并返回空DataFrame
        if error_messages:
            message = "合并后应用最终查询条件出错:\n\n"
            for idx, err in enumerate(error_messages, 1):
                message += f"{idx}. {err}\n"
            message += "\n请检查查询条件是否与合并数据兼容。"
            
            MessageBox("查询结果", message, self).exec()
            return pd.DataFrame()
        
        # 返回经过筛选的数据
        return merged_df[mask]
            
    def _getAllQueryFields(self):
        """获取所有查询字段"""
        return [field for field in self.query_fields if len(field) >= 3 and field[2].text().strip()]

    def _getSheetSpecificQueryFields(self, sheet_name):
        """获取特定工作表的查询字段"""
        sheet_query_fields = []
        
        for field in self.query_fields:
            if len(field) >= 3:
                column_full = field[0].currentText()
                
                # 检查列名是否属于特定工作表
                if "." in column_full:
                    field_sheet, field_col = column_full.split(".", 1)
                    if field_sheet == sheet_name:
                        # 创建一个新的查询字段元组，但将列名修改为不带工作表前缀的版本
                        new_field = list(field)
                        # 创建一个临时ComboBox来替换原始的列选择框
                        temp_combo = ComboBox()
                        temp_combo.addItem(field_col)
                        temp_combo.setCurrentIndex(0)
                        new_field[0] = temp_combo
                        sheet_query_fields.append(tuple(new_field))
                else:
                    # 如果列名不包含"."，则假定它是一个共同列，可以直接使用
                    sheet_query_fields.append(field)
            else:
                # 处理旧格式的查询字段
                sheet_query_fields.append(field)
        
        return sheet_query_fields
        
    def _addQueryField(self):
        """添加查询字段"""
        if not self.sheets or not self.selected_sheets:
            return
        
        # 获取所有选择的工作表中的列
        columns = self._getAllQueryColumns()
        if not columns:
            InfoBar.warning(
                title="无法添加查询字段",
                content="所选工作表没有可用列",
                parent=self,
                position=InfoBarPosition.TOP,
                duration=3000
            )
            return
            
        # 创建字段选择器组件
        fieldWidget = QWidget()
        layout = QHBoxLayout(fieldWidget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)
        
        # 如果已经有查询字段，添加且/或选择器
        logicCombo = None
        if self.query_fields:
            logicCombo = ComboBox()
            logicCombo.addItems(["且", "或", "非"])
            logicCombo.setCurrentIndex(0)  # 默认选择"且"
            logicCombo.setFixedWidth(70)  # 增加宽度，从50改为70
            layout.addWidget(logicCombo)
        
        # 列选择下拉框
        comboBox = ComboBox()
        comboBox.addItems(columns)
        comboBox.setMinimumWidth(150)  # 增加最小宽度以适应带工作表名的更长列名
        # 默认选择第一个字段
        if columns:
            comboBox.setCurrentIndex(0)
        
        # 比较操作符下拉框
        operatorCombo = ComboBox()
        operatorCombo.addItems(["包含", "不包含", "等于", "不等于", "大于", "小于", "大于等于", "小于等于", "介于"])
        operatorCombo.setCurrentIndex(0)  # 默认选择"包含"
        operatorCombo.setMinimumWidth(100)  # 增加最小宽度以确保完整显示内容
        
        # 操作符变化时更新输入框提示文本
        def updatePlaceholder(index):
            op = operatorCombo.currentText()
            if op == "介于":
                valueEdit.setPlaceholderText("最小值,最大值")
            elif op == "包含":
                valueEdit.setPlaceholderText("包含文本")
            elif op == "不包含":
                valueEdit.setPlaceholderText("不包含文本")
            else:
                valueEdit.setPlaceholderText("输入值")
        
        operatorCombo.currentIndexChanged.connect(updatePlaceholder)
        
        # 值输入框
        valueEdit = LineEdit()
        valueEdit.setPlaceholderText("包含文本")
        valueEdit.setMinimumWidth(150)  # 设置最小宽度
        valueEdit.setMaximumWidth(300)  # 设置最大宽度，防止输入框过长
        valueEdit.setClearButtonEnabled(True)
        
        # 添加文本改变事件，根据内容自动调整宽度（在最大宽度范围内）
        def adjustWidth():
            text = valueEdit.text()
            # 计算文本大概宽度 (每个字符约10像素)
            textWidth = len(text) * 10
            # 设置宽度，最小150，最大300，根据内容自动调整
            valueEdit.setMinimumWidth(min(max(150, textWidth + 30), 300))
        
        valueEdit.textChanged.connect(adjustWidth)
        
        # 删除按钮
        removeButton = ToolButton(FluentIcon.DELETE)
        removeButton.setToolTip("移除此查询条件")
        removeButton.setIconSize(QSize(14, 14))
        removeButton.clicked.connect(lambda: self._removeQueryField(fieldWidget))
        
        layout.addWidget(comboBox)
        layout.addWidget(operatorCombo)
        layout.addWidget(valueEdit)
        layout.addWidget(removeButton)
        layout.addStretch(1)
        
        # 将字段组件添加到查询字段容器
        self.queryFieldsLayout.addWidget(fieldWidget)
        
        # 保存查询字段信息，现在包括列选择、操作符和值输入框，以及可选的逻辑选择器
        if logicCombo:
            self.query_fields.append((comboBox, operatorCombo, valueEdit, logicCombo))
        else:
            self.query_fields.append((comboBox, operatorCombo, valueEdit, None))
        
        # 更新执行按钮状态
        self._updateExecuteButtonState()
    
    def _getAllQueryColumns(self):
        """获取所有可用于查询的列，包括所有工作表的所有列"""
        # 获取共同列（用于常规模式）
        common_columns = self._getCommonColumns()
        
        # 获取当前选择的工作表
        selected_sheet_names = [button.text() for button in self.selected_sheets if button.isChecked()]
        
        # 处理模式 - 获取当前模式
        processing_mode = self.processingModeCombo.currentText() if hasattr(self, 'processingModeCombo') else "堆叠"
        
        # 如果是合并模式，还要添加带工作表名前缀的所有列
        if processing_mode == "合并" and len(selected_sheet_names) >= 2:
            # 收集所有工作表的所有列
            all_columns = []
            
            # 先添加常见列作为基础选项
            if common_columns:
                all_columns.extend(common_columns)
            
            # 然后为每个工作表添加带前缀的列
            for sheet_name in selected_sheet_names:
                if sheet_name in self.sheets:
                    df = self.sheets[sheet_name]
                    for column in df.columns:
                        # 添加带工作表名前缀的列名，例如"工作表1.列1"
                        prefixed_column = f"{sheet_name}.{column}"
                        # 确保不重复添加列
                        if prefixed_column not in all_columns:
                            all_columns.append(prefixed_column)
            
            return all_columns
        else:
            # 对于堆叠模式，只返回常见列
            return common_columns

    def _getAllMatchColumns(self):
        """获取所有可用于结果显示的列"""
        # 获取当前选择的工作表
        selected_sheet_names = [button.text() for button in self.selected_sheets if button.isChecked()]
        
        # 处理模式
        processing_mode = self.processingModeCombo.currentText() if hasattr(self, 'processingModeCombo') else "堆叠"
        
        # 对于堆叠模式，我们需要所有可能的列
        if processing_mode == "堆叠":
            # 收集所有选定工作表的所有唯一列
            all_columns = set()
            for sheet_name in selected_sheet_names:
                if sheet_name in self.sheets:
                    df = self.sheets[sheet_name]
                    all_columns.update(df.columns)
            
            # 转换为有序列表
            return ["显示全部列"] + sorted(list(all_columns))
        
        # 对于合并模式，我们需要考虑合并后的所有列
        elif processing_mode == "合并" and len(selected_sheet_names) >= 2:
            # 先获取共同列
            common_columns = self._getCommonColumns()
            
            # 初始化所有可能的列
            all_columns = ["显示全部列"]
            
            # 添加共同列
            if common_columns:
                all_columns.extend(common_columns)
            
            # 为每个工作表添加带前缀的列
            for sheet_name in selected_sheet_names:
                if sheet_name in self.sheets:
                    df = self.sheets[sheet_name]
                    for column in df.columns:
                        # 如果不是共同列，则添加带工作表名前缀的列名
                        if column not in common_columns:
                            prefixed_column = f"{sheet_name}.{column}"
                            if prefixed_column not in all_columns:
                                all_columns.append(prefixed_column)
            
            return all_columns
        else:
            # 如果只有一个工作表或其他情况
            common_columns = self._getCommonColumns()
            return ["显示全部列"] + common_columns if common_columns else ["显示全部列"]

    def _showMergeKeySelectionDialog(self, common_columns):
        """显示合并键选择对话框"""
        # 创建对话框
        dialog = Dialog(self)
        dialog.setWindowTitle("选择合并键")
        
        # 设置对话框内容
        content = QWidget()
        layout = QVBoxLayout(content)
        
        # 添加说明文本
        label = QLabel("请选择用于合并工作表的关联列:")
        layout.addWidget(label)
        
        # 添加选择列表
        comboBox = ComboBox()
        comboBox.addItems(common_columns)
        comboBox.setCurrentIndex(0)  # 默认选择第一个
        layout.addWidget(comboBox)
        
        # 添加合并方式选择
        groupBox = QGroupBox("合并方式")
        radioLayout = QVBoxLayout()
        
        outerJoinRadio = QPushButton("外连接 (保留所有数据)")
        outerJoinRadio.setCheckable(True)
        outerJoinRadio.setChecked(True)
        
        innerJoinRadio = QPushButton("内连接 (仅保留匹配数据)")
        innerJoinRadio.setCheckable(True)
        
        leftJoinRadio = QPushButton("左连接 (保留第一个表的所有数据)")
        leftJoinRadio.setCheckable(True)
        
        # 互斥按钮组
        def select_outer():
            outerJoinRadio.setChecked(True)
            innerJoinRadio.setChecked(False)
            leftJoinRadio.setChecked(False)
            
        def select_inner():
            outerJoinRadio.setChecked(False)
            innerJoinRadio.setChecked(True)
            leftJoinRadio.setChecked(False)
            
        def select_left():
            outerJoinRadio.setChecked(False)
            innerJoinRadio.setChecked(False)
            leftJoinRadio.setChecked(True)
            
        outerJoinRadio.clicked.connect(select_outer)
        innerJoinRadio.clicked.connect(select_inner)
        leftJoinRadio.clicked.connect(select_left)
        
        radioLayout.addWidget(outerJoinRadio)
        radioLayout.addWidget(innerJoinRadio)
        radioLayout.addWidget(leftJoinRadio)
        
        groupBox.setLayout(radioLayout)
        layout.addWidget(groupBox)
        
        # 设置对话框内容
        dialog.setWidget(content)
        dialog.setSizePolicy(QWidget.Minimum, QWidget.Minimum)
        
        # 添加按钮
        dialog.yesButton.setText("确定")
        dialog.cancelButton.setText("取消")
        
        # 显示对话框并获取结果
        if dialog.exec():
            selected_key = comboBox.currentText()
            
            # 获取选择的合并方式
            how = 'outer'  # 默认
            if innerJoinRadio.isChecked():
                how = 'inner'
            elif leftJoinRadio.isChecked():
                how = 'left'
                
            # 设置全局合并方式
            self.merge_how = how
            
            return selected_key
        else:
            return None

    def _processAndDisplayResults(self, df):
        """处理和显示查询结果"""
        # 检查数据是否为空
        if df is None or df.empty:
            # 如果数据为空，弹窗提示并清空结果表格
            MessageBox(
                "查询结果", 
                "未找到匹配记录，请检查查询条件。", 
                self
            ).exec()
            self.clearResultTable()
            return
        
        # 过滤掉全为空值的行（所有列都是NA/NaN/None的行）
        original_count = len(df)
        df = df.dropna(how='all')  # 删除所有列都为NaN的行
        filtered_count = original_count - len(df)
        
        if filtered_count > 0:
            InfoBar.info(
                title="数据清理",
                content=f"已过滤 {filtered_count} 行全空数据",
                parent=self,
                position=InfoBarPosition.TOP,
                duration=3000
            )
            
        # 再次检查清理后的数据是否为空
        if df.empty:
            MessageBox(
                "查询结果", 
                "过滤空值后无匹配记录，请检查查询条件。", 
                self
            ).exec()
            self.clearResultTable()
            return
            
        # 提取显示列
        display_columns = []
        
        for combo, _ in self.match_fields:
            column = combo.currentText()
            
            # 特殊处理：如果选择"显示全部列"
            if column == "显示全部列":
                display_columns = list(df.columns)
                break
                
            # 处理带工作表前缀的列名
            if "." in column and column not in df.columns:
                sheet_name, col_name = column.split(".", 1)
                # 寻找合并后对应的列名
                matched_cols = []
                for df_col in df.columns:
                    # 检查是否可能是带后缀的列，如 "列名_工作表1"
                    if col_name in df_col and (f"_{sheet_name}" in df_col or df_col == col_name):
                        matched_cols.append(df_col)
                
                # 如果找到匹配的列，添加到显示列中
                if matched_cols:
                    display_columns.extend(matched_cols)
                    continue
            
            if column and column in df.columns:
                display_columns.append(column)
                    
        # 如果指定了显示字段，则过滤列
        if display_columns:
            # 确保始终包括"数据来源"列
            if '数据来源' not in display_columns and '数据来源' in df.columns:
                display_columns.append('数据来源')
                
            # 确保所有指定的列都存在
            existing_columns = [col for col in display_columns if col in df.columns]
            if existing_columns:
                # 确保数据来源列在最左侧
                if '数据来源' in existing_columns:
                    existing_columns.remove('数据来源')
                    existing_columns.insert(0, '数据来源')
                    
                df = df[existing_columns]
        
        # 最终检查，确保有可显示的内容
        if df.empty:
            MessageBox(
                "查询结果", 
                "处理后无可显示的数据，请检查查询条件和显示字段设置。", 
                self
            ).exec()
            self.clearResultTable()
            return
            
        # 显示最终结果
        self.displayResults(df)

    def _findCommonColumns(self, dataframes):
        """查找多个DataFrame之间的共同列"""
        if not dataframes:
            return []
            
        # 获取每个DataFrame的列集合
        column_sets = [set(df.columns) for df in dataframes]
        
        # 计算交集得到共同列
        common_columns = set.intersection(*column_sets)
        
        # 按照第一个DataFrame的列顺序返回共同列
        if common_columns and dataframes:
            return [col for col in dataframes[0].columns if col in common_columns]
        
        return list(common_columns)

    def clearResultTable(self):
        """清空结果表格"""
        self.resultTable.clear()
        self.resultTable.setRowCount(0)
        self.resultTable.setColumnCount(0)
        self.result_data = None

    def _applyQueryConditions(self, df, query_fields):
        """应用查询条件到数据框，使用更高效的pandas查询语法，并检测逻辑矛盾"""
        # 如果没有查询条件，则返回原数据
        if not query_fields:
            return df
            
        # 收集所有错误信息
        error_messages = []
            
        # 获取有效的查询条件 (有值的查询字段)
        active_query_fields = []
        for field in query_fields:
            # 检查字段元组的长度，适配新旧格式
            if len(field) >= 3 and field[2].text().strip():
                # 新格式: (列选择框, 操作符选择框, 值输入框, 逻辑选择框)
                column = field[0].currentText()
                operator = field[1].currentText() if len(field) > 1 else "包含"
                value = field[2].text().strip()
                logic = field[3].currentText() if len(field) > 3 and field[3] is not None else "且"
                active_query_fields.append((column, operator, value, logic))
            elif len(field) == 2 and field[1].text().strip():
                # 旧格式: (列选择框, 值输入框)
                column = field[0].currentText()
                value = field[1].text().strip()
                active_query_fields.append((column, "包含", value, "且"))
                
        # 如果没有有效的查询条件，则返回原数据
        if not active_query_fields:
            return df
            
        # 构建布尔索引而不是使用series的累积操作
        all_masks = []  # 存储所有条件的掩码
        all_logics = []  # 存储所有逻辑操作符
        all_conditions = []  # 存储所有条件的描述，用于矛盾检测

        # 预处理数据框 - 对日期列进行一次性转换
        date_columns = {}  # 存储已转换的日期列
        
        # 检测逻辑矛盾
        conflict_columns = {}  # 按列存储可能冲突的条件
        
        # 常用日期格式列表，用于尝试解析日期
        date_formats = [
            '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', 
            '%Y/%m/%d', '%Y-%m-%d %H:%M:%S',
            '%d-%m-%Y', '%m-%d-%Y'
        ]
        
        # 禁用pandas日期解析警告
        import warnings
        warnings.filterwarnings("ignore", category=UserWarning, module="pandas.core.tools.datetimes")
        
        # 处理每个查询条件
        for i, (column, operator, value, logic) in enumerate(active_query_fields):
            # 第一个条件不需要记录逻辑
            if i > 0:
                all_logics.append(logic)
                
            if column not in df.columns:
                error_messages.append(f"查询字段中的列 '{column}' 在工作表中不存在")
                continue
            
            # 记录条件
            condition_info = {"column": column, "operator": operator, "value": value}
            all_conditions.append(condition_info)
            
            # 如果列已有条件，添加进冲突检测字典
            if column not in conflict_columns:
                conflict_columns[column] = []
            conflict_columns[column].append(condition_info)
                
            # 创建当前条件的掩码
            current_mask = None
            
            try:
                # 检测日期列并缓存转换结果
                is_datetime_column = False
                date_col = None
                
                # 检查是否已将此列识别为日期
                if column in date_columns:
                    is_datetime_column = True
                    date_col = date_columns[column]
                else:
                    # 尝试检测日期列
                    try:
                        if pd.api.types.is_datetime64_any_dtype(df[column].dtype):
                            is_datetime_column = True
                            date_col = df[column]
                            date_columns[column] = date_col
                        else:
                            # 尝试转换第一个非空值来检测是否可能是日期字符串
                            sample = df[column].dropna().iloc[0] if not df[column].dropna().empty else None
                            if sample and isinstance(sample, str):
                                try:
                                    # 尝试使用常见日期格式进行解析，避免dateutil自动推断警告
                                    for date_format in date_formats:
                                        try:
                                            date_col = pd.to_datetime(df[column], format=date_format, errors='coerce')
                                            # 如果大部分值不是NaT，说明找到了正确的格式
                                            if date_col.notna().sum() > 0.5 * len(date_col):
                                                is_datetime_column = True
                                                date_columns[column] = date_col
                                                break
                                        except:
                                            continue
                                            
                                    # 如果所有尝试失败，最后尝试自动推断，但这里会静默处理警告
                                    if not is_datetime_column:
                                        with warnings.catch_warnings():
                                            warnings.simplefilter("ignore")
                                            date_col = pd.to_datetime(df[column], errors='coerce')
                                            # 如果转换成功（不全是NaT），则视为日期列
                                            if not date_col.isna().all() and date_col.notna().sum() > 0.5 * len(date_col):
                                                is_datetime_column = True
                                                date_columns[column] = date_col
                                except:
                                    pass
                    except:
                        pass
                    
                # 根据操作符和字段类型构建查询条件
                if operator == "包含":
                    if is_datetime_column:
                        # 对日期列进行字符串包含查询
                        str_col = df[column].astype(str)
                        current_mask = str_col.str.contains(value, case=False, na=False)
                    elif pd.api.types.is_numeric_dtype(df[column]):
                        # 数值列转字符串后包含查询
                        current_mask = df[column].astype(str).str.contains(value, case=False, na=False)
                    else:
                        # 字符串列直接包含查询
                        current_mask = df[column].astype(str).str.contains(value, case=False, na=False)
                
                elif operator == "不包含":
                    if is_datetime_column:
                        str_col = df[column].astype(str)
                        current_mask = ~str_col.str.contains(value, case=False, na=False)
                    elif pd.api.types.is_numeric_dtype(df[column]):
                        current_mask = ~df[column].astype(str).str.contains(value, case=False, na=False)
                    else:
                        current_mask = ~df[column].astype(str).str.contains(value, case=False, na=False)
                
                elif operator == "等于":
                    if is_datetime_column:
                        try:
                            # 尝试使用已知的日期格式解析查询值
                            query_date = None
                            for date_format in date_formats:
                                try:
                                    query_date = pd.to_datetime(value, format=date_format)
                                    break
                                except:
                                    continue
                                    
                            # 如果格式化失败，尝试自动推断
                            if query_date is None:
                                query_date = pd.to_datetime(value)
                                
                            current_mask = date_col.dt.date == query_date.date()
                        except:
                            current_mask = df[column].astype(str) == value
                    elif pd.api.types.is_numeric_dtype(df[column]):
                        try:
                            current_mask = df[column] == float(value)
                        except ValueError:
                            current_mask = df[column].astype(str) == value
                    else:
                        current_mask = df[column].astype(str) == value
                
                elif operator == "不等于":
                    if is_datetime_column:
                        try:
                            # 尝试使用已知的日期格式解析查询值
                            query_date = None
                            for date_format in date_formats:
                                try:
                                    query_date = pd.to_datetime(value, format=date_format)
                                    break
                                except:
                                    continue
                                    
                            # 如果格式化失败，尝试自动推断
                            if query_date is None:
                                query_date = pd.to_datetime(value)
                                
                            current_mask = date_col.dt.date != query_date.date()
                        except:
                            current_mask = df[column].astype(str) != value
                    elif pd.api.types.is_numeric_dtype(df[column]):
                        try:
                            current_mask = df[column] != float(value)
                        except ValueError:
                            current_mask = df[column].astype(str) != value
                    else:
                        current_mask = df[column].astype(str) != value
                
                elif operator == "大于":
                    if is_datetime_column:
                        try:
                            # 尝试使用已知的日期格式解析查询值
                            query_date = None
                            for date_format in date_formats:
                                try:
                                    query_date = pd.to_datetime(value, format=date_format)
                                    break
                                except:
                                    continue
                                    
                            # 如果格式化失败，尝试自动推断
                            if query_date is None:
                                query_date = pd.to_datetime(value)
                                
                            current_mask = date_col > query_date
                        except:
                            current_mask = df[column].astype(str) > value
                    else:
                        try:
                            current_mask = df[column] > float(value)
                        except (ValueError, TypeError):
                            current_mask = df[column].astype(str) > value
                
                elif operator == "小于":
                    if is_datetime_column:
                        try:
                            # 尝试使用已知的日期格式解析查询值
                            query_date = None
                            for date_format in date_formats:
                                try:
                                    query_date = pd.to_datetime(value, format=date_format)
                                    break
                                except:
                                    continue
                                    
                            # 如果格式化失败，尝试自动推断
                            if query_date is None:
                                query_date = pd.to_datetime(value)
                                
                            current_mask = date_col < query_date
                        except:
                            current_mask = df[column].astype(str) < value
                    else:
                        try:
                            current_mask = df[column] < float(value)
                        except (ValueError, TypeError):
                            current_mask = df[column].astype(str) < value
                
                elif operator == "大于等于":
                    if is_datetime_column:
                        try:
                            # 尝试使用已知的日期格式解析查询值
                            query_date = None
                            for date_format in date_formats:
                                try:
                                    query_date = pd.to_datetime(value, format=date_format)
                                    break
                                except:
                                    continue
                                    
                            # 如果格式化失败，尝试自动推断
                            if query_date is None:
                                query_date = pd.to_datetime(value)
                                
                            current_mask = date_col >= query_date
                        except:
                            current_mask = df[column].astype(str) >= value
                    else:
                        try:
                            current_mask = df[column] >= float(value)
                        except (ValueError, TypeError):
                            current_mask = df[column].astype(str) >= value
                
                elif operator == "小于等于":
                    if is_datetime_column:
                        try:
                            # 尝试使用已知的日期格式解析查询值
                            query_date = None
                            for date_format in date_formats:
                                try:
                                    query_date = pd.to_datetime(value, format=date_format)
                                    break
                                except:
                                    continue
                                    
                            # 如果格式化失败，尝试自动推断
                            if query_date is None:
                                query_date = pd.to_datetime(value)
                                
                            current_mask = date_col <= query_date
                        except:
                            current_mask = df[column].astype(str) <= value
                    else:
                        try:
                            current_mask = df[column] <= float(value)
                        except (ValueError, TypeError):
                            current_mask = df[column].astype(str) <= value
                            
                elif operator == "介于":
                    try:
                        min_val, max_val = value.split(",", 1)
                        min_val = min_val.strip()
                        max_val = max_val.strip()
                        
                        if is_datetime_column:
                            try:
                                # 尝试使用已知的日期格式解析最小日期
                                min_date = None
                                for date_format in date_formats:
                                    try:
                                        min_date = pd.to_datetime(min_val, format=date_format)
                                        break
                                    except:
                                        continue
                                        
                                # 如果格式化失败，尝试自动推断
                                if min_date is None:
                                    min_date = pd.to_datetime(min_val)
                                    
                                # 尝试使用已知的日期格式解析最大日期
                                max_date = None
                                for date_format in date_formats:
                                    try:
                                        max_date = pd.to_datetime(max_val, format=date_format)
                                        break
                                    except:
                                        continue
                                        
                                # 如果格式化失败，尝试自动推断
                                if max_date is None:
                                    max_date = pd.to_datetime(max_val)
                                    
                                current_mask = (date_col >= min_date) & (date_col <= max_date)
                            except:
                                current_mask = (df[column].astype(str) >= min_val) & (df[column].astype(str) <= max_val)
                        else:
                            try:
                                min_num = float(min_val)
                                max_num = float(max_val)
                                # 检查范围是否有效
                                if min_num > max_num:
                                    raise ValueError(f"列 '{column}' 的范围无效: {min_num} 大于 {max_num}")
                                current_mask = (df[column] >= min_num) & (df[column] <= max_num)
                            except (ValueError, TypeError):
                                current_mask = (df[column].astype(str) >= min_val) & (df[column].astype(str) <= max_val)
                    except ValueError as e:
                        if "范围无效" in str(e):
                            raise
                        current_mask = df[column].astype(str) == value
                
                # 将当前掩码添加到列表中
                all_masks.append(current_mask)
                
                # 检查单个条件是否满足
                if not current_mask.any():
                    error_messages.append(f"条件 '{column} {operator} {value}' 没有匹配的数据")
                
                # 如果不是第一个条件，检查组合结果
                if i > 0:
                    # 计算到当前条件为止的组合结果
                    temp_mask = all_masks[0]
                    for j in range(len(all_logics)):
                        if j < i:  # 计算到当前条件
                            if all_logics[j] == "且":
                                temp_mask = temp_mask & all_masks[j+1]
                            elif all_logics[j] == "或":
                                temp_mask = temp_mask | all_masks[j+1]
                            elif all_logics[j] == "非":
                                temp_mask = temp_mask & (~all_masks[j+1])
                    
                    # 检查组合结果是否为空
                    if not temp_mask.any():
                        error_messages.append(f"条件 '{column} {operator} {value}' 与前面的条件组合后没有匹配数据")
                    
            except Exception as e:
                # 记录任何其他异常
                error_messages.append(f"应用条件 '{column} {operator} {value}' 时出错: {str(e)}")
        
        # 在应用所有条件前，检查是否有错误信息
        if error_messages:
            # 构建清晰的错误消息
            message = "查询条件错误:\n\n"
            for idx, err in enumerate(error_messages, 1):
                message += f"{idx}. {err}\n"
            message += "\n请检查查询条件。"
            
            # 显示消息并返回空结果
            MessageBox("查询结果", message, self).exec()
            return df.iloc[0:0]  # 返回空DataFrame
            
        # 检查是否有逻辑矛盾
        contradictions = self._checkLogicalContradictions(conflict_columns)
        if contradictions:
            message = "查询条件存在逻辑矛盾:\n\n"
            for idx, contra in enumerate(contradictions, 1):
                message += f"{idx}. {contra}\n"
            message += "\n请修改条件。"
            
            MessageBox("条件矛盾", message, self).exec()
            return df.iloc[0:0]  # 返回空DataFrame
        
        # 合并所有条件
        if not all_masks:
            return df
            
        # 从第一个条件开始
        result_mask = all_masks[0]
        
        # 根据逻辑运算符合并后续条件
        for i in range(len(all_logics)):
            logic = all_logics[i]
            mask = all_masks[i+1]
            
            if logic == "且":
                result_mask = result_mask & mask
            elif logic == "或":
                result_mask = result_mask | mask
            elif logic == "非":
                result_mask = result_mask & (~mask)
        
        # 应用最终的条件掩码
        final_result = df[result_mask]
        
        # 再次确认结果不为空
        if final_result.empty:
            MessageBox("查询结果", "没有满足所有查询条件的数据。\n请检查查询条件或选择其他工作表。", self).exec()
            return df.iloc[0:0]  # 返回空DataFrame
            
        return final_result

    def _checkLogicalContradictions(self, conflict_columns):
        """检查查询条件中的逻辑矛盾，返回矛盾列表"""
        contradictions = []
        
        for column, conditions in conflict_columns.items():
            # 如果列只有一个条件，跳过检查
            if len(conditions) < 2:
                continue
                
            # 检查数值型条件的矛盾
            numeric_conditions = []
            for cond in conditions:
                operator = cond["operator"]
                value = cond["value"]
                
                # 只处理明确的数值比较
                if operator in ["大于", "小于", "大于等于", "小于等于", "等于", "不等于"]:
                    try:
                        # 尝试转换为数值
                        num_value = float(value)
                        numeric_conditions.append({
                            "operator": operator,
                            "value": num_value,
                            "original": cond
                        })
                    except (ValueError, TypeError):
                        # 非数值类型，跳过
                        continue
                elif operator == "介于":
                    try:
                        min_val, max_val = value.split(",", 1)
                        min_val = float(min_val.strip())
                        max_val = float(max_val.strip())
                        
                        # 添加两个条件
                        numeric_conditions.append({
                            "operator": "大于等于",
                            "value": min_val,
                            "original": cond,
                            "part_of_range": True
                        })
                        numeric_conditions.append({
                            "operator": "小于等于",
                            "value": max_val,
                            "original": cond,
                            "part_of_range": True
                        })
                    except (ValueError, TypeError):
                        # 格式错误，跳过
                        continue
            
            # 如果有足够的数值条件，检查矛盾
            if len(numeric_conditions) >= 2:
                # 检查例如 "大于10" 且 "小于5" 这样的矛盾
                min_bound = float('-inf')
                max_bound = float('inf')
                equal_values = set()
                not_equal_values = set()
                
                for cond in numeric_conditions:
                    op = cond["operator"]
                    val = cond["value"]
                    
                    if op == "大于":
                        min_bound = max(min_bound, val)
                    elif op == "大于等于":
                        min_bound = max(min_bound, val)
                    elif op == "小于":
                        max_bound = min(max_bound, val)
                    elif op == "小于等于":
                        max_bound = min(max_bound, val)
                    elif op == "等于":
                        equal_values.add(val)
                    elif op == "不等于":
                        not_equal_values.add(val)
                
                # 检查各种矛盾
                if min_bound > max_bound:
                    contradictions.append(f"列 '{column}' 的条件矛盾: 大于 {min_bound} 且 小于 {max_bound}")
                
                # 检查等于值是否在范围内
                for val in equal_values:
                    if val < min_bound or val > max_bound:
                        contradictions.append(f"列 '{column}' 的条件矛盾: 等于 {val} 但范围是 {min_bound} 到 {max_bound}")
                
                # 检查多个不同的等于值
                if len(equal_values) > 1:
                    contradictions.append(f"列 '{column}' 的条件矛盾: 同时等于多个不同的值 {', '.join(map(str, equal_values))}")
        
        return contradictions

    def _applyDisplayColumns(self, df, match_fields):
        """应用显示字段到数据框"""
        # 获取要显示的列
        display_columns = [combo.currentText() for combo in match_fields]
        
        # 如果没有指定显示字段，则返回原数据
        if not display_columns:
            return df
            
        # 检查指定的列是否存在
        missing_cols = [col for col in display_columns if col not in df.columns]
        if missing_cols:
            # 过滤掉不存在的列
            display_columns = [col for col in display_columns if col in df.columns]
            
            # 如果所有指定的列都不存在，则显示警告并返回原数据
            if not display_columns:
                InfoBar.warning(
                    title="注意",
                    content=f"指定的显示字段列不存在",
                    parent=self,
                    duration=5000,
                    position=InfoBarPosition.TOP
                )
                return df
                
            # 显示警告但继续使用存在的列
            InfoBar.warning(
                title="注意",
                content=f"部分显示字段列不存在: {', '.join(missing_cols)}",
                parent=self,
                duration=5000,
                position=InfoBarPosition.TOP
            )
            
        # 返回只包含指定列的数据
        return df[display_columns]

    def onSheetChanged(self, index):
        """工作表变更时的处理"""
        if index < 0 or not self.sheets:
            return

        sheet_name = self.sheetComboBox.currentText()
        if sheet_name not in self.sheets:
             MessageBox("错误", f"找不到工作表 '{sheet_name}' 的数据。", self).exec()
             return

        self.current_sheet = sheet_name
        # Ensure data is a DataFrame
        current_df = self.sheets.get(self.current_sheet)
        if not isinstance(current_df, pd.DataFrame):
             MessageBox("错误", f"工作表 '{sheet_name}' 的数据格式不正确。", self).exec()
             self.columns = []
             # Disable further actions
             self.addQueryFieldButton.setEnabled(False)
             self.addMatchFieldButton.setEnabled(False)
             self.executeQueryButton.setEnabled(False)
             return

        self.columns = list(current_df.columns)

        # 清空之前的查询和匹配字段
        self._clearAllFields()
        self.clearResultTable()

        # 启用添加字段按钮 if columns exist
        has_columns = bool(self.columns)
        self.addQueryFieldButton.setEnabled(has_columns)
        self.addMatchFieldButton.setEnabled(has_columns)
        # Execute button should only be enabled when fields are added
        self.executeQueryButton.setEnabled(False)

    def _reflowSheetSelectionLayout(self):
        """重新排列工作表选择布局，填充空白区域"""
        # 直接请求重新计算布局
        self.sheetSelectionLayout.update()
        self.sheetSelectionContainer.updateGeometry()
        self.sheetSelectionContainer.update()

    def _showModeInfo(self):
        """显示数据处理模式的详细说明"""
        html = """
        <h3>数据处理模式说明</h3>
        <p><b>堆叠模式:</b></p>
        <ul>
            <li>将多个工作表的数据垂直堆叠在一起</li>
            <li>适用于工作表结构相似的情况</li>
            <li>每个工作表的查询结果会按行合并</li>
            <li>所有结果会增加"数据来源"列标识数据来自哪个工作表</li>
        </ul>
        <p><b>合并模式:</b></p>
        <ul>
            <li>通过共同字段将多个工作表的数据横向合并</li>
            <li>适用于工作表之间有关联关系的情况</li>
            <li>会自动检测工作表间的共同列作为关联键</li>
            <li>没有共同列或合并失败时，会自动切换回堆叠模式</li>
        </ul>
        """
        
        # 使用MessageBox替代Flyout，避免参数问题
        MessageBox(
            title="数据处理模式说明", 
            content=html,
            parent=self
        ).exec()

    def _addMatchField(self):
        """添加显示字段"""
        if not self.sheets or not self.selected_sheets:
            return
            
        # 获取所有可用于显示的列
        columns = self._getAllMatchColumns()
        if not columns:
            InfoBar.warning(
                title="无法添加显示字段",
                content="所选工作表没有可用列",
                parent=self,
                position=InfoBarPosition.TOP,
                duration=3000
            )
            return
            
        # 创建字段选择器组件（简化为只有列选择和删除按钮）
        fieldWidget = QWidget()
        layout = QHBoxLayout(fieldWidget)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)
        
        comboBox = ComboBox()
        comboBox.addItems(columns)
        comboBox.setMinimumWidth(150)  # 增加最小宽度以适应带工作表名的更长列名
        # 默认选择第一个字段
        comboBox.setCurrentIndex(0)  # 默认选择"显示全部列"
        
        removeButton = ToolButton(FluentIcon.DELETE)
        removeButton.setToolTip("移除此显示字段")
        removeButton.setIconSize(QSize(14, 14))
        removeButton.clicked.connect(lambda: self._removeMatchField(fieldWidget))
        
        layout.addWidget(comboBox)
        layout.addWidget(removeButton)
        
        # 添加到FlowLayout
        self.matchFieldsLayout.addWidget(fieldWidget)
        
        # 保存显示字段信息（不再需要自定义标题输入框）
        self.match_fields.append((comboBox, None))
        
        # 更新执行按钮状态
        self._updateExecuteButtonState()

    def _updateExecuteButtonState(self):
        """更新执行查询按钮状态"""
        # 检查是否有选择的工作表
        has_selected_sheets = False
        for button in self.selected_sheets:
            if button.isChecked():
                has_selected_sheets = True
                break
                
        # 更新执行按钮状态
        self.executeQueryButton.setEnabled(has_selected_sheets)
        
        # 处理模式
        processing_mode = self.processingModeCombo.currentText() if hasattr(self, 'processingModeCombo') else "堆叠"
        
        # 更新添加字段按钮状态
        if has_selected_sheets:
            # 合并模式下，即使没有共同列也可以添加查询和显示字段
            if processing_mode == "合并":
                self.addQueryButton.setEnabled(True)
                self.addMatchButton.setEnabled(True)
            else:
                # 堆叠模式下，需要检查是否有共同列
                has_common_columns = bool(self._getCommonColumns())
                self.addQueryButton.setEnabled(has_common_columns)
                self.addMatchButton.setEnabled(has_common_columns)
        else:
            self.addQueryButton.setEnabled(False)
            self.addMatchButton.setEnabled(False)

    def _removeMatchField(self, widget):
        """移除显示字段"""
        # 查找组件在列表中的索引
        found_index = -1
        for i, (combo, _) in enumerate(self.match_fields):
            if combo.parentWidget() == widget:
                found_index = i
                break
                
        if found_index != -1:
            # 从列表中移除
            self.match_fields.pop(found_index)
            
            # 从布局中移除并删除组件
            widget.deleteLater()
            
            # 立即更新布局
            self._reflowMatchFieldsLayout()
            
            # 更新执行按钮状态
            self._updateExecuteButtonState()
    
    def _reflowMatchFieldsLayout(self):
        """重新排列显示字段布局，填充空白区域"""
        # 直接请求重新计算布局
        self.matchFieldsLayout.update()
        self.matchFieldsContainer.updateGeometry()
        self.matchFieldsContainer.update()
    
    def _reflowQueryFieldsLayout(self):
        """重新排列查询字段布局，填充空白区域"""
        # 直接请求重新计算布局
        self.queryFieldsLayout.update()
        self.queryFieldsContainer.updateGeometry()
        self.queryFieldsContainer.update()
    
    def _removeQueryField(self, widget):
        """移除查询字段"""
        # 查找组件在列表中的索引
        found_index = -1
        for i, field_tuple in enumerate(self.query_fields):
            # 字段元组现在可能是(列选择框, 操作符选择框, 值输入框, 逻辑选择框)
            if len(field_tuple) >= 3 and field_tuple[0].parentWidget() == widget:
                found_index = i
                break
                
        if found_index != -1:
            # 从列表中移除
            self.query_fields.pop(found_index)
            
            # 从布局中移除并删除组件
            widget.deleteLater()
            
            # 如果删除后只剩一个查询字段，需要移除其逻辑选择器（因为只有一个条件不需要逻辑选择器）
            if len(self.query_fields) == 1 and len(self.query_fields[0]) == 4 and self.query_fields[0][3] is not None:
                # 将第一个字段的逻辑选择器设为None
                col_combo, op_combo, value_edit, _ = self.query_fields[0]
                self.query_fields[0] = (col_combo, op_combo, value_edit, None)
            
            # 立即更新布局
            self._reflowQueryFieldsLayout()
            
            # 更新执行按钮状态
            self._updateExecuteButtonState()
            
    def _getCommonColumns(self):
        """获取所有选择的工作表中的共同列，保持第一个工作表中列的原始顺序"""
        if not self.selected_sheets:
            return []
            
        # 获取每个选择的工作表的列
        sheet_columns = []
        first_sheet_columns_ordered = []
        
        # 处理每个选中的工作表
        first_sheet_processed = False
        
        for button in self.selected_sheets:
            if not button.isChecked():
                continue
                
            sheet_name = button.text()
            if sheet_name and sheet_name in self.sheets:
                df = self.sheets[sheet_name]
                if isinstance(df, pd.DataFrame) and not df.empty:
                    # 如果是第一个工作表，记录其列顺序
                    if not first_sheet_processed:
                        first_sheet_columns_ordered = list(df.columns)
                        first_sheet_processed = True
                    
                    # 将工作表的列添加到列集合中
                    sheet_columns.append(set(df.columns))
        
        # 如果没有有效的工作表，返回空列表
        if not sheet_columns:
            return []
            
        # 获取所有工作表的共同列
        common_columns_set = set.intersection(*sheet_columns)
        
        # 按照第一个工作表的列顺序排序共同列
        common_columns_ordered = [col for col in first_sheet_columns_ordered if col in common_columns_set]
        
        return common_columns_ordered

    def _onProcessingModeChanged(self, index):
        """处理模式变化时的处理"""
        # 获取当前模式
        current_mode = self.processingModeCombo.currentText()
        
        # 清空现有的查询和显示字段
        self._clearAllFields()
        
        # 如果已经加载了Excel文件，则重新添加查询和显示字段
        if self.sheets and len(self.selected_sheets) > 0:
            # 添加一个新的查询字段
            self._addQueryField()
            
            # 添加一个新的显示字段
            self._addMatchField()
            
            # 更新执行按钮状态
            self._updateExecuteButtonState()
            
            # 显示模式变化提示
            InfoBar.info(
                title="模式变化",
                content=f"已切换到{current_mode}模式，查询和显示字段已更新",
                parent=self,
                position=InfoBarPosition.TOP,
                duration=3000
            )

    def onResize(self, event):
        """窗口大小变化时的处理"""
        # 调整分割器大小
        if self.splitter:
            width = self.width()
            if width < 800:
                # 小窗口时，左侧占比更大
                self.splitter.setSizes([int(width * 0.6), int(width * 0.4)])
            else:
                # 大窗口时，右侧占比更大
                self.splitter.setSizes([int(width * 0.4), int(width * 0.6)])
        
        # 调用父类的resizeEvent
        super().resizeEvent(event)
        
        # 窗口大小变化时重新平衡左侧三个区域
        if hasattr(self, 'leftScrollContent') and hasattr(self, 'leftScrollLayout'):
            # 获取可用高度
            available_height = self.leftWidget.height() - 40  # 减去按钮区域的高度
            if available_height > 0:
                # 根据窗口高度调整三个区域的高度
                self._adjustLeftPanelSizes(available_height)
    
    def _adjustLeftPanelSizes(self, available_height):
        """根据可用高度调整左侧面板各部分大小"""
        try:
            # 获取三个主要区域
            sheet_section = self.leftScrollLayout.itemAt(0).widget()
            query_section = self.leftScrollLayout.itemAt(1).widget()
            display_section = self.leftScrollLayout.itemAt(2).widget()
            
            # 计算每个区域的高度 - 均分可用高度
            section_height = int(available_height / 3)
            
            # 设置最小高度，确保内容可见
            min_height = 150
            
            # 根据工作表数量和查询字段数量调整区域高度
            sheet_count = len(self.selected_sheets) if hasattr(self, 'selected_sheets') else 0
            query_count = len(self.query_fields) if hasattr(self, 'query_fields') else 0
            match_count = len(self.match_fields) if hasattr(self, 'match_fields') else 0
            
            # 设置最小高度
            sheet_section.setMinimumHeight(min_height)
            query_section.setMinimumHeight(min_height)
            display_section.setMinimumHeight(min_height)
            
            # 根据内容比例适当调整高度（工作表少时可以分配更少空间）
            if sheet_count <= 2 and query_count > 2:
                # 如果工作表少但查询条件多，给查询条件更多空间
                sheet_height = int(section_height * 0.7)
                query_height = int(section_height * 1.5)
                display_height = available_height - sheet_height - query_height
            elif match_count > 5 and query_count <= 2:
                # 如果显示字段很多但查询条件少，给显示字段更多空间
                display_height = int(section_height * 1.5)
                query_height = int(section_height * 0.7)
                sheet_height = available_height - display_height - query_height
            else:
                # 默认均匀分配
                sheet_height = section_height
                query_height = section_height
                display_height = section_height
                
            # 更新区域高度
            sheet_section.setFixedHeight(sheet_height)
            query_section.setFixedHeight(query_height)
            display_section.setFixedHeight(display_height)
            
        except Exception as e:
            # 出错时不阻止程序继续运行
            print(f"调整布局大小时出错: {str(e)}")

    def _alignDataFrameColumns(self, dataframes):
        """对齐多个DataFrame的列，确保可以垂直堆叠
        
        策略:
        1. 找出所有DataFrame中的所有唯一列
        2. 对于每个DataFrame，添加缺失的列并填充NaN
        3. 返回列对齐后的DataFrame列表
        """
        if not dataframes:
            return []
            
        # 收集所有数据框中的所有列
        all_columns = set()
        for df in dataframes:
            all_columns.update(df.columns)
            
        # 对每个数据框添加缺失的列
        aligned_dfs = []
        for df in dataframes:
            # 找出当前数据框缺失的列
            missing_columns = all_columns - set(df.columns)
            
            # 如果有缺失的列，创建一个新的数据框并添加缺失的列
            if missing_columns:
                # 创建一个新的数据框，包含原始列和缺失的列
                new_df = df.copy()
                for col in missing_columns:
                    new_df[col] = pd.NA  # 使用pandas的NA表示缺失值
                aligned_dfs.append(new_df)
            else:
                # 如果没有缺失的列，直接使用原始数据框
                aligned_dfs.append(df)
                
        return aligned_dfs

    def displayResults(self, df):
        """显示查询结果"""
        # 保存结果数据
        self.result_data = df

        # 清空表格
        self.resultTable.clear() # Clear headers too

        # 此时df不应该为空，因为在_processAndDisplayResults中已经检查过了
        # 但再次检查以增加健壮性
        if df.empty:
            self.resultTable.setRowCount(0)
            self.resultTable.setColumnCount(0)
            return

        # 设置表格列数和标题
        columns = list(df.columns)
        col_count = len(columns)
        self.resultTable.setColumnCount(col_count)
        self.resultTable.setHorizontalHeaderLabels(columns)

        # 设置表格行数
        row_count = len(df)
        self.resultTable.setRowCount(row_count)
        
        # 设置表格为不可编辑
        self.resultTable.setEditTriggers(QTableWidget.NoEditTriggers)

        # 填充数据
        # Using itertuples for potentially better performance than iloc in loop
        for row_idx, data_row in enumerate(df.itertuples(index=False, name=None)):
            for col_idx in range(col_count):
                value = data_row[col_idx]
                # Convert value to string for QTableWidgetItem
                # Handle None/NaN gracefully
                if pd.isna(value):
                    item_text = ""
                else:
                    # 保持原始格式
                    item_text = str(value)

                # 创建表格项
                table_item = QTableWidgetItem(item_text)
                
                # 所有单元格默认居中对齐
                table_item.setTextAlignment(Qt.AlignCenter)

                self.resultTable.setItem(row_idx, col_idx, table_item)

        # 显示结果统计
        InfoBar.success(
            title="查询完成",
            content=f"共找到 {row_count} 条匹配记录",
            parent=self,
            position=InfoBarPosition.TOP,
            duration=3000
        )


def main():
    # 启用高DPI支持
    # if hasattr(Qt, 'AA_EnableHighDpiScaling'):
    #     QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    # if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
    #     QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)

    # 创建应用程序
    app = QApplication(sys.argv)

    # 设置应用程序主题
    setTheme(Theme.AUTO)

    # 创建并显示主窗口
    window = ExcelMatchWindow()
    window.show()

    # 运行应用程序
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

# --- END OF FILE main.py ---