# -*- coding: utf-8 -*-
"""
UI组件模块 - 定义各种UI组件和布局
"""

from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QTableWidgetItem,
                               QSplitter, QScrollArea, QFrame, QLabel, QFormLayout)
from PySide6.QtCore import Qt, Signal, QSize, QRect, QMargins
from PySide6.QtGui import QIcon, QFont

from qfluentwidgets import (SubtitleLabel, PushButton, ComboBox, TableWidget,
                            ToolButton, FluentIcon, LineEdit, SmoothScrollArea, 
                            FlowLayout, PrimaryPushButton, TogglePushButton)


class QueryFieldWidget(QWidget):
    """查询字段组件"""
    
    removeRequested = Signal(object)  # 删除请求信号
    
    def __init__(self, parent=None, columns=None):
        super().__init__(parent)
        self.columns = columns or []
        self.initUI()
    
    def initUI(self):
        """初始化UI"""
        # 创建水平布局
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)
        
        # 列选择下拉框
        self.columnCombo = ComboBox(self)
        self.columnCombo.setMinimumWidth(120)
        self.columnCombo.setMaximumWidth(200)
        self.updateColumns(self.columns)
        
        # 操作符下拉框
        self.operatorCombo = ComboBox(self)
        self.operatorCombo.setMinimumWidth(80)
        self.operatorCombo.setMaximumWidth(100)
        self.operatorCombo.addItems(["=", "!=", "<", "<=", ">", ">=", "包含", "不包含", "开头是", "结尾是", "为空", "不为空"])
        
        # 值输入框
        self.valueEdit = LineEdit(self)
        self.valueEdit.setMinimumWidth(100)
        self.valueEdit.setPlaceholderText("输入查询值")
        
        # 逻辑运算符下拉框
        self.logicCombo = ComboBox(self)
        self.logicCombo.setMinimumWidth(60)
        self.logicCombo.setMaximumWidth(80)
        self.logicCombo.addItems(["AND", "OR"])
        
        # 删除按钮
        self.removeButton = ToolButton(FluentIcon.REMOVE, self)
        self.removeButton.setToolTip("删除此查询条件")
        self.removeButton.clicked.connect(lambda: self.removeRequested.emit(self))
        
        # 添加组件到布局
        layout.addWidget(self.columnCombo)
        layout.addWidget(self.operatorCombo)
        layout.addWidget(self.valueEdit, 1)  # 1表示可伸缩
        layout.addWidget(self.logicCombo)
        layout.addWidget(self.removeButton)
        
        # 设置为空和不为空操作符时禁用值输入框
        self.operatorCombo.currentIndexChanged.connect(self.updateValueEditState)
    
    def updateValueEditState(self, index):
        """更新值输入框状态"""
        operator = self.operatorCombo.currentText()
        if operator in ["为空", "不为空"]:
            self.valueEdit.setEnabled(False)
            self.valueEdit.setPlaceholderText("无需输入值")
        else:
            self.valueEdit.setEnabled(True)
            self.valueEdit.setPlaceholderText("输入查询值")
    
    def updateColumns(self, columns):
        """更新列选择下拉框"""
        self.columnCombo.clear()
        if columns:
            self.columnCombo.addItems(columns)
    
    def getQueryCondition(self):
        """获取查询条件"""
        return (
            self.columnCombo.currentText(),
            self.operatorCombo.currentText(),
            self.valueEdit.text(),
            self.logicCombo.currentText()
        )


class MatchFieldWidget(QWidget):
    """匹配字段组件"""
    
    removeRequested = Signal(object)  # 删除请求信号
    
    def __init__(self, parent=None, columns=None):
        super().__init__(parent)
        self.columns = columns or []
        self.initUI()
    
    def initUI(self):
        """初始化UI"""
        # 创建水平布局
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)
        
        # 列选择下拉框
        self.columnCombo = ComboBox(self)
        self.columnCombo.setMinimumWidth(120)
        self.columnCombo.setMaximumWidth(200)
        self.updateColumns(self.columns)
        
        # 自定义标题输入框
        self.titleEdit = LineEdit(self)
        self.titleEdit.setMinimumWidth(100)
        self.titleEdit.setPlaceholderText("自定义显示标题(可选)")
        
        # 删除按钮
        self.removeButton = ToolButton(FluentIcon.REMOVE, self)
        self.removeButton.setToolTip("删除此显示字段")
        self.removeButton.clicked.connect(lambda: self.removeRequested.emit(self))
        
        # 添加组件到布局
        layout.addWidget(self.columnCombo)
        layout.addWidget(self.titleEdit, 1)  # 1表示可伸缩
        layout.addWidget(self.removeButton)
    
    def updateColumns(self, columns):
        """更新列选择下拉框"""
        self.columnCombo.clear()
        if columns:
            self.columnCombo.addItems(columns)
    
    def getMatchField(self):
        """获取匹配字段"""
        return (
            self.columnCombo.currentText(),
            self.titleEdit.text()
        )


class MergeKeyDialog(QWidget):
    """合并键选择对话框"""
    
    def __init__(self, parent=None, common_columns=None):
        super().__init__(parent)
        self.common_columns = common_columns or []
        self.selected_key = None
        self.merge_how = 'outer'  # 默认合并方式
        self.initUI()
    
    def initUI(self):
        """初始化UI"""
        # 创建垂直布局
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # 标题
        titleLabel = SubtitleLabel("选择合并键和合并方式", self)
        layout.addWidget(titleLabel)
        
        # 说明文本
        descLabel = QLabel("请选择用于关联不同工作表的列，以及合并方式：", self)
        layout.addWidget(descLabel)
        
        # 合并键选择
        keyLayout = QFormLayout()
        keyLayout.setContentsMargins(0, 0, 0, 0)
        keyLayout.setSpacing(5)
        
        self.keyCombo = ComboBox(self)
        self.keyCombo.setMinimumWidth(200)
        if self.common_columns:
            self.keyCombo.addItems(self.common_columns)
            self.selected_key = self.common_columns[0]  # 默认选择第一个
        
        keyLayout.addRow("合并键:", self.keyCombo)
        layout.addLayout(keyLayout)
        
        # 合并方式选择
        mergeLayout = QVBoxLayout()
        mergeLayout.setContentsMargins(0, 0, 0, 0)
        mergeLayout.setSpacing(5)
        
        mergeLabel = QLabel("合并方式:", self)
        mergeLayout.addWidget(mergeLabel)
        
        # 外连接按钮
        self.outerButton = PushButton("外连接 (保留所有数据)", self)
        self.outerButton.setCheckable(True)
        self.outerButton.setChecked(True)  # 默认选中
        self.outerButton.clicked.connect(self.selectOuter)
        mergeLayout.addWidget(self.outerButton)
        
        # 内连接按钮
        self.innerButton = PushButton("内连接 (仅保留匹配的数据)", self)
        self.innerButton.setCheckable(True)
        self.innerButton.clicked.connect(self.selectInner)
        mergeLayout.addWidget(self.innerButton)
        
        # 左连接按钮
        self.leftButton = PushButton("左连接 (保留左侧工作表的所有数据)", self)
        self.leftButton.setCheckable(True)
        self.leftButton.clicked.connect(self.selectLeft)
        mergeLayout.addWidget(self.leftButton)
        
        layout.addLayout(mergeLayout)
        
        # 确认按钮
        self.confirmButton = PrimaryPushButton("确认", self)
        layout.addWidget(self.confirmButton, 0, Qt.AlignRight)
        
        # 连接信号
        self.keyCombo.currentIndexChanged.connect(self.updateSelectedKey)
    
    def updateSelectedKey(self, index):
        """更新选中的合并键"""
        if index >= 0 and index < len(self.common_columns):
            self.selected_key = self.common_columns[index]
    
    def selectOuter(self):
        """选择外连接"""
        self.merge_how = 'outer'
        self.outerButton.setChecked(True)
        self.innerButton.setChecked(False)
        self.leftButton.setChecked(False)
    
    def selectInner(self):
        """选择内连接"""
        self.merge_how = 'inner'
        self.outerButton.setChecked(False)
        self.innerButton.setChecked(True)
        self.leftButton.setChecked(False)
    
    def selectLeft(self):
        """选择左连接"""
        self.merge_how = 'left'
        self.outerButton.setChecked(False)
        self.innerButton.setChecked(False)
        self.leftButton.setChecked(True)
    
    def getSelectedKey(self):
        """获取选中的合并键"""
        return self.selected_key
    
    def getMergeHow(self):
        """获取合并方式"""
        return self.merge_how