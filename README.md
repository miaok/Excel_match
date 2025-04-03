# Excel多条件多表查询匹配工具

这是一个使用PySide6和QFluentWidgets开发的Excel多条件多表查询匹配工具，类似于Excel中filter加vstack函数实现的多条件多表查询匹配功能。

## 功能特点

- 加载Excel文件并读取所有工作簿
- 根据用户选择的工作簿列标题进行查询和匹配
- 支持添加多个查询字段（通过下拉列表选择）
- 支持添加多个匹配字段（通过下拉列表选择）
- 查询和匹配字段支持增加删除
- 表格显示匹配结果
- 现代化的Fluent设计界面

## 安装依赖

```bash
pip install -r requirements.txt
```

## 运行程序

```bash
python main.py
```

## 使用说明

1. 点击「选择文件」按钮，选择要加载的Excel文件
2. 从下拉列表中选择要操作的工作表
3. 点击「添加查询字段」按钮，选择要查询的列和输入查询值
4. 点击「添加匹配字段」按钮，选择要显示的列
5. 点击「执行查询」按钮，查看匹配结果

## 技术栈

- PySide6：Qt for Python
- PySide6-Fluent-Widgets：基于PySide6的Fluent Design风格组件库
- pandas：用于数据处理和分析的Python库
- openpyxl/xlrd：用于读取Excel文件的Python库