# WordHelper - An Office VSTO Addin for Word

## 功能

- 文档内部变量管理

## Issue

- [ ] VariableControl.DataGrid: 自然排序
- [ ] Ribbon Group: 缩小窗口时显示的图标

## TODO

### 插件功能

- 错误检测：
  - 检查整篇文档是否存在“错误！”情况的配对；
    - 使用“搜索”功能实现？
  - 检查是否存在没有题注的图表：
    - 类似搜索列表结果显示？
    - 能定位？
- 文本编辑：
  - [x] 清除选中段落空行（用户自行用Ctrl+A选全文）；
  - [x] 清除选中段落每行结尾空白字符；
  - [x] 清除选中段落每行结尾的换行，拼接为一个段落；
  - ~~选中段落的全角半角标点符号转换；（两个按钮，半角/全角）~~
  - 选中段落中英文格式化：自动在单个英文字母或单词两边添加/删除空格；
  - 选中段落半角符号格式化：自动在每个半角符号右边留一个空格，左边不留；
  - 选中段落末尾批量增加符号（比如句号，分号等）；
  - [x] 批量转换转换段落中的软回车（由快捷键`Shift+Enter`输入的换行）为硬回车；
- 表格处理：
  - [x] 单元格快速行/列拆分
  - [ ] 单元格扩散复制（将第一个单元格内容复制到选中的其它所有单元格）
- 书签管理：
  - 至少能批量删除
- 文档变量管理：
    - [x] 显示当前变量；
    - [x] 编辑变量；
    - [x] 从 Excel 文件批量导入变量；
- 未知分类：
  - 同时复制标题和编号？
  - 快速引用标题？同时显示编号与标题？

### 插件外观

- 完善Ribbon界面的图标，尽量使用 Office 内置
- 添加 WinForm 界面部分的图标
