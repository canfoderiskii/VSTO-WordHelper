# WordHelper - An Office VSTO Addin for Word

## 功能

- 文档内部变量管理

## Issue

- [ ] VariableControl: 长度应可以与 DataGrid 一同增长到某一界限。 需要获取当前 Word 界面长度？
- [ ] VariableControl.DataGrid: 显示上下拉的侧边栏
- [ ] VariableControl.DataGrid: 自然排序
- [x] Ribbon: 显示变量的开关状态似乎反了
- [ ] Edit.TrimTrailing: 总是会遗留一个空格

## TODO

### 插件功能

- 错误检测：
  - 检查整篇文档是否存在“错误！”情况的配对；
    - 使用“搜索”功能实现？
  - 检查是否存在没有题注的图表：
    - 类似搜索列表结果显示？
    - 能定位？
- 文本编辑：
  - 清除选中段落空行（用户自行用Ctrl+A选全文）；
  - 清除选中段落每行结尾或开头空白字符；
  - 清除选中段落每行结尾的换行，拼接为一个段落；
- 文本格式化：
  - 选中段落的全角半角标点符号转换；（两个按钮，半角/全角）
  - 选中段落中英文格式化：自动在单个英文字母或单词两边添加/删除空格；
  - 选中段落半角符号格式化：自动在每个半角符号右边留一个空格，左边不留；
  - 选中段落末尾批量增加符号（比如句号，分号等）
- 书签管理：
  - 至少能批量删除
- 未知分类：
  - 同时复制标题和编号？
  - 快速引用标题？同时显示编号与标题？

### 插件外观

- 完善Ribbon界面的图标，尽量使用 Office 内置
- 添加 WinForm 界面部分的图标
