# Excel 文件合并工具

一个基于 Web 的 Excel 文件合并工具，支持多种合并方式和格式转换。

## 功能特点

### 文件支持
- 支持 Excel 文件 (.xlsx, .xls) 和 CSV 文件
- 支持多文件同时上传
- 文件大小限制：16MB
- 支持显示每个文件的行数和总行数

### 合并方式
1. **按原表名分sheet**
   - 每个文件保持原有内容
   - 以文件名作为新的 sheet 名
   - 自动处理重复的 sheet 名
   - 支持自定义编辑每个sheet的名称

2. **相同sheet名合并**
   - 将具有相同 sheet 名的内容合并到一起
   - 保留所有 sheet 的结构

3. **合并成一个sheet**
   - 将所有数据合并到单个 sheet 中
   - 可选功能：
     - 去除重复表头（可设置从第几行开始）
     - 添加来源列（可设置插入位置）
     - 支持自定义来源列名称
   - 适合需要统一处理数据的场景

### 输出选项
- 支持输出为 Excel (.xlsx) 或 CSV 格式
- 自定义输出文件名
- 文件名自动包含合并方式和日期信息

### 预览功能
- 动态预览功能：
  - 按原表名分sheet时：预览生成的sheet名称
  - 合并成一个sheet时：预览来源列名称（如果启用）
- 显示每个文件的行数和总行数
- 支持编辑sheet名称或来源列名称
- 显示文件大小和格式信息

## 使用方法

1. **选择文件**
   - 点击"选择文件"或拖拽文件到指定区域
   - 可以选择多个文件
   - 支持选择/取消选择单个文件
   - 支持一键删除选中文件

2. **设置合并选项**
   - 选择合并方式：
     - 分sheet：每个文件生成独立的sheet
     - 同名合并：相同名称的sheet合并到一起
     - 单sheet合并：所有数据合并到一个sheet
   - 单sheet合并的额外选项：
     - 去除重复表头：可设置从第几行开始
     - 添加来源列：可设置插入位置和自定义名称
   - 选择输出格式（Excel/CSV）
   - 可选择自定义输出文件名

3. **预览和编辑**
   - 查看每个文件的行数
   - 根据合并方式预览不同内容：
     - sheet名称预览
     - 来源列名称预览
   - 支持编辑预览的名称
   - 确认输出文件名

4. **执行合并**
   - 点击"开始合并"按钮
   - 等待处理完成
   - 下载合并后的文件

## 技术栈

- 后端：Python + Flask
- 前端：HTML + CSS + JavaScript
- 数据处理：Pandas

## 安装和运行

1. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```

2. 运行应用：
   ```bash
   python app.py
   ```

3. 访问：
   打开浏览器访问 `http://localhost:5000`

## 注意事项

- 建议在合并前备份原始文件
- 大文件处理可能需要较长时间
- 确保文件格式正确且未损坏
- 建议保持稳定的网络连接

## 文件命名规则

- 默认文件名格式：`合并文件_[合并方式]_[日期].xlsx`
- 示例：`合并文件_单表合并_20240315.xlsx`
- 可以自定义文件名，系统会自动添加正确的扩展名

## 贡献

欢迎提交 Issue 和 Pull Request 来帮助改进这个工具。

## 许可证

MIT License 