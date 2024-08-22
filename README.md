# 邮编区域匹配脚本

该脚本用于从 Excel 文件中读取邮编数据，并根据邮编的区域信息为每个邮编添加一个对应的区域列。

## 使用说明

### 准备工作

1. **创建虚拟环境**（如果尚未创建）：
    ```bash
    python3 -m venv myenv
    ```

2. **激活虚拟环境**：
    ```bash
    source myenv/bin/activate  # 在 macOS 或 Linux 上
    ```

3. **安装所需的 Python 包**：
    确保你在虚拟环境中安装了 `pandas` 和 `openpyxl`：
    ```bash
    pip install pandas openpyxl
    ```

### 准备 Excel 文件

1. **创建 Excel 文件**：在代码的目录下放置一个名为 `邮编.xlsx` 的 Excel 文件。确保 Excel 文件中包含一列名为 `收件人编码` 的邮编数据。

2. **修改文件名**（如有需要）：如果你的 Excel 文件名称不同，请在脚本中将 `input_file` 和 `output_file` 变量的值修改为你的文件名。

### 运行脚本

1. **确保虚拟环境已激活**：
    ```bash
    source myenv/bin/activate  # 在 macOS 或 Linux 上
    # 或
    myenv\Scripts\activate     # 在 Windows 上
    ```

2. **运行脚本**：
    ```bash
    python3 postcodeToArea.py
    ```

### 脚本功能

- **`find_area(postal_code)`**：根据邮编查找对应的区域。
- **`process_excel(input_file, output_file)`**：从指定的 Excel 文件读取邮编数据，处理每个邮编并添加区域列，然后将结果保存到新的 Excel 文件中。

### 示例

如果你的输入文件名为 `postal_codes.xlsx`，输出文件名为 `postal_codes_with_areas.xlsx`，可以将脚本中的 `input_file` 和 `output_file` 修改如下：
```python
input_file = 'postal_codes.xlsx'
output_file = 'postal_codes_with_areas.xlsx'