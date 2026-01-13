# Excel转Markdown工具

一个简单实用的Python工具，将Excel文件转换为Markdown表格格式。提供命令行和图形界面两种使用方式。

## 功能特点

- 支持多种Excel格式（.xlsx, .xls）
- 自动转换所有工作表或指定工作表
- 保留表格结构和数据
- 提供命令行和图形界面两种方式
- 支持自定义输出路径
- 实时预览转换结果（GUI版本）
- 多选工作表支持（GUI版本）

## 安装

1. 克隆或下载此项目

2. 安装依赖：
```bash
pip install -r requirements.txt
```

## 使用方法

### 方式一：独立可执行文件（Windows，最简单）

如果你使用 Windows 系统，可以直接使用打包好的可执行文件：

1. 双击 `dist/Excel转Markdown工具.exe` 即可启动程序
2. 无需安装 Python 环境或任何依赖

**优点：**
- ✅ 开箱即用，无需安装任何依赖
- ✅ 单文件，可以单独分发
- ✅ 包含完整的 Python 运行环境

### 方式二：图形界面（推荐，需要 Python 环境）

**Windows用户：**
双击运行 `run_gui.bat` 或在命令行执行：
```bash
python excel2md_gui.py
```

**Mac/Linux用户：**
```bash
python3 excel2md_gui.py
# 或
chmod +x run_gui.sh
./run_gui.sh
```

**GUI功能：**
- 📁 浏览并选择Excel文件
- 📋 多选工作表（支持全选/清空）
- 👁️ 实时预览转换结果
- 💾 自定义输出路径和文件名
- 📊 自动生成Markdown表格

**GUI使用步骤：**
1. 点击"浏览"按钮选择要转换的Excel文件
2. 在工作表列表中选择需要转换的Sheet（可多选）
3. （可选）修改输出文件路径和名称
4. 点击"预览"查看转换结果
5. 点击"转换并保存"生成Markdown文件

### 方式二：命令行

**基本用法：**
```bash
python excel2md.py data.xlsx
```
这将在同一目录下生成 `data.md` 文件。

**指定输出文件：**
```bash
python excel2md.py data.xlsx -o output.md
```

**转换指定工作表：**
```bash
python excel2md.py data.xlsx -s Sheet1
```

**查看帮助：**
```bash
python excel2md.py -h
```

## 项目文件

```
Excel2Md/
├── dist/                # 打包输出目录
│   └── Excel转Markdown工具.exe  # Windows独立可执行文件
├── excel2md.py          # 命令行版本主程序
├── excel2md_gui.py      # 图形界面版本
├── create_example.py    # 创建示例Excel文件
├── test.py              # 测试脚本
├── run_gui.bat          # Windows启动脚本
├── run_gui.sh           # Linux/Mac启动脚本
├── requirements.txt     # 依赖列表
└── README.md            # 使用说明
```

## 示例

假设有一个 `sales.xlsx` 文件，包含以下数据：

| 产品 | 数量 | 单价 |
|------|------|------|
| 苹果 | 100  | 5.5  |
| 香蕉 | 200  | 3.2  |

运行转换命令后，将生成对应的Markdown表格：

```markdown
## 产品销售

| 产品 | 数量 | 单价 |
|---|---|---|
| 苹果 | 100 | 5.5 |
| 香蕉 | 200 | 3.2 |
```

## 快速测试

1. 创建示例Excel文件：
```bash
python create_example.py
```

2. 运行测试：
```bash
python test.py
```

3. 或使用GUI界面：
```bash
python excel2md_gui.py
```

## 注意事项

- 空单元格会被转换为空字符串
- 支持多个工作表的Excel文件
- 确保Excel文件路径正确且可访问
- GUI版本需要操作系统支持图形界面

## 依赖库

- pandas: 数据处理
- openpyxl: 读写.xlsx文件
- xlrd: 读取.xls文件
- tkinter: 图形界面（Python内置）

## 许可证

MIT License
