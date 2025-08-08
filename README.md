# 📊 Excel 文件批量比较工具

一个功能强大的Excel文件批量比较工具，支持自动文件配对、智能差异检测和详细报告生成。特别适用于需要批量比较大量Excel文件的场景。

## ✨ 主要特性

### 🔄 智能文件配对
- **自动配对**: 根据文件命名规则自动匹配文件对（基础名称-A.xlsx 和 基础名称-B.xlsx）
- **批量处理**: 一次性选择多个文件，自动识别并配对
- **状态显示**: 实时显示文件配对状态和处理进度

### 📈 灵活比较配置
- **自定义行范围**: 支持多种行选择格式
  - 连续范围：`1-100`（比较第1到100行）
  - 离散选择：`1,3,4,9`（比较指定行）
  - 全文件比较：留空则比较所有行
- **智能数据处理**: 自动处理空值、空字符串等特殊情况

### 📊 详细比较报告
- **单文件报告**: 每对文件生成独立的详细比较报告
- **批量汇总报告**: 生成包含所有文件对的汇总分析
- **多维度分析**: 
  - 比较概览（相似度、差异统计）
  - 差异详情（具体位置、差异类型、原值对比）
  - 原始数据（文件A和文件B的完整数据）
  - 统计分析（差异分布、类型统计）

### 🎯 精准差异检测
- **单元格级别**: 精确到每个单元格的差异
- **差异分类**: 
  - 值不同
  - 文件A为空值/文件B为空值
  - 文件A为空字符串/文件B为空字符串
- **相似度计算**: 自动计算文件相似度百分比

### 🖥️ 友好用户界面
- **图形化界面**: 基于Tkinter的直观操作界面
- **实时状态**: 操作状态实时显示，进度一目了然
- **文件管理**: 可视化文件列表，支持添加、清空操作

## 🚀 快速开始

### 环境要求
- Python 3.7+
- 支持的操作系统：Windows、macOS、Linux

## 💻 代码执行命令

### 方法一：一键启动（推荐）

#### macOS/Linux 系统：
```bash
# 进入项目目录
cd /path/to/Excel文件比较工具

# 给启动脚本添加执行权限
chmod +x run.sh

# 运行启动脚本
./run.sh
```

#### Windows 系统：
```cmd
# 进入项目目录
cd C:\path\to\Excel文件比较工具

# 运行启动脚本（如果有Git Bash）
bash run.sh

# 或者手动执行以下命令
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python excel_compare_tool.py
```

### 方法二：手动执行

#### 1. 创建虚拟环境
```bash
# 创建Python虚拟环境
python3 -m venv venv

# 或者使用python命令（Windows）
python -m venv venv
```

#### 2. 激活虚拟环境
```bash
# macOS/Linux
source venv/bin/activate

# Windows (Command Prompt)
venv\Scripts\activate

# Windows (PowerShell)
venv\Scripts\Activate.ps1

# Windows (Git Bash)
source venv/Scripts/activate
```

#### 3. 安装依赖包
```bash
# 安装所有必需的依赖
pip install -r requirements.txt

# 或者逐个安装
pip install pandas>=1.3.0
pip install openpyxl>=3.0.0
pip install matplotlib>=3.3.0
pip install seaborn>=0.11.0
pip install numpy>=1.20.0
```

#### 4. 运行程序
```bash
# 启动Excel比较工具
python excel_compare_tool.py

# 或者使用python3（某些系统）
python3 excel_compare_tool.py
```

### 方法三：直接运行（如果已安装依赖）
```bash
# 如果系统已全局安装所需依赖，可直接运行
python excel_compare_tool.py
```

### 方法四：使用IDE运行
```bash
# 在VS Code中打开项目
code .

# 然后在VS Code中：
# 1. 打开excel_compare_tool.py文件
# 2. 按F5或点击运行按钮
# 3. 选择Python解释器（建议选择venv中的Python）
```

## 🔧 启动脚本说明

项目包含的 `run.sh` 脚本会自动执行以下操作：

```bash
#!/bin/bash
# 1. 检查虚拟环境是否存在，不存在则创建
# 2. 激活虚拟环境
# 3. 安装/更新所有依赖包
# 4. 启动Excel比较工具
# 5. 保持虚拟环境激活状态
```

## 📦 依赖包详解

```txt
pandas>=1.3.0          # 数据处理和分析核心库
openpyxl>=3.0.0         # Excel文件读写支持
matplotlib>=3.3.0       # 图表绘制（预留功能）
seaborn>=0.11.0         # 统计图表美化（预留功能）
numpy>=1.20.0           # 数值计算基础库
```

### 可选依赖（用于增强功能）
```bash
# 如果需要更好的Excel写入性能
pip install xlsxwriter

# 如果需要处理旧版Excel文件
pip install xlrd
```

## 📖 使用指南

### 1. 文件准备
确保您的Excel文件按照以下命名规则：
- 文件A：`基础名称-A.xlsx`
- 文件B：`基础名称-B.xlsx`

例如：
- `test-A.xlsx`
- `test-B.xlsx`

### 2. 操作步骤
1. **启动程序**: 使用上述任一方法启动程序
2. **选择文件**: 点击"📁 选择多个Excel文件"按钮，选择要比较的所有Excel文件
3. **查看配对**: 程序会自动显示文件配对状态
4. **设置范围**: 在"数据比较行序"中输入要比较的行范围（可选）
5. **开始比较**: 点击"🔍 开始批量比较"按钮
6. **查看结果**: 比较完成后，在桌面的"Excel比较结果"文件夹中查看报告

### 3. 行范围格式说明
```bash
# 连续范围
1-100          # 比较第1行到第100行
2-50           # 比较第2行到第50行

# 离散选择
1,3,4,9        # 只比较第1、3、4、9行
1,5,10-20,25   # 比较第1、5行，第10-20行，第25行

# 全部比较
(留空)         # 比较所有行
```

### 4. 报告文件说明
比较完成后会在桌面生成以下文件结构：
桌面/Excel比较结果_YYYYMMDD_HHMMSS/
├── 基础名称1_比较报告.xlsx
├── 基础名称2_比较报告.xlsx
├── ...
└── 批量比较汇总.xlsx

## 🐛 故障排除

### 常见错误及解决方案

#### 1. ModuleNotFoundError
```bash
# 错误信息：ModuleNotFoundError: No module named 'pandas'
# 解决方案：
pip install -r requirements.txt
```

#### 2. 权限错误（macOS/Linux）
```bash
# 错误信息：Permission denied
# 解决方案：
chmod +x run.sh
```

#### 3. Python版本问题
```bash
# 检查Python版本
python --version
python3 --version

# 如果版本低于3.7，请升级Python
```

#### 4. 虚拟环境激活失败
```bash
# Windows PowerShell执行策略问题
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# 然后重新激活虚拟环境
venv\Scripts\Activate.ps1
```

#### 5. Excel文件无法读取
```bash
# 确保文件未被其他程序占用
# 检查文件格式是否为.xlsx或.xls
# 尝试用Excel打开文件确认文件完整性
```

## 📋 报告内容详解

### 单文件比较报告
每个比较报告包含以下工作表：
- **概览**: 基本信息、统计数据、相似度
- **差异详情**: 具体差异位置、类型、原值对比
- **文件A数据**: 完整的文件A数据（带行号）
- **文件B数据**: 完整的文件B数据（带行号）

### 批量汇总报告
汇总报告包含：
- **汇总统计**: 总体比较统计信息
- **文件对概览**: 所有文件对的比较结果概览
- **所有差异详情**: 合并所有文件对的差异信息
- **差异统计分析**: 差异类型分布统计
- **有差异的文件**: 仅显示存在差异的文件列表

## 🔧 高级功能

### 命令行参数（开发中）
```bash
# 未来版本将支持命令行参数
python excel_compare_tool.py --help
python excel_compare_tool.py --batch /path/to/files --output /path/to/results
```

### 配置文件（开发中）
```bash
# 支持配置文件自定义设置
config.json
```

## 🤝 开发者指南

### 项目结构
Excel文件比较工具/
├── excel_compare_tool.py    # 主程序文件
├── simple_excel_compare.py  # 简单版本（单文件比较）
├── requirements.txt         # 依赖包列表
├── run.sh                  # 启动脚本
├── README.md               # 项目说明
└── venv/                   # 虚拟环境目录

### 运行测试
```bash
# 激活虚拟环境后运行
python -m pytest tests/  # 如果有测试文件
```

## 📄 许可证

本项目采用MIT许可证，详情请查看LICENSE文件。

## 🤝 贡献

欢迎提交Issue和Pull Request来帮助改进这个项目！

### 贡献步骤
```bash
# 1. Fork项目
# 2. 创建特性分支
git checkout -b feature/new-feature

# 3. 提交更改
git commit -am 'Add new feature'

# 4. 推送到分支
git push origin feature/new-feature

# 5. 创建Pull Request
```

## 📞 支持

如果您在使用过程中遇到问题，请：
1. 查看本README的故障排除部分
2. 检查程序状态显示区域的错误信息
3. 提交Issue描述具体问题

### 联系方式
- 📧 Email: [您的邮箱]
- 🐛 Issues: [项目Issues页面]
- 📖 Wiki: [项目Wiki页面]

---

**享受高效的Excel文件比较体验！** 🎉

## 🔖 版本历史

- **v1.0.0**: 初始版本，支持基本的批量文件比较功能
- **v1.1.0**: 添加了智能文件配对和详细报告生成
- **v1.2.0**: 优化了用户界面和错误处理机制

---

*最后更新时间：2025年8月8日*