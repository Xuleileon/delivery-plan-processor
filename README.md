# Excel到货计划处理工具

## 项目简介
这是一个专门用于处理Excel格式到货计划的自动化工具。该工具可以预处理Excel文件、转换数据格式并合并SKU信息，支持本地运行和云函数部署。同时提供了图形界面，方便用户操作。

## 功能特性
- Excel文件预处理（公式计算、工作表清理）
- 到货计划格式转换和标准化
- SKU数据合并和去重
- 支持本地运行和云函数部署
- 完整的日志记录
- 可配置的处理流程
- 直观的图形用户界面
- 支持从飞书文档导入数据
- 实时处理状态反馈
- 自定义输出目录

## 系统要求
- Python 3.8+
- Windows操作系统（因使用了win32com组件）
- 安装有Microsoft Excel

## 安装说明
1. 克隆项目到本地
2. 创建并激活虚拟环境：
```bash
python -m venv venv
.\venv\Scripts\activate
```
3. 安装依赖：
```bash
pip install -r requirements.txt
```

## 使用方法
### 图形界面模式
1. 运行GUI应用：
```bash
python gui_app.py
```
2. 在界面中选择：
   - 选择输入文件或配置飞书导入
   - 设置输出目录（可选）
   - 选择配置文件（可选）
   - 点击处理按钮开始处理

### 命令行模式
```bash
python main.py --input <输入文件路径> --output <输出目录> --config <配置文件路径>
```

## 配置说明
配置文件（config.yaml）支持以下选项：
- 输入/输出目录设置
- 数据转换规则
- 飞书集成配置
- 日志级别设置

## 主要依赖
- pandas: 数据处理
- openpyxl: Excel文件操作
- PyYAML: 配置文件解析
- tkinter: 图形界面
- requests: 网络请求

## 开发指南
- 使用pre-commit hooks进行代码质量控制
- 运行测试：`pytest`
- 代码格式化：`black .`
- 类型检查：`mypy .`

## 许可证
MIT License

## 联系方式
如有问题或建议，请提交Issue或Pull Request。