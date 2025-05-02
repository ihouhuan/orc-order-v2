# OCR订单处理系统 v2.0

基于百度OCR API的订单处理系统，用于识别采购订单图片并生成Excel采购单。

## 功能特点

- **图像OCR识别**：支持对采购单图片进行OCR识别并生成Excel文件
- **Excel数据处理**：读取OCR识别的Excel文件并提取商品信息
- **采购单生成**：按照模板格式生成标准采购单Excel文件
- **采购单合并**：支持多个采购单合并为一个总单
- **批量处理**：支持批量处理多张图片
- **图形界面**：提供简洁直观的图形界面，方便操作
- **命令行支持**：支持命令行方式调用，便于自动化处理

## 系统架构

### 目录结构

```
orc-order-v2/
│
├── app/                  # 应用主目录
│   ├── config/           # 配置目录
│   │   ├── settings.py   # 配置管理
│   │   └── defaults.py   # 默认配置值
│   │
│   ├── core/             # 核心功能
│   │   ├── ocr/          # OCR相关功能
│   │   │   ├── baidu_ocr.py      # 百度OCR接口
│   │   │   └── table_ocr.py      # 表格OCR处理
│   │   │
│   │   ├── excel/        # Excel处理
│   │   │   ├── processor.py      # Excel处理核心
│   │   │   ├── merger.py         # 订单合并功能
│   │   │   └── converter.py      # 单位转换与规格处理
│   │   │
│   │   └── utils/        # 工具函数
│   │       ├── file_utils.py     # 文件操作工具
│   │       ├── log_utils.py      # 日志工具
│   │       └── string_utils.py   # 字符串处理工具
│   │
│   └── services/         # 业务服务
│       ├── ocr_service.py        # OCR服务
│       └── order_service.py      # 订单处理服务
│
├── data/                 # 数据目录
│   ├── input/            # 输入文件
│   ├── output/           # 输出文件
│   └── temp/             # 临时文件
│
├── logs/                 # 日志目录
│
├── templates/            # 模板文件
│   └── 银豹-采购单模板.xls       # 采购单模板
│
├── 启动器.py               # 图形界面启动器
├── run.py                # 命令行入口
├── config.ini            # 配置文件
└── requirements.txt      # 依赖包列表
```

### 主要模块说明

- **配置模块**：统一管理系统配置，支持默认值和配置文件读取
- **OCR模块**：调用百度OCR API进行表格识别，生成Excel文件
- **Excel处理模块**：读取OCR生成的Excel文件，提取商品信息
- **单位转换模块**：处理商品规格和单位转换
- **订单合并模块**：合并多个采购单为一个总单
- **文件工具模块**：处理文件读写、路径管理等
- **启动器**：提供图形界面操作

## 使用方法

### 环境准备

1. 安装Python 3.6+
2. 安装依赖包：
   ```
   pip install -r requirements.txt
   ```
3. 配置百度OCR API密钥：
   - 在`config.ini`中填写您的API密钥和Secret密钥

### 图形界面使用

1. 运行启动器：
   ```
   python 启动器.py
   ```
2. 使用界面上的功能按钮进行操作：
   - **OCR图像识别**：批量处理`data/input`目录下的图片
   - **处理单个图片**：选择并处理单个图片
   - **处理Excel文件**：处理OCR识别后的Excel文件，生成采购单
   - **合并采购单**：合并所有生成的采购单
   - **完整处理流程**：按顺序执行所有处理步骤
   - **整理项目文件**：整理文件到规范目录结构

### 命令行使用

系统提供命令行方式调用，便于集成到自动化流程中：

```bash
# OCR识别
python run.py ocr [--input 图片路径] [--batch]

# Excel处理
python run.py excel [--input Excel文件路径]

# 订单合并
python run.py merge [--input 采购单文件路径列表]

# 完整流程
python run.py pipeline
```

## 文件处理流程

1. **OCR识别处理**：
   - 读取`data/input`目录下的图片文件
   - 调用百度OCR API进行表格识别
   - 保存识别结果为Excel文件到`data/output`目录

2. **Excel处理**：
   - 读取OCR识别生成的Excel文件
   - 提取商品信息（条码、名称、规格、单价、数量等）
   - 按照采购单模板格式生成标准采购单Excel文件
   - 输出文件命名为"采购单_原文件名.xls"

3. **采购单合并**：
   - 读取所有采购单Excel文件
   - 合并相同商品的数量
   - 生成总采购单

## 配置说明

系统配置文件`config.ini`包含以下主要配置：

```ini
[API]
api_key = 您的百度API Key
secret_key = 您的百度Secret Key
timeout = 30
max_retries = 3
retry_delay = 2
api_url = https://aip.baidubce.com/rest/2.0/ocr/v1/table

[Paths]
input_folder = data/input
output_folder = data/output
temp_folder = data/temp

[Performance]
max_workers = 4
batch_size = 5
skip_existing = true

[File]
allowed_extensions = .jpg,.jpeg,.png,.bmp
excel_extension = .xlsx
max_file_size_mb = 4
```

## 注意事项

1. 系统依赖百度OCR API，使用前请确保已配置正确的API密钥
2. 图片质量会影响OCR识别结果，建议使用清晰的原始图片
3. 处理大量图片时可能会受到API调用频率限制
4. 所有处理好的文件会保存在`data/output`目录中

## 错误排查

- **OCR识别失败**：检查API密钥是否正确，图片是否符合要求
- **Excel处理失败**：检查OCR识别结果是否包含必要的列（条码、数量、单价等）
- **模板填充错误**：确保模板文件存在且格式正确

## 开发说明

如需进行二次开发或扩展功能，请参考以下说明：

1. 核心逻辑位于`app/core`目录
2. 添加新功能建议遵循已有的模块化结构
3. 使用`app/services`目录中的服务类调用核心功能
4. 日志记录已集成到各模块，便于调试

## 许可证

MIT License

## 联系方式

如有问题，请提交Issue或联系开发者。 