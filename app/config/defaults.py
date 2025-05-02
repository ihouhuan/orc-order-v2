"""
默认配置
-------
包含系统的默认配置值。
"""

# 默认配置
DEFAULT_CONFIG = {
    'API': {
        'api_key': '',  # 将从配置文件中读取
        'secret_key': '',  # 将从配置文件中读取
        'timeout': '30',
        'max_retries': '3',
        'retry_delay': '2',
        'api_url': 'https://aip.baidubce.com/rest/2.0/ocr/v1/table'
    },
    'Paths': {
        'input_folder': 'data/input',
        'output_folder': 'data/output',
        'temp_folder': 'data/temp',
        'template_folder': 'templates',
        'processed_record': 'data/processed_files.json'
    },
    'Performance': {
        'max_workers': '4',
        'batch_size': '5',
        'skip_existing': 'true'
    },
    'File': {
        'allowed_extensions': '.jpg,.jpeg,.png,.bmp',
        'excel_extension': '.xlsx',
        'max_file_size_mb': '4'
    },
    'Templates': {
        'purchase_order': '银豹-采购单模板.xls'
    }
} 