# astrbot_plugin_file_reader

## 使用

先下载`requirements.txt`里面的库（`pip install -r requirements.txt`）

`Linux` 用户可能需要安装 `libmagic`
`sudo apt-get install libmagic1`

**在使用时发送文件并不会直接呼起llm，而是将文件内容加入prompt里面，在发送文字内容时就附带发送了**

支持了更多的后缀名，按需自取（不用的自己注释掉），字典如下：

 ```
SUPPORTED_EXTENSIONS: Dict[str, str] = {
    # 文档格式
    'pdf': 'read_pdf_to_text',
    'docx': 'read_docx_to_text',
    'doc': 'read_docx_to_text',
    'rtf': 'read_txt_to_text',
    'odt': 'read_txt_to_text',
    
    # 电子表格
    'xlsx': 'read_excel_to_text',
    'xls': 'read_excel_to_text',
    'ods': 'read_excel_to_text',
    'csv': 'read_csv_to_text', 
    
    # 演示文稿
    'pptx': 'read_pptx_to_text',
    'ppt': 'read_pptx_to_text',
    'odp': 'read_pptx_to_text',
    
    # 编程语言源代码
    'py': 'read_txt_to_text',
    'java': 'read_txt_to_text',
    'cpp': 'read_txt_to_text',
    'c': 'read_txt_to_text',
    'h': 'read_txt_to_text',
    'hpp': 'read_txt_to_text',
    'cs': 'read_txt_to_text',
    'js': 'read_txt_to_text',
    'ts': 'read_txt_to_text',
    'php': 'read_txt_to_text',
    'rb': 'read_txt_to_text',
    'go': 'read_txt_to_text',
    'rs': 'read_txt_to_text',
    'swift': 'read_txt_to_text',
    'kt': 'read_txt_to_text',
    'scala': 'read_txt_to_text',
    'sh': 'read_txt_to_text',
    'bash': 'read_txt_to_text',
    'ps1': 'read_txt_to_text',
    'bat': 'read_txt_to_text',
    'cmd': 'read_txt_to_text',
    'vbs': 'read_txt_to_text',
    
    # 标记语言
    'html': 'read_txt_to_text',
    'htm': 'read_txt_to_text',
    'xml': 'read_txt_to_text',
    'json': 'read_txt_to_text',
    'yaml': 'read_txt_to_text',
    'yml': 'read_txt_to_text',
    'md': 'read_txt_to_text',
    'markdown': 'read_txt_to_text',
    
    # 配置文件
    'ini': 'read_txt_to_text',
    'cfg': 'read_txt_to_text',
    'conf': 'read_txt_to_text',
    'properties': 'read_txt_to_text',
    'env': 'read_txt_to_text',
    
    # 数据库/查询
    'sql': 'read_txt_to_text',
    
    # 其他文本格式
    'txt': 'read_txt_to_text',
    'log': 'read_txt_to_text',
    '': 'read_txt_to_text',  # 无扩展名文件
    
    # 构建/项目文件
    'toml': 'read_txt_to_text',
    'lock': 'read_txt_to_text',
    'gitignore': 'read_txt_to_text',
    
    # 网络相关
    'url': 'read_txt_to_text',
    'webloc': 'read_txt_to_text',
}
 ```

**What's New?**

让DS优化了一下