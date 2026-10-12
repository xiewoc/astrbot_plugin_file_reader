# astrbot_plugin_file_reader

## 使用

先下载`requirements.txt`里面的库（`pip install -r requirements.txt`）

`Linux` 用户可能需要安装 `libmagic`
`sudo apt-get install libmagic1`

**上传文件不会主动调用 LLM。** 插件会在收到附件时读取文本并暂存在内存中；只有后续问题明确提到文件、文档或文件名时，才将相应内容加入本轮 LLM 请求。

### 文件缓存与权限

- 群内成员可共同分析本群附件；不同群、不同平台适配器实例及私聊之间相互隔离。
- 使用 `/readfiles` 查看当前会话已读取的文件；多个文件未指定文件名时，机器人会提示用户澄清。
- 相同会话中按附件内容去重，默认最多缓存 10 个文件、60 分钟；全局缓存正文总计不超过 400,000 字符。
- 单文件大小上限为 20 MB，单文件最多提取 40,000 字符。文本超限会标明内容不完整。
- 插件不另存原文件；附件下载和清理由 AstrBot 本身管理。插件重启后内存缓存会丢失。
- 文件内容以不可信资料的形式注入用户消息，不执行文件中的指令或脚本。扫描件、图片文字不提供 OCR。
- 避免多机器人重复发送回执，上传后默认不主动发确认消息，可使用 `/readfiles` 检查是否已读取。

**兼容说明**：使用 AstrBot 通用 `File.get_file()` 接口；QQ 官方机器人和个人微信等平台的实际附件接收情况需要各自测试。多个机器人即使处于同一个群，只要平台实例 ID 不同，就不会自动共享文件缓存，以免误合并不同群。

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
