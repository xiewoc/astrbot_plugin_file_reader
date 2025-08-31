from astrbot.api.event import filter, AstrMessageEvent          # pyright: ignore[reportMissingImports]
from astrbot.api.star import Context, Star, register            # pyright: ignore[reportMissingImports]
from astrbot.api.provider import ProviderRequest                # pyright: ignore[reportMissingImports]
import astrbot.api.message_components as Comp                   # pyright: ignore[reportMissingImports] 
from astrbot.api.all import *                                   # pyright: ignore[reportMissingImports] 
from astrbot.api import logger                                  # pyright: ignore[reportMissingImports]

import os
from pdfminer.high_level import extract_text
import docx2txt
import pandas as pd
from docx import Document
from pptx import Presentation
from typing import Dict, Optional
import chardet
from pathlib import Path

# 使用字典存储支持的文件类型和对应的处理函数
SUPPORTED_EXTENSIONS: Dict[str, str] = {
    # 文档格式
    "pdf": "read_pdf_to_text",
    "docx": "read_docx_to_text",
    "doc": "read_docx_to_text",
    "rtf": "read_txt_to_text",
    "odt": "read_txt_to_text",

    # 电子表格
    "xlsx": "read_excel_to_text",
    "xls": "read_excel_to_text",
    "ods": "read_excel_to_text",
    "csv": "read_csv_to_text",

    # 演示文稿
    "pptx": "read_pptx_to_text",
    "ppt": "read_pptx_to_text",
    "odp": "read_pptx_to_text",

    # 编程语言源代码
    "py": "read_txt_to_text",
    "java": "read_txt_to_text",
    "cpp": "read_txt_to_text",
    "c": "read_txt_to_text",
    "h": "read_txt_to_text",
    "hpp": "read_txt_to_text",
    "cs": "read_txt_to_text",
    "js": "read_txt_to_text",
    "ts": "read_txt_to_text",
    "php": "read_txt_to_text",
    "rb": "read_txt_to_text",
    "go": "read_txt_to_text",
    "rs": "read_txt_to_text",
    "swift": "read_txt_to_text",
    "kt": "read_txt_to_text",
    "scala": "read_txt_to_text",
    "sh": "read_txt_to_text",
    "bash": "read_txt_to_text",
    "ps1": "read_txt_to_text",
    "bat": "read_txt_to_text",
    "cmd": "read_txt_to_text",
    "vbs": "read_txt_to_text",

    # 标记语言
    "html": "read_txt_to_text",
    "htm": "read_txt_to_text",
    "xml": "read_txt_to_text",
    "json": "read_txt_to_text",
    "yaml": "read_txt_to_text",
    "yml": "read_txt_to_text",
    "md": "read_txt_to_text",
    "markdown": "read_txt_to_text",

    # 配置文件
    "ini": "read_txt_to_text",
    "cfg": "read_txt_to_text",
    "conf": "read_txt_to_text",
    "properties": "read_txt_to_text",
    "env": "read_txt_to_text",

    # 数据库/查询
    "sql": "read_txt_to_text",

    # 其他文本格式
    "txt": "read_txt_to_text",
    "log": "read_txt_to_text",
    "": "read_txt_to_text",  # 无扩展名文件

    # 构建/项目文件
    "toml": "read_txt_to_text",
    "lock": "read_txt_to_text",
    "gitignore": "read_txt_to_text",

    # 网络相关
    "url": "read_txt_to_text",
    "webloc": "read_txt_to_text",
}

def get_file_type(file_path: str) -> Optional[str]:
    """安全获取文件扩展名（优先MIME检测，后备扩展名）"""
    
    # 首先检查文件是否存在
    if not os.path.isfile(file_path):
        raise FileNotFoundError
    
    try:
        # 方案1：使用python-magic（推荐）
        import magic
        mime = magic.from_file(file_path, mime=True)
        
        # 处理常见的MIME类型映射
        mime_to_ext = {
            'application/pdf': 'pdf',
            'image/jpeg': 'jpg',
            'image/png': 'png',
            'image/gif': 'gif',
            'text/plain': 'txt',
            'application/zip': 'zip',
            'application/x-rar-compressed': 'rar',
            'application/x-tar': 'tar',
            'application/gzip': 'gz'
        }
        
        # 特殊处理Office文档
        if "vnd.openxmlformats-officedocument" in mime:
            # 提取具体的Office文档类型
            if 'wordprocessingml' in mime:
                return 'docx'
            elif 'spreadsheetml' in mime:
                return 'xlsx'
            elif 'presentationml' in mime:
                return 'pptx'
        
        # 检查映射表
        if mime in mime_to_ext:
            return mime_to_ext[mime]
        
        # 通用处理：从MIME类型中提取扩展名
        if '/' in mime:
            mime_type = mime.split("/")[-1]
            # 处理复合类型如vnd.ms-excel
            if mime_type.startswith('vnd.'):
                mime_type = mime_type[4:]
            if mime_type.startswith('x-'):
                mime_type = mime_type[2:]
            return mime_type
        
        return mime
    
    except ImportError:
        # 方案2：后备使用扩展名
        ext = os.path.splitext(file_path)[1]
        if ext:
            return ext[1:].lower()  # 去掉点号并转为小写
        else:
            raise ImportError
    
def complete_filename(file_path: str) -> str:
    """补全文件名（如果缺少扩展名则自动添加）"""
    if not os.path.isfile(file_path):
        return file_path
    
    # 如果已经有扩展名，直接返回
    if os.path.splitext(file_path)[1]:
        return file_path
    
    # 获取文件类型并补全扩展名
    file_type = get_file_type(file_path)
    if file_type:
        return f"{file_path}.{file_type}"
    
    return file_path  # 无法确定类型，返回原文件名

def read_csv_to_text(file_path: str) -> str:
    """读取CSV文件并返回格式化的文本"""
    try:
        df = pd.read_csv(file_path)
        return df.to_string(index=False)
    except Exception as e:
        raise RuntimeError(f"读取CSV文件失败: {str(e)}")

def read_pdf_to_text(file_path: str) -> str:
    """使用pdfminer.six提取PDF文本（效果更好）"""
    try:
        return extract_text(file_path)
    except Exception as e:
        raise RuntimeError(f"读取PDF文件失败: {str(e)}")

def convert_doc_to_docx(doc_file: str, docx_file: str) -> None:
    """将doc文档转为docx文档"""
    try:
        doc = Document(doc_file)
        doc.save(docx_file)
    except Exception as e:
        raise RuntimeError(f"转换DOC到DOCX失败: {str(e)}")

def read_docx_to_text(file_path: str) -> str:
    """读取DOCX或DOC文件内容并返回文本"""
    try:
        # 统一处理路径
        file_path = os.path.abspath(file_path)

        if file_path.lower().endswith(".doc"):
            # 处理DOC文件
            file_dir, file_name = os.path.split(file_path)
            file_base = os.path.splitext(file_name)[0]
            docx_file = os.path.join(file_dir, f"{file_base}.docx")

            # 转换DOC到DOCX
            convert_doc_to_docx(file_path, docx_file)

            # 处理转换后的文件
            text = docx2txt.process(docx_file)

            # 删除临时转换的文件
            try:
                os.remove(docx_file)
            except:
                pass
        else:
            # 直接处理DOCX文件
            text = docx2txt.process(file_path)

        return text
    except Exception as e:
        raise RuntimeError(f"读取Word文件失败: {str(e)}")

def read_excel_to_text(file_path: str) -> str:
    """读取Excel文件内容并返回文本"""
    try:
        excel_file = pd.ExcelFile(file_path)
        text_list = []

        for sheet_name in excel_file.sheet_names:
            df = excel_file.parse(sheet_name)
            text = df.to_string(index=False)
            text_list.append(f"=== {sheet_name} ===\n{text}")

        return "\n\n".join(text_list)
    except Exception as e:
        raise RuntimeError(f"读取Excel文件失败: {str(e)}")

def read_pptx_to_text(file_path: str) -> str:
    """读取PPTX文件内容并返回文本"""
    try:
        prs = Presentation(file_path)
        text_list = []

        for slide in prs.slides:
            slide_text = []
            for shape in slide.shapes:
                if hasattr(shape, "text_frame") and shape.has_text_frame:
                    text_frame = shape.text_frame
                    if text_frame.text.strip():
                        slide_text.append(text_frame.text.strip())

            if slide_text:  # 只添加有内容的幻灯片
                text_list.append("\n".join(slide_text))

        return "\n\n".join(text_list)
    except Exception as e:
        raise RuntimeError(f"读取PPTX文件失败: {str(e)}")

def read_txt_to_text(file_path: str) -> str:
    """读取文本文件，自动检测编码"""
    try:
        with open(file_path, "rb") as f:
            raw_data = f.read()
            encoding = chardet.detect(raw_data)["encoding"] or "utf-8"
        return raw_data.decode(encoding)
    except Exception as e:
        raise RuntimeError(f"读取文本文件失败: {str(e)}")

def read_any_file_to_text(file_path: str) -> str:
    """
    根据文件扩展名自动选择适当的读取函数
    返回文件内容文本或错误信息
    """
    try:
        # 修复路径编码问题
        if isinstance(file_path, bytes):
            try:
                file_path = file_path.decode('utf-8')
            except UnicodeDecodeError:
                file_path = file_path.decode('latin1')
        
        # 标准化路径（处理反斜杠和特殊字符）
        file_path = os.path.abspath(os.path.normpath(file_path))
        
        # 检查文件是否存在
        if not os.path.exists(file_path):
            return f"文件不存在: {file_path}"
            
        # 获取文件扩展名（小写，不带点）
        file_ext = get_file_type(file_path)
        if file_ext is None or file_ext == "":
            file_ext = os.path.splitext(file_path)[1][1:].lower()
        if not file_ext:
            file_ext = "txt"  # 默认文本类型

        # 后续处理逻辑保持不变...
        func_name = SUPPORTED_EXTENSIONS.get(file_ext)
        if not func_name:
            return f"不支持 {file_ext} 格式"
            
        # 使用函数映射
        func_map = {
            "read_pdf_to_text": read_pdf_to_text,
            "read_docx_to_text": read_docx_to_text,
            "read_excel_to_text": read_excel_to_text,
            "read_pptx_to_text": read_pptx_to_text,
            "read_txt_to_text": read_txt_to_text,
            "read_csv_to_text": read_csv_to_text,
        }
        
        func = func_map.get(func_name)
        if func is None:
            return f"找不到处理 {file_ext} 文件的函数"
            
        return func(file_path)
        
    except Exception as e:
        return f"读取文件时出错: {str(e)}"

@register("astrbot_plugin_file_reader", "xiewoc", "一个将文件内容传给llm的插件", "1.0.2", "https://github.com/xiewoc/astrbot_plugin_file_reader")
class astrbot_plugin_file_reader(Star):
    def __init__(self, context: Context):
        self.file_name = ""
        self.file_dir = ""
        self.content = ""
        super().__init__(context)

    @event_message_type(EventMessageType.ALL)               # type: ignore
    async def on_receive_msg(self, event: AstrMessageEvent):
        """当获取到有文件时"""
        if event.is_at_or_wake_command:# 如果是被唤醒的状态，即：先被at一下后发送
            for item in event.message_obj.message:
                if isinstance(item, Comp.File):# 判断有无File组件
                    try:
                        file_path = await item.get_file() # 获取文件
                        completed_name = complete_filename(file_path)
                        file_dir, self.file_name = os.path.split(file_path)
                        logger.info(f"接收到文件: {completed_name}, 文件路径：{file_path}")
                        # 读取文件内容
                        self.content = read_any_file_to_text(file_path)
                        if self.content:
                            logger.info(f"读取文件{completed_name}内容成功")
                        else:
                            logger.warning(f"读取文件{completed_name}内容为空")
                    except Exception as e:
                        logger.error(f"读取文件失败: {str(e)}")
                        self.content = ""
                        self.file_name = ""

    @filter.on_llm_request()
    async def on_request(self, event: AstrMessageEvent, req: ProviderRequest):
        if self.content != "" and self.file_name != "":
            logger.info(f"将文件{self.file_name}内容添加到请求中")
            req.prompt += "文件名:" + self.file_name + "文件内容:" + self.content
            self.content = ""
            self.file_name = ""