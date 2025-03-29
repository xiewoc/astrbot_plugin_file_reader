from astrbot.api.event import filter, AstrMessageEvent, MessageEventResult
from astrbot.api.star import Context, Star, register
from astrbot.api.provider import ProviderRequest
import astrbot.api.message_components as Comp
from astrbot.api.all import *

import os
from pdfminer.high_level import extract_text
import docx2txt
import pandas as pd
from docx import Document
from pptx import Presentation
from typing import Dict, Callable, Optional

# 使用字典存储支持的文件类型和对应的处理函数
SUPPORTED_EXTENSIONS: Dict[str, Callable] = {
    'pdf': 'read_pdf_to_text',
    'docx': 'read_docx_to_text',
    'doc': 'read_docx_to_text',  # 添加对doc的支持
    'xlsx': 'read_excel_to_text',
    'pptx': 'read_pptx_to_text',
    'csv': 'read_txt_to_text',
    'txt': 'read_txt_to_text',
}

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
        
        if file_path.lower().endswith('.doc'):
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
            
        return '\n\n'.join(text_list)
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
                if shape.has_text_frame:
                    text_frame = shape.text_frame
                    if text_frame.text.strip():
                        slide_text.append(text_frame.text.strip())
            
            if slide_text:  # 只添加有内容的幻灯片
                text_list.append('\n'.join(slide_text))
                
        return '\n\n'.join(text_list)
    except Exception as e:
        raise RuntimeError(f"读取PPTX文件失败: {str(e)}")

def read_txt_to_text(file_path: str) -> str:
    """读取文本文件内容"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        raise RuntimeError(f"读取文本文件失败: {str(e)}")

def read_any_file_to_text(file_path: str) -> str:
    """
    根据文件扩展名自动选择适当的读取函数
    返回文件内容文本或错误信息
    """
    try:
        # 获取文件扩展名（小写，不带点）
        file_ext = os.path.splitext(file_path)[1][1:].lower()
        
        # 获取对应的处理函数名
        func_name = SUPPORTED_EXTENSIONS.get(file_ext)
        if func_name is None:
            return f"暂不支持 {file_ext} 文件格式"
        
        # 使用字典映射替代eval()更安全
        func_map = {
            'read_pdf_to_text': read_pdf_to_text,
            'read_docx_to_text': read_docx_to_text,
            'read_excel_to_text': read_excel_to_text,
            'read_pptx_to_text': read_pptx_to_text,
            'read_txt_to_text': read_txt_to_text,
        }
        
        func = func_map.get(func_name)
        if func is None:
            return f"找不到处理 {file_ext} 文件的函数"
            
        return func(file_path)
        
    except Exception as e:
        return f"读取文件时出错: {str(e)}"

@register("astrbot_plugin_file_reader", "xiewoc", "一个将文件内容传给llm的插件", "1.0.0", "https://github.com/xiewoc/astrbot_plugin_file_reader")
class astrbot_plugin_file_reader(Star):
    def __init__(self, context: Context):
        super().__init__(context)
    
    
    @event_message_type(EventMessageType.ALL) 
    async def on_receive_msg(self, event: AstrMessageEvent):
        '''当获取到有文件时'''
        for item in event.message_obj.message:
            if isinstance(item, File):
                print("get file:",item.file)
                print("content:",read_any_file_to_text(item.file))

                global file_name
                file_dir, file_name = os.path.split(item.file)

                global content
                content = read_any_file_to_text(item.file)


    @filter.on_llm_request()
    async def my_custom_hook_1(self, event: AstrMessageEvent, req: ProviderRequest): 
        global content, file_name
        print(req)
        if content != '' and file_name != '':
            req.prompt += '文件名：' + file_name + '文件内容:' + content
            content = ''
            file_name = ''
