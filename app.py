import pandas as pd
import streamlit as st
from docx import Document
import os
import re
import sys
import zipfile
import io
import tempfile
from typing import Dict, List
from pathlib import Path

# 设置页面标题
st.set_page_config(page_title="文档批量生成工具", layout="wide")
st.title("商标文档批量生成工具")

# 初始化session状态
if 'processing_stage' not in st.session_state:
    st.session_state.processing_stage = 0  # 0: 未开始, 1: 处理完成
if 'output_dir' not in st.session_state:
    st.session_state.output_dir = ""
if 'generated_files' not in st.session_state:
    st.session_state.generated_files = []

# 占位符处理器
class PlaceholderHandler:
    """占位符处理器，支持多种占位符格式"""
    PLACEHOLDER_PATTERNS = [
        r'\{\{\s*(.*?)\s*\}\}',  # {{key}}
        r'\$\{\s*(.*?)\s*\}',    # ${key}
        r'\{\s*(.*?)\s*\}',      # {key}
        r'\[\[\s*(.*?)\s*\]\]'   # [[key]]
    ]

    @classmethod
    def find_placeholders(cls, text: str) -> List[str]:
        """查找文本中的所有占位符"""
        placeholders = []
        for pattern in cls.PLACEHOLDER_PATTERNS:
            matches = re.findall(pattern, text)
            placeholders.extend(matches)
        return list(set(placeholders))

    @classmethod
    def replace_placeholder(cls, text: str, placeholder: str, value: str) -> str:
        """替换特定格式的占位符"""
        for pattern in cls.PLACEHOLDER_PATTERNS:
            wrapped_placeholder = pattern.replace(r'(.*?)', re.escape(placeholder))
            if re.search(wrapped_placeholder, text):
                text = re.sub(wrapped_placeholder, str(value), text)
        return text

def replace_text_in_paragraph(paragraph, replacements: Dict[str, str]):
    """替换段落中的占位符，保留原有格式"""
    all_placeholders = []
    
    # 检查段落中是否包含任何占位符
    for key in replacements:
        for pattern in PlaceholderHandler.PLACEHOLDER_PATTERNS:
            wrapped_key = pattern.replace(r'(.*?)', re.escape(key))
            if re.search(wrapped_key, paragraph.text):
                all_placeholders.append(key)
                break
    
    if not all_placeholders:
        return
    
    # 获取段落的完整文本（合并所有run）
    full_text = ''.join([run.text for run in paragraph.runs])
    
    # 替换所有占位符
    for placeholder in all_placeholders:
        value = replacements.get(placeholder, '')
        for pattern in PlaceholderHandler.PLACEHOLDER_PATTERNS:
            wrapped_placeholder = pattern.replace(r'(.*?)', re.escape(placeholder))
            full_text = re.sub(wrapped_placeholder, str(value), full_text)
    
    # 清空原有runs并添加新文本
    for run in paragraph.runs:
        run.text = ""
    
    if paragraph.runs:
        paragraph.runs[0].text = full_text
    else:
        paragraph.add_run(full_text)

def process_document(template_path: str, output_path: str, replacements: Dict[str, str]) -> bool:
    """处理整个Word文档的替换"""
    try:
        doc = Document(template_path)
        
        # 处理正文段落
        for paragraph in doc.paragraphs:
            replace_text_in_paragraph(paragraph, replacements)
        
        # 处理表格
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        replace_text_in_paragraph(paragraph, replacements)
        
        # 处理页眉
        for section in doc.sections:
            for paragraph in section.header.paragraphs:
                replace_text_in_paragraph(paragraph, replacements)
            
            # 处理页脚
            for paragraph in section.footer.paragraphs:
                replace_text_in_paragraph(paragraph, replacements)
        
        doc.save(output_path)
        return True
    except Exception as e:
        st.error(f"处理文档时发生错误: {str(e)}")
        return False

def generate_output_filename(row: Dict[str, str], filename_columns: List[str]) -> str:
    """生成输出文件名"""
    filename_parts = []
    
    for col in filename_columns:
        if col in row and row[col]:
            filename_parts.append(str(row[col]))
    
    if not filename_parts:
        # 如果用户没有指定列，则使用前三个非空值
        for i, (key, value) in enumerate(row.items()):
            if i >= 3:
                break
            if value:
                filename_parts.append(str(value))
    
    if not filename_parts:
        filename_parts.append("generated_document")
    
    filename = "_".join(filename_parts)
    # 移除非法字符
    filename = re.sub(r'[\\/*?:"<>|]', "", filename).strip()
    return filename[:100]  # 限制文件名长度

def validate_dataframe(df: pd.DataFrame, required_columns: List[str]) -> None:
    """验证DataFrame是否包含必需的列"""
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Excel中缺少必要列: {', '.join(missing_columns)}")

# 文件上传区域
st.header("1. 上传文件")

col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader("上传Excel数据文件", type=["xlsx", "xls"])
with col2:
    template_file = st.file_uploader("上传Word模板文件", type=["docx"])

# 配置区域
st.header("2. 配置选项")

required_columns = st.text_input(
    "必填列（用逗号分隔）",
    value="客户案号",
    help="这些列必须在Excel中存在，否则会报错"
)

filename_columns = st.text_input(
    "文件名生成列（用逗号分隔）",
    value="客户案号",
    help="这些列的值将用于生成文件名"
)

# 处理按钮
if st.button("开始生成文档") and excel_file and template_file:
    # 创建临时目录
    temp_dir = tempfile.mkdtemp()
    st.session_state.output_dir = os.path.join(temp_dir, "生成文档")
    os.makedirs(st.session_state.output_dir, exist_ok=True)
    
    # 保存上传的文件
    excel_path = os.path.join(temp_dir, excel_file.name)
    with open(excel_path, "wb") as f:
        f.write(excel_file.getbuffer())
    
    template_path = os.path.join(temp_dir, template_file.name)
    with open(template_path, "wb") as f:
        f.write(template_file.getbuffer())
    
    # 处理Excel数据
    try:
        # 解析配置
        required_cols = [col.strip() for col in required_columns.split(",") if col.strip()]
        filename_cols = [col.strip() for col in filename_columns.split(",") if col.strip()]
        
        # 读取Excel
        df = pd.read_excel(excel_path).astype(str)
        validate_dataframe(df, required_cols)
        
        # 生成文档
        progress_bar = st.progress(0)
        status_text = st.empty()
        generated_files = []
        
        total_rows = len(df)
        success_count = 0
        
        for index, row in df.iterrows():
            # 更新进度
            progress = (index + 1) / total_rows
            progress_bar.progress(progress)
            status_text.text(f"正在处理 {index+1}/{total_rows}...")
            
            # 准备替换数据
            replacements = row.to_dict()
            
            # 生成文件名
            base_filename = generate_output_filename(replacements, filename_cols)
            output_path = os.path.join(st.session_state.output_dir, f"{base_filename}.docx")
            
            # 处理文档
            if process_document(template_path, output_path, replacements):
                success_count += 1
                generated_files.append({
                    "name": f"{base_filename}.docx",
                    "path": output_path
                })
        
        # 保存结果
        st.session_state.generated_files = generated_files
        st.session_state.processing_stage = 1
        
        # 显示结果
        st.success(f"文档生成完成！成功: {success_count}/{total_rows}")
        if success_count < total_rows:
            st.warning(f"有 {total_rows - success_count} 个文档生成失败")
        
    except Exception as e:
        st.error(f"处理过程中发生错误: {str(e)}")
        st.session_state.processing_stage = 0

# 下载区域
if st.session_state.processing_stage == 1 and st.session_state.generated_files:
    st.header("3. 下载生成的文件")
    
    # 创建ZIP文件
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for file_info in st.session_state.generated_files:
            if os.path.exists(file_info["path"]):
                zip_file.write(file_info["path"], file_info["name"])
    
    zip_buffer.seek(0)
    
    # 提供下载按钮
    st.download_button(
        label="下载所有文档 (ZIP)",
        data=zip_buffer,
        file_name="generated_documents.zip",
        mime="application/zip"
    )
    
    # 显示生成的文件列表
    st.subheader("生成的文档列表")
    for file_info in st.session_state.generated_files:
        if os.path.exists(file_info["path"]):
            with open(file_info["path"], "rb") as f:
                st.download_button(
                    label=f"下载 {file_info['name']}",
                    data=f.read(),
                    file_name=file_info["name"],
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

# 重置按钮
if st.button("重置系统"):
    # 清除session状态
    keys_to_clear = list(st.session_state.keys())
    for key in keys_to_clear:
        del st.session_state[key]
    
    # 重新初始化
    st.session_state.processing_stage = 0
    st.session_state.output_dir = ""
    st.session_state.generated_files = []
    
    st.success("系统已重置，可以开始新的处理流程！")
    st.experimental_rerun()

# 使用说明
st.sidebar.header("使用说明")
st.sidebar.markdown("""
1. **上传文件**:
   - Excel数据文件（包含占位符数据）
   - Word模板文件（包含占位符标记）

2. **配置选项**:
   - 必填列：Excel中必须存在的列
   - 文件名生成列：用于生成文件名的列

3. **开始生成**:
   - 点击"开始生成文档"按钮
   - 系统会处理所有数据行

4. **下载文档**:
   - 下载单个文档或所有文档(ZIP)
""")

st.sidebar.header("占位符格式")
st.sidebar.markdown("""
支持以下占位符格式:
- `{{key}}`
- `${key}`
- `{key}`
- `[[key]]`

在Word模板中使用这些格式标记需要替换的位置。
""")