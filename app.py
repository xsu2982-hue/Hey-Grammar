import streamlit as st
import docx
import language_tool_python
from io import BytesIO

# 初始化语法检查引擎
@st.cache_resource
def load_tool():
    return language_tool_python.LanguageTool('en-US')

tool = load_tool()

# --- 网页界面设计 ---
st.set_page_config(page_title="Word剧本语法批量检查工具", page_icon="📝", layout="centered")

st.title("📝 Word剧本语法一键检查与导出")
st.write("直接上传你的 .docx 格式英文剧本，系统将自动扫描全篇语法，并允许你下载修复后的新文档。")

st.markdown("---")

# 文件上传组件（限制只能上传 docx）
uploaded_file = st.file_uploader("请上传你的英文剧本 (.docx)", type=["docx"])

if uploaded_file is not None:
    st.info("正在读取文档并检查语法，请稍候...")
    
    try:
        # 读取上传的 Word 文档
        doc = docx.Document(uploaded_file)
        new_doc = docx.Document()
        
        error_count = 0
        
        # 逐段检查并修复
        for paragraph in doc.paragraphs:
            original_text = paragraph.text
            
            if original_text.strip():
                # 检查语法错误
                matches = tool.check(original_text)
                
                if matches:
                    error_count += len(matches)
                    # 自动修复该段落的语法
                    corrected_text = language_tool_python.utils.correct(original_text, matches)
                    new_para = new_doc.add_paragraph(corrected_text)
                else:
                    # 没有错误则保持原样
                    new_para = new_doc.add_paragraph(original_text)
                    
                # 保持原有的段落对齐格式（如角色名居中）
                new_para.alignment = paragraph.alignment
            else:
                # 保留空行
                new_doc.add_paragraph("")
        
        # 将新文档保存到内存中
        output = BytesIO()
        new_doc.save(output)
        output.seek(0)
        
        st.balloons()
        if error_count > 0:
            st.success(f"✨ 检查完成！共自动修复了 {error_count} 处拼写或语法错误。")
        else:
            st.success("🎉 检查完成！全篇文档非常完美，未检测到明显错误。")
            
        # 下载按钮
        st.download_button(
            label="📥 下载语法修复版剧本 (.docx)",
            data=output,
            file_name=f"Corrected_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    except Exception as e:
        st.error(f"解析文档时出错，请确保文件未损坏。错误信息: {e}")