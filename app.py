import streamlit as st
import docx
from docx.shared import RGBColor
import language_tool_python
import re
from io import BytesIO

# 初始化语法检查引擎
@st.cache_resource
def load_tool():
    return language_tool_python.LanguageTool('en-US')

tool = load_tool()

# 判断是否包含中文
def has_chinese(text):
    return bool(re.search(r'[\u4e00-\u9fa5]', text))

# --- 网页界面设计 ---
st.set_page_config(page_title="双语剧本语法标红校对工具", page_icon="🎬", layout="centered")

st.title("🎬 双语剧本语法标红校对工具")
st.write("上传你的中英双语剧本。系统不会破坏你的原稿，而是将英文错误地方【标红】，并附带修改建议。")

st.markdown("---")

uploaded_file = st.file_uploader("请上传你的双语剧本 (.docx)", type=["docx"])

if uploaded_file is not None:
    st.info("正在扫描全剧本语法，请稍候...")
    
    try:
        doc = docx.Document(uploaded_file)
        new_doc = docx.Document()
        
        total_errors = 0
        
        for paragraph in doc.paragraphs:
            original_text = paragraph.text
            
            # 如果是空行，直接保留
            if not original_text.strip():
                new_doc.add_paragraph("")
                continue
                
            # 建立一个新段落，并保持原有的对齐格式
            new_para = new_doc.add_paragraph()
            new_para.alignment = paragraph.alignment
            
            # 如果这一行包含中文（通常是中文翻译、动作描写），为了不误报，跳过语法检查直接写入原样
            if has_chinese(original_text):
                run = new_para.add_run(original_text)
                continue
                
            # 如果是纯英文段落，进行精准的错词定位标红
            matches = tool.check(original_text)
            
            if not matches:
                # 没有任何错误，保持原样写入
                new_para.add_run(original_text)
            else:
                total_errors += len(matches)
                last_idx = 0
                
                # 逐个对段落内的错误进行切片和标红处理
                for match in matches:
                    start = match.offset
                    end = match.offset + match.errorLength
                    
                    # 1. 写入错误发生前的正常文本
                    if start > last_idx:
                        new_para.add_run(original_text[last_idx:start])
                        
                    # 2. 写入错误的文本，并将其【标红】
                    error_run = new_para.add_run(original_text[start:end])
                    error_run.font.color.rgb = RGBColor(255, 0, 0)  # 红色
                    error_run.bold = True
                    
                    # 3. 紧跟其后插入【绿色的修改建议】
                    if match.replacements:
                        suggestion = match.replacements[0]
                        suggest_run = new_para.add_run(f"【👉 建议改为: {suggestion}】")
                        suggest_run.font.color.rgb = RGBColor(0, 128, 0)  # 绿色
                        suggest_run.bold = True
                        
                    last_idx = end
                    
                # 4. 写入段落尾部的剩余文本
                if last_idx < len(original_text):
                    new_para.add_run(original_text[last_idx:])
                    
        # 保存到内存提供下载
        output = BytesIO()
        new_doc.save(output)
        output.seek(0)
        
        st.balloons()
        if total_errors > 0:
            st.success(f"✨ 校对完成！全篇共发现 {total_errors} 处英文语法/拼写可疑点，已在文档中为您标红。")
        else:
            st.success("🎉 全篇非常完美！未检测到任何明显的英文语法错误。")
            
        st.download_button(
            label="📥 下载【标红批注版】剧本 (.docx)",
            data=output,
            file_name=f"Marked_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    except Exception as e:
        st.error(f"处理文档时出错: {e}")