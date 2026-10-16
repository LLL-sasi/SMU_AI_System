import streamlit as st
import streamlit.components.v1 as components
import requests
import regex as re
import os
import base64
import json
import time
import io
import concurrent.futures
from PIL import Image
import genanki
from datetime import datetime, date
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import pdfplumber
from docx import Document
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==================== 1. 配置 API ====================
DEEPSEEK_API_KEY = st.secrets.get("DEEPSEEK_API_KEY", "")
DEEPSEEK_URL = "https://api.deepseek.com/chat/completions"
ZHIPU_API_KEY = st.secrets.get("ZHIPU_API_KEY", "")
ZHIPU_URL = "https://open.bigmodel.cn/api/paas/v4/chat/completions"

# ==================== 2. 状态初始化 ====================
if 'last_analysis' not in st.session_state: st.session_state['last_analysis'] = ""
if 'last_mistake_diagnosis' not in st.session_state: st.session_state['last_mistake_diagnosis'] = ""
if 'parsed_ppt_text' not in st.session_state: st.session_state['parsed_ppt_text'] = ""
if 'parsed_note_text' not in st.session_state: st.session_state['parsed_note_text'] = ""
if 'thinking_stage' not in st.session_state: st.session_state['thinking_stage'] = "idle"
if 'user_feedback' not in st.session_state: st.session_state['user_feedback'] = ""
if 'last_anki' not in st.session_state: st.session_state['last_anki'] = ""
if 'mistake_book' not in st.session_state: st.session_state['mistake_book'] = []


# ==================== 3. 数据存储机制（本地会话版） ====================
def load_mistakes(): return st.session_state['mistake_book']


def save_mistake(question, diagnosis, source="文字"):
    st.session_state['mistake_book'].append({
        "time": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "source": source, "question": question, "diagnosis": diagnosis
    })


# ==================== 4. Anki 自动打包 ====================
def generate_apkg(csv_text):
    try:
        my_model = genanki.Model(1607392319, 'SMU_AI_Med_Model',
                                 fields=[{'name': 'Front'}, {'name': 'Back'}],
                                 templates=[{'name': 'Medical_Card', 'qfmt': '{{Front}}',
                                             'afmt': '{{FrontSide}}<hr id="answer">{{Back}}'}])
        my_deck = genanki.Deck(2059400110, 'SMU_AI_Medical_Deck')
        valid_cards = 0
        for line in csv_text.strip().split('\n'):
            if "|" not in line: continue
            parts = line.split('|', 1)
            if len(parts) >= 2:
                my_deck.add_note(genanki.Note(model=my_model, fields=[parts[0].strip(), parts[1].strip()]))
                valid_cards += 1
        if valid_cards == 0: return None
        apkg_buffer = io.BytesIO()
        genanki.Package(my_deck).write_to_file(apkg_buffer)
        apkg_buffer.seek(0)
        return apkg_buffer
    except Exception:
        return None


# ==================== 5. 图片识别与文本提取 ====================
def describe_image_with_glm(base64_image):
    prompt = "你是一名医学助教。请详细描述这张医学图片的核心内容，包括主要结构、机制、关键节点和标注文字。"
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {ZHIPU_API_KEY}"}
    data = {
        "model": "glm-4v-flash",
        "messages": [{"role": "user", "content": [
            {"type": "text", "text": prompt},
            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}]}],
        "temperature": 0.3, "max_tokens": 1500
    }
    try:
        response = requests.post(ZHIPU_URL, headers=headers, json=data, timeout=60)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"[图片识别失败: {str(e)}]"


def process_single_image(shape):
    try:
        image_bytes = shape.image.blob
        image = Image.open(io.BytesIO(image_bytes))
        if image.width > 1000:
            ratio = 1000 / image.width
            image = image.resize((1000, int(image.height * ratio)), Image.Resampling.LANCZOS)
        if image.mode in ("RGBA", "P"): image = image.convert("RGB")
        img_buffer = io.BytesIO()
        image.save(img_buffer, format="JPEG", quality=85)
        base64_image = base64.b64encode(img_buffer.getvalue()).decode('utf-8')
        return describe_image_with_glm(base64_image)
    except Exception:
        return None


def extract_ppt_with_images(uploaded_file):
    prs = Presentation(uploaded_file)
    full_text = ""
    total_images = 0
    for slide in prs.slides:
        slide_text = ""
        picture_shapes = []
        for shape in slide.shapes:
            try:
                if hasattr(shape, "text_frame") and shape.has_text_frame:
                    slide_text += shape.text_frame.text + "\n"
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and shape.width > 1828800:
                    picture_shapes.append(shape)
            except Exception:
                continue
        slide_images_desc = ""
        if picture_shapes:
            with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
                futures = [executor.submit(process_single_image, shape) for shape in picture_shapes]
                for i, future in enumerate(concurrent.futures.as_completed(futures)):
                    res = future.result()
                    if res: slide_images_desc += f"\n【图片{total_images + i + 1}内容描述】\n{res}\n"
            total_images += len(picture_shapes)
        full_text += slide_text + slide_images_desc + "\n---\n"
    if total_images > 0: st.sidebar.success(f"PPT解析完成！识别了{total_images}张关键图片。")
    return full_text.strip()


def extract_text_from_file(uploaded_file):
    if uploaded_file.name.endswith('.pptx'):
        return extract_ppt_with_images(uploaded_file)
    elif uploaded_file.name.endswith('.docx'):
        doc = Document(uploaded_file)
        return "\n".join([para.text for para in doc.paragraphs]).strip()
    elif uploaded_file.name.endswith('.pdf'):
        with pdfplumber.open(uploaded_file) as pdf:
            return "\n".join([page.extract_text() for page in pdf.pages]).strip()
    elif uploaded_file.name.endswith('.txt'):
        return str(uploaded_file.read(), "utf-8")
    return ""


# ==================== 6. 核心 AI 处理 ====================
def call_deepseek(system_prompt, user_input, temperature=0.4, max_tokens=4000):
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {DEEPSEEK_API_KEY}"}
    data = {"model": "deepseek-chat",
            "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": user_input}],
            "temperature": temperature, "stream": False, "max_tokens": max_tokens}
    try:
        response = requests.post(DEEPSEEK_URL, headers=headers, json=data, timeout=120)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"⚠️ 系统繁忙或网络错误：{str(e)}"


def get_ai_response(ppt_text, audio_text, mode, work_mode="微观伴读模式", thinking_mode=True,
                    user_feedback="", enable_pubmed=False):
    combined_input = f"【PPT文本】\n{ppt_text}\n\n【课堂笔记】\n{audio_text}"

    teacher_rule = """【红黄绿灯分级】
🔴【必考核心】：老师明确说"必考"的内容。
🟡【理解重点】：老师强调"注意、记住"的内容。
⚪【了解即可】：背景知识。
所有核心医学名词必须用 [[ ]] 包裹！"""

    if mode == "解构":
        if "全篇融合模式" in work_mode:
            system_prompt = f"你是一名医学导师。请融合PPT和笔记，整理复习大纲。\n{teacher_rule}"
        else:
            if thinking_mode and not user_feedback:
                system_prompt = "你是一位医学导师。请抛出一个临床情景题。只输出题目，不输出解析。"
            elif thinking_mode and user_feedback == "我不会":
                system_prompt = f"你是一位医学导师。用大白话拆解知识点。\n{teacher_rule}"
            elif thinking_mode and user_feedback == "我会了":
                system_prompt = f"你是一位医学导师。先输出【知识点大纲】，换行输出 === 变式题 ===。下方格式必须为'(题干)|(答案解析)'，每道题独占一行！\n{teacher_rule}"
            else:
                system_prompt = f"你是AI导师团队。提取核心考点。\n{teacher_rule}"
    elif mode == "出题":
        system_prompt = f"你是医学出题官。出3道选择题。{teacher_rule}必须用竖线'|'分隔。"

    return call_deepseek(system_prompt, combined_input, temperature=0.4)


def get_diagnose_with_context(question_text, image_base64=None):
    system_prompt = "你是资深诊断学导师。进行错题会诊：1.错题识别 2.考查核心 3.错因剖析 4.正确推导 5.临床推理路径 6.处方建议 7.举一反三。直接输出诊断报告。"
    if image_base64:
        headers = {"Content-Type": "application/json", "Authorization": f"Bearer {ZHIPU_API_KEY}"}
        data = {"model": "glm-4v-flash", "messages": [{"role": "user", "content": [
            {"type": "text", "text": system_prompt},
            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_base64}"}}]}],
                "temperature": 0.4}
        try:
            response = requests.post(ZHIPU_URL, headers=headers, json=data, timeout=120)
            response.raise_for_status()
            return response.json()["choices"][0]["message"]["content"]
        except Exception as e:
            return f"❌ 图片诊断出错：{str(e)}"
    else:
        return call_deepseek(system_prompt, question_text, temperature=0.4)


def get_knowledge_graph(topic_text):
    system_prompt = "你是医学知识图谱架构师。输出 Mermaid 代码块。节点ID用英文字母，分号后强制换行。不要解释。"
    return call_deepseek(system_prompt, "请生成图谱", temperature=0.3)


def chat_with_ai(user_input, chat_history):
    system_prompt = "你是一名专业的医学AI助教。基于用户提问回答。用 [[ ]] 标出核心医学名词。"
    return call_deepseek(system_prompt, user_input, temperature=0.5)


# ==================== 7. 解析与渲染 ====================
def apply_highlight(text, hl_color="#FFEB3B", teacher_color="#FFCDD2"):
    text = text.replace("<teacher_important>",
                        f"<span style='background-color:{teacher_color}; color:#000; padding:2px 6px; border-radius:4px; font-weight:bold;'>").replace(
        "</teacher_important>", "</span>")
    text = text.replace("[[",
                        f"<span style='background-color:{hl_color}; color:#000; padding:2px 4px; border-radius:4px; font-weight:bold;'>").replace(
        "]]", "</span>")
    return text


# ==================== 8. 网页前端 ====================
st.title("🏥 南医大医学AI学习系统（云端测试版）")
st.caption("基于多智能体协作 + 云端DeepSeek API | 内测版")
st.info("👋 欢迎使用！请在左侧上传PPT或笔记，然后点击下方的'一键全自动完成'按钮开始体验。")

with st.sidebar:
    st.header("📥 数据导入")
    uploaded_ppt = st.file_uploader("① 上传PPT/教材", type=['pptx', 'pdf', 'txt'])
    uploaded_doc = st.file_uploader("③ 上传课堂笔记/逐字稿", type=['docx', 'txt', 'pdf'])
    st.divider()

if uploaded_ppt is not None:
    if st.session_state.get('last_uploaded_ppt') != uploaded_ppt.name:
        with st.spinner("正在解析PPT..."):
            st.session_state['parsed_ppt_text'] = extract_text_from_file(uploaded_ppt)
            st.session_state['last_uploaded_ppt'] = uploaded_ppt.name

if uploaded_doc is not None:
    if st.session_state.get('last_uploaded_doc') != uploaded_doc.name:
        with st.spinner("正在解析笔记..."):
            st.session_state['parsed_note_text'] = extract_text_from_file(uploaded_doc)
            st.session_state['last_uploaded_doc'] = uploaded_doc.name

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📚 课程伴读", "🧠 错题诊断", "📇 知识图谱", "📕 我的错题本", "💬 自由问答"])

with tab1:
    st.subheader("📥 输入区")
    col1, col2 = st.columns(2)
    with col1:
        ppt_text = st.text_area("PPT原文", value=st.session_state.get('parsed_ppt_text', ''), height=250)
    with col2:
        audio_text = st.text_area("课堂笔记", value=st.session_state.get('parsed_note_text', ''), height=250)
    btn_all = st.button("🚀 一键全自动完成（融合解构 + 生成Anki卡片）", use_container_width=True)

    if st.session_state['thinking_stage'] == "idle":
        if btn_all:
            if not ppt_text and not audio_text:
                st.warning("请先输入内容！")
            else:
                with st.status("AI导师团队处理中...", expanded=True) as status:
                    result = get_ai_response(ppt_text, audio_text, "解构", "微观伴读模式", False, enable_pubmed=False)
                    st.session_state['last_analysis'] = result
                    st.session_state['thinking_stage'] = "done"
                    st.rerun()
    elif st.session_state['thinking_stage'] == "done":
        st.subheader("📤 AI 解构结果")
        res = st.session_state['last_analysis']
        st.markdown(apply_highlight(res), unsafe_allow_html=True)

        st.download_button("📄 下载 Markdown 文件", data=res,
                           file_name=f"医学知识点_{datetime.now().strftime('%Y%m%d')}.md", mime="text/markdown")

        if st.button("🔄 结束本次学习"):
            st.session_state['thinking_stage'] = "idle"
            st.session_state['last_analysis'] = ""
            st.rerun()

with tab2:
    st.subheader("🧠 错题诊断室")
    wrong_input = st.text_area("粘贴错题文字", height=150)
    if st.button("开始会诊") and wrong_input:
        with st.spinner("诊断中..."):
            res = get_diagnose_with_context(wrong_input)
            st.session_state['last_mistake_diagnosis'] = res
            st.markdown(apply_highlight(res), unsafe_allow_html=True)
            save_mistake(wrong_input, res)

with tab3:
    st.subheader("📇 知识网络图谱")
    if st.button("生成知识网络图"):
        with st.spinner("绘制中..."):
            graph_result = get_knowledge_graph("输出整章框架")
            mermaid_code = graph_result.replace("```mermaid", "").replace("```", "").strip()
            st.code(mermaid_code, language="markdown")

with tab4:
    st.subheader("📕 我的错题本")
    mistakes = load_mistakes()
    if not mistakes:
        st.info("暂无错题记录。")
    else:
        for item in mistakes:
            with st.expander(f"{item['time']} | {item['source']}"):
                st.markdown(f"**原题：**\n{item['question']}")
                st.markdown(f"**诊断：**\n{item['diagnosis']}")

with tab5:
    st.subheader("💬 自由问答")
    if prompt := st.chat_input("输入你的医学问题..."):
        with st.chat_message("user"): st.markdown(prompt)
        with st.chat_message("assistant"):
            with st.spinner("AI正在思考中..."):
                response = chat_with_ai(prompt, [])
                st.markdown(apply_highlight(response), unsafe_allow_html=True)