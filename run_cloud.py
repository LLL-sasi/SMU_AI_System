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
from supabase import create_client, Client

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==================== 0. 云端环境初始化 ====================
DEEPSEEK_API_KEY = st.secrets.get("DEEPSEEK_API_KEY", "")
DEEPSEEK_URL = "https://api.deepseek.com/chat/completions"

ZHIPU_API_KEY = st.secrets.get("ZHIPU_API_KEY", "")
ZHIPU_URL = "https://open.bigmodel.cn/api/paas/v4/chat/completions"

SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "")
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

DAILY_LIMIT = int(st.secrets.get("DAILY_LIMIT", "20"))


# ==================== 1. 用户认证模块 ====================
def auth_user():
    """处理云端用户的登录与注册"""
    if 'user' not in st.session_state:
        st.session_state['user'] = None

    if st.session_state['user'] is None:
        st.title("🏥 南医大医学AI学习系统（云端版）")
        st.subheader("请登录或注册")
        email = st.text_input("邮箱")
        password = st.text_input("密码", type="password")
        col1, col2 = st.columns(2)

        if col1.button("登录", use_container_width=True):
            try:
                res = supabase.auth.sign_in_with_password({"email": email, "password": password})
                st.session_state['user'] = res.user
                st.rerun()
            except Exception as e:
                st.error(f"登录失败：{e}")

        if col2.button("注册", use_container_width=True):
            try:
                res = supabase.auth.sign_up({"email": email, "password": password})
                st.success("注册成功，请登录。")
            except Exception as e:
                st.error(f"注册失败：{e}")
        st.stop()


auth_user()
USER_ID = st.session_state['user'].id

# ==================== 2. 状态初始化 ====================
if 'last_analysis' not in st.session_state: st.session_state['last_analysis'] = ""
if 'last_mistake_diagnosis' not in st.session_state: st.session_state['last_mistake_diagnosis'] = ""
if 'parsed_ppt_text' not in st.session_state: st.session_state['parsed_ppt_text'] = ""
if 'parsed_note_text' not in st.session_state: st.session_state['parsed_note_text'] = ""
if 'thinking_stage' not in st.session_state: st.session_state['thinking_stage'] = "idle"
if 'user_feedback' not in st.session_state: st.session_state['user_feedback'] = ""
if 'last_anki' not in st.session_state: st.session_state['last_anki'] = ""
if 'tasks' not in st.session_state: st.session_state['tasks'] = []


# ==================== 3. 云端数据库操作 ====================
def check_daily_limit():
    """检查用户今日调用次数是否超限"""
    today = date.today().isoformat()
    res = supabase.table("usage_records").select("count").eq("user_id", USER_ID).eq("date", today).execute()
    if res.data:
        return res.data[0]["count"] < DAILY_LIMIT
    return True


def update_daily_limit():
    """更新用户调用次数"""
    today = date.today().isoformat()
    res = supabase.table("usage_records").select("id, count").eq("user_id", USER_ID).eq("date", today).execute()
    if res.data:
        new_count = res.data[0]["count"] + 1
        supabase.table("usage_records").update({"count": new_count}).eq("id", res.data[0]["id"]).execute()
    else:
        supabase.table("usage_records").insert({"user_id": USER_ID, "date": today, "count": 1}).execute()


def save_mistake_to_cloud(question, diagnosis, source="文字"):
    supabase.table("mistake_book").insert({
        "user_id": USER_ID, "time": datetime.now().isoformat(),
        "source": source, "question": question, "diagnosis": diagnosis
    }).execute()


def load_mistakes_from_cloud():
    res = supabase.table("mistake_book").select("*").eq("user_id", USER_ID).order("time", desc=True).execute()
    return res.data if res.data else []


def save_analysis_to_cloud(content, analysis_type):
    supabase.table("analysis_history").insert({
        "user_id": USER_ID, "time": datetime.now().isoformat(),
        "type": analysis_type, "content": content
    }).execute()


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
    prompt = """你是一名专业的医学助教。请详细描述这张医学图片的核心内容，包括：1. 图中展示的主要结构和分子机制；2. 图中的关键节点和信号流向；3. 图中标注的文字和符号。"""
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
            with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:  # 云端限制单线程
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


# ==================== 6. 核心 AI 处理（无 ChromaDB 轻量版） ====================
def call_deepseek(system_prompt, user_input, temperature=0.4, max_tokens=4000):
    if not check_daily_limit():
        return "⚠️ 今日云端使用次数已达上限（防刷机制），请明天再来。"
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {DEEPSEEK_API_KEY}"}
    data = {"model": "deepseek-chat",
            "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": user_input}],
            "temperature": temperature, "stream": False, "max_tokens": max_tokens}
    try:
        response = requests.post(DEEPSEEK_URL, headers=headers, json=data, timeout=120)
        response.raise_for_status()
        update_daily_limit()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"⚠️ 系统繁忙：{str(e)}"


def get_ai_response(ppt_text, audio_text, mode, work_mode="微观伴读模式", thinking_mode=True,
                    user_feedback="", enable_pubmed=False):
    combined_input = f"【PPT文本】\n{ppt_text}\n\n【课堂笔记】\n{audio_text}"

    teacher_rule = """【红黄绿灯分级】...（此处保留你本地版的规则）
    🔴【必考核心】 🟡【理解重点】 ⚪【了解即可】
    【死命令】必须带 <teacher_important> 标签，最多3-5个。"""

    obsidian_rule = """（此处保留你本地版的Obsidian死模板要求）"""

    if mode == "解构":
        if "全篇融合模式" in work_mode:
            system_prompt = f"你是一名医学导师。请融合PPT和笔记，整理大纲。\n{teacher_rule}\n{obsidian_rule}"
        else:
            if thinking_mode and not user_feedback:
                system_prompt = "你是一位医学导师。抛出一个临床情景题。只输出题目，不输出解析。"
            elif thinking_mode and user_feedback == "我不会":
                system_prompt = f"你是一位医学导师。用大白话拆解知识点。\n{teacher_rule}"
            elif thinking_mode and user_feedback == "我会了":
                system_prompt = f"你是一位医学导师。先输出【知识点大纲】，换行输出 === 变式题 ===。下方必须使用'(题干)|(答案解析)'，每道题独占一行！"
            else:
                system_prompt = f"你是多智能体AI导师团队。提取核心考点。\n{teacher_rule}"
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
    system_prompt = "你是一名专业的医学AI助教。基于用户的提问回答。用 [[ ]] 标出核心医学名词。"
    return call_deepseek(system_prompt, user_input, temperature=0.5)


# ==================== 7. 解析与渲染（去掉了 Obsidian 原生化，保留网页清爽感） ====================
def obsidian_to_html(text):
    if not text: return text
    text = re.sub(r'==(.*?)==',
                  r'<span style="background-color: #FFEB3B; font-weight: bold; padding: 2px 4px; border-radius: 4px;">\1</span>',
                  text)
    text = text.replace("> [!important]", "> ⭐ **核心考点**")
    text = text.replace("> [!warning]", "> ⚠️ **易错点**")
    text = text.replace("> [!danger]", "> 🚨 **考试避坑**")
    text = text.replace("> [!abstract]", "> 💡 **一句话总结**")
    text = text.replace("> [!info]", "> ℹ️ **章节导航**")
    text = text.replace("> [!question]", "> ❓ **随堂检测**")
    text = text.replace("> [!success]-", "> ✅ **答案**")
    text = text.replace("[[", "🔗 ").replace("]]", "")
    text = text.replace("```mermaid", "```text")
    return text


def apply_highlight(text, hl_color, teacher_color, is_clinical_mode=False):
    text = text.replace("<teacher_important>",
                        f"<span style='background-color:{teacher_color}; color:#000; padding:2px 6px; border-radius:4px; font-weight:bold;'>").replace(
        "</teacher_important>", "</span>")
    lines = text.split('\n')
    new_lines = []
    for line in lines:
        if '|' in line or line.strip().startswith('#') or line.strip().startswith('```'):
            new_lines.append(line)
            continue
        line = line.replace("[[",
                            f"<span style='background-color:{hl_color}; color:#000; padding:2px 4px; border-radius:4px; font-weight:bold;'>").replace(
            "]]", "</span>")
        new_lines.append(line)
    return '\n'.join(new_lines)


# ==================== 8. 网页前端 ====================
st.title("🏥 南医大医学AI学习系统（云端版）")
st.caption(f"当前用户：{st.session_state['user'].email} | 每日剩余次数防刷限制已开启")
if st.sidebar.button("退出登录"):
    supabase.auth.sign_out()
    st.session_state['user'] = None
    st.rerun()

with st.sidebar:
    st.header("📥 数据导入")
    uploaded_ppt = st.file_uploader("① 上传PPT/教材", type=['pptx', 'pdf', 'txt'])
    uploaded_doc = st.file_uploader("③ 上传课堂笔记/逐字稿", type=['docx', 'txt', 'pdf'])
    st.caption("⚠️ 云端暂不支持音频转写，请先使用通义听悟导出TXT后上传。")
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

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["📚 课程伴读", "🧠 错题诊断", "📇 卡片与图谱", "📕 我的错题本", "💬 自由问答", "📅 任务规划"])

with tab1:
    st.subheader("📥 输入区")
    col1, col2 = st.columns(2)
    with col1:
        ppt_text = st.text_area("PPT原文", value=st.session_state.get('parsed_ppt_text', ''), height=250)
    with col2:
        audio_text = st.text_area("课堂笔记", value=st.session_state.get('parsed_note_text', ''), height=250)
    btn_all = st.button("🚀 一键全自动完成", use_container_width=True)

    if st.session_state['thinking_stage'] == "idle":
        if btn_all:
            if not ppt_text and not audio_text:
                st.warning("请先输入内容！")
            else:
                with st.status("AI导师团队处理中...", expanded=True) as status:
                    result = get_ai_response(ppt_text, audio_text, "解构", "微观伴读模式", True, enable_pubmed=False)
                    st.session_state['last_analysis'] = result
                    st.session_state['thinking_stage'] = "done"
                    st.rerun()
    elif st.session_state['thinking_stage'] == "done":
        st.subheader("📤 AI 解构结果")
        res = st.session_state['last_analysis']
        st.markdown(apply_highlight(obsidian_to_html(res), "#FFEB3B", "#FFCDD2"), unsafe_allow_html=True)

        st.download_button("📄 下载 Markdown 文件", data=res,
                           file_name=f"医学知识点_{datetime.now().strftime('%Y%m%d')}.md", mime="text/markdown")
        st.info("💡 云端版已移除 Obsidian 自动推送功能，请下载 Markdown 文件后手动放入 Obsidian 文件夹。")

with tab2:
    st.subheader("🧠 错题诊断室")
    wrong_input = st.text_area("粘贴错题文字", height=150)
    if st.button("开始会诊") and wrong_input:
        with st.spinner("诊断中..."):
            res = get_diagnose_with_context(wrong_input)
            st.session_state['last_mistake_diagnosis'] = res
            st.markdown(apply_highlight(obsidian_to_html(res), "#FFEB3B", "#FFCDD2"), unsafe_allow_html=True)
            save_mistake_to_cloud(wrong_input, res)

with tab3:
    st.subheader("📇 知识网络与卡片图鉴")
    if st.button("生成知识网络图"):
        with st.spinner("绘制中..."):
            graph_result = get_knowledge_graph("输出整章框架")
            mermaid_code = graph_result.replace("```mermaid", "").replace("```", "").strip()
            st.code(mermaid_code, language="markdown")

with tab4:
    st.subheader("📕 我的错题本")
    mistakes = load_mistakes_from_cloud()
    if not mistakes:
        st.info("暂无错题记录。")
    else:
        for item in mistakes:
            with st.expander(f"{item['time'][:10]} | {item['source']}"):
                st.markdown(f"**原题：**\n{item['question']}")
                st.markdown(f"**诊断：**\n{item['diagnosis']}")

with tab5:
    st.subheader("💬 自由问答")
    if prompt := st.chat_input("输入你的医学问题..."):
        with st.chat_message("user"): st.markdown(prompt)
        with st.chat_message("assistant"):
            with st.spinner("AI正在思考中..."):
                response = chat_with_ai(prompt, [])
                st.markdown(apply_highlight(obsidian_to_html(response), "#FFEB3B", "#FFCDD2"), unsafe_allow_html=True)

with tab6:
    st.subheader("📅 本周任务规划")
    st.info("（此处省略任务规划器的移植代码，直接使用你本地版 tab7 的逻辑即可，它不依赖特殊库）")