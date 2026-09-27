import streamlit as st
import requests
import re
import os
import base64
import json
import time
import io
import genanki
from datetime import datetime
from pptx import Presentation
import pdfplumber
from docx import Document

# ==================== 1. 配置 API（云端部署，从 Secrets 读取） ====================
DEEPSEEK_API_KEY = st.secrets.get("DEEPSEEK_API_KEY", "")
DEEPSEEK_URL = "https://api.deepseek.com/chat/completions"
ZHIPU_API_KEY = st.secrets.get("ZHIPU_API_KEY", "")
ZHIPU_URL = "https://open.bigmodel.cn/api/paas/v4/chat/completions"

# ==================== 2. 状态初始化 ====================
if 'last_analysis' not in st.session_state:
    st.session_state['last_analysis'] = ""
if 'last_mistake_diagnosis' not in st.session_state:
    st.session_state['last_mistake_diagnosis'] = ""
if 'parsed_ppt_text' not in st.session_state:
    st.session_state['parsed_ppt_text'] = ""
if 'parsed_note_text' not in st.session_state:
    st.session_state['parsed_note_text'] = ""
if 'thinking_stage' not in st.session_state:
    st.session_state['thinking_stage'] = "idle"
if 'user_feedback' not in st.session_state:
    st.session_state['user_feedback'] = ""
if 'last_anki' not in st.session_state:
    st.session_state['last_anki'] = ""


# ==================== 3. 数据存储机制（云端版已关闭本地写入） ====================
def load_mistakes():
    return []  # 云端内测期间，不加载本地错题本


def save_mistake(question, diagnosis, source="文字"):
    return  # 云端内测期间，禁止写入本地文件，防止数据串台


def save_analysis_to_history(content, analysis_type):
    return  # 云端内测期间，禁止写入本地文件，防止数据串台


# ==================== 4. Anki 自动打包 ====================
def generate_apkg(csv_text):
    try:
        my_model = genanki.Model(
            1607392319, 'SMU_AI_Med_Model',
            fields=[{'name': 'Front'}, {'name': 'Back'}],
            templates=[{'name': 'Medical_Card', 'qfmt': '{{Front}}', 'afmt': '{{FrontSide}}<hr id="answer">{{Back}}'}]
        )
        my_deck = genanki.Deck(2059400110, 'SMU_AI_Medical_Deck')
        lines = csv_text.strip().split('\n')
        for line in lines:
            parts = line.split('|')
            if len(parts) >= 2:
                note = genanki.Note(model=my_model, fields=[parts[0].strip(), parts[1].strip()])
                my_deck.add_note(note)
        apkg_buffer = io.BytesIO()
        genanki.Package(my_deck).write_to_file(apkg_buffer)
        apkg_buffer.seek(0)
        return apkg_buffer
    except Exception as e:
        return None


# ==================== 5. 本地知识库与 PubMed ====================
def get_local_knowledge():
    knowledge = ""
    files = {"名师思维链库": "expert_cot.txt", "医学类比库": "analogy_lib.txt", "历年真题考频": "exam_freq.txt"}
    for name, filename in files.items():
        if os.path.exists(filename):
            with open(filename, 'r', encoding='utf-8') as f:
                content = f.read().strip()
                if content: knowledge += f"\n【{name}参考】\n{content}\n"
    return knowledge


def search_pubmed(query, max_results=3):
    try:
        search_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?db=pubmed&term={query}&retmode=json&retmax={max_results}"
        search_res = requests.get(search_url, timeout=10).json()
        id_list = search_res['esearchresult']['idlist']
        if not id_list:
            return "未找到相关文献。"
        ids = ",".join(id_list)
        fetch_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&id={ids}&retmode=text&rettype=abstract"
        fetch_res = requests.get(fetch_url, timeout=15)
        time.sleep(1)
        return fetch_res.text
    except Exception as e:
        return f"PubMed检索失败：{str(e)}"


# ==================== 6. 核心 AI 处理 ====================
def get_ai_response(ppt_text, audio_text, mode, work_mode="微观伴读模式", thinking_mode=True, user_feedback="",
                    enable_pubmed=False):
    combined_input = f"【PPT文本】\n{ppt_text}\n\n【课堂录音/笔记】\n{audio_text}"
    local_knowledge = get_local_knowledge()

    pubmed_context = ""
    if enable_pubmed and ppt_text.strip():
        with st.spinner("🌐 正在检索 PubMed 最新文献..."):
            pubmed_context = f"\n【PubMed最新文献参考】\n{search_pubmed(ppt_text[:100])}\n"

    # 老师原话精炼化（保留原意，去口语化，转书面语）
    teacher_rule = "【最高优先级】如果老师在录音里强调了'必考'、'重点'、'注意'、'记住'，请提取老师强调时的核心内容。要求：1. 必须保留老师的原意和强调力度（如'必须掌握'、'期末必考'）；2. 必须去除口语化废话（如'啊、吧、呢、同学们'等语气词）；3. 转化为严谨、精炼的医学书面语，控制在1-3句话以内；4. 用 <teacher_important> 和 </teacher_important> 包裹起来，绝对不能遗漏老师强调的话！"

    # 重点高亮（从单纯的名词升级为高频考点和结论）
    highlight_rule = "请用 [[ ]] 把最核心的【高频考点】、【必须掌握的结论】或【关键医学名词】包裹起来。不要只标名词，要把整句的重点结论或机制也用 [[ ]] 标出。"

    if mode == "解构":
        if "临床趣味脑洞模式" in work_mode:
            system_prompt = f"""你是三甲医院的一名资深主治医师，带教经验丰富，说话幽默风趣。
请根据我提供的【PPT文本】和【课堂笔记】，找到最核心的 1-2 个医学机制，把它转化为一段风趣、生动的“查房情景喜剧”。
【创作要求】：1. 包含医学准确性；2. 结构化输出：🩺【临床查房情景剧】、💊【看病如破案】、🤣【记忆梗】；3. {highlight_rule}
{local_knowledge}{pubmed_context}
直接输出段子，把医学生逗笑你就赢了！"""
        elif "全篇融合模式" in work_mode:
            system_prompt = f"""你是一名资深的医学学霸导师。请将【PPT文本】与【课堂录音/笔记】完美融合，整理出一份详尽、连贯、有深度的整章知识点复习大纲。
【工作原则】：1. 全面覆盖，不再删减；2. 重点凸显：{teacher_rule}；3. 深度串联：穿插“微观机制 -> 病理生理联系 -> 药理干预”；4. 保留高亮：{highlight_rule}
【输出格式】：请使用清晰的多级标题和列表。若提供了最新文献，请务必在最后加上一段“前沿进展”。
{local_knowledge}{pubmed_context}"""
        else:
            if thinking_mode and not user_feedback:
                system_prompt = f"""你是一位医学导师。请根据以下文本内容，抛出一个临床情景题，引导学生思考。
【绝对铁律】：只输出题目，不要输出任何解析！直接输出题目即可！
{local_knowledge}{pubmed_context}"""
            elif thinking_mode and user_feedback == "我不会":
                system_prompt = f"""你是一位医学导师。学生表示他暂时不会。请用大白话和生活化的比喻，详细拆解该知识点的微观机制 -> 病理生理联系 -> 药理干预。
{teacher_rule}
【输出结构】：🔴 必考核心；💡 白话拆解；🧠 思维发散。{highlight_rule}
{local_knowledge}{pubmed_context}"""
            elif thinking_mode and user_feedback == "我会了":
                system_prompt = f"""你是一位医学导师。学生表示他已经掌握了该知识点。
【死命令】：
1. 【全局大纲】第一件事：必须先输出一份本章节完整的【知识点大纲】，用于帮学生建立整体框架！
2. 【重点保留】在大纲中，必须保留老师强调的内容！用 <teacher_important> 标签包裹这些强调句！（严禁删掉它们）
3. 【题目分界线】大纲输出完毕后，必须换行输出 === 变式题 === 作为分隔符。
4. 【题目生成】分隔符下方，直接生成 3 道变式题。第一句话必须是“第一题：...”
5. 【难度限制】本系统主要面向大一医学生，题目难度必须贴合大一《细胞生物学》水平！重点是基础概念、分子机制，绝对不能出超纲的临床综合题！
6. 【格式要求】每道题必须是一道完整的选择题！必须包含：题干 + A、B、C、D四个选项。
7. 【切分规则】每道题的输出格式必须严格是：'(题干+A、B、C、D四个选项)|(详细答案解析)'，用英文竖线分隔。绝对不能让选项出现在竖线后面！
{local_knowledge}{pubmed_context}"""
            else:
                system_prompt = f"""你是由三位资深医学专家组成的“多智能体AI导师团队”。请从文本中提取核心考点。
{teacher_rule}
【输出结构】：🔴 必考核心（提取3-8个核心考点，详尽写出微观机制 -> 病理生理联系 -> 药理干预。）；💡 白话拆解；🧠 思维发散。
【铁律】：必须包含以上三段！{highlight_rule}
{local_knowledge}{pubmed_context}"""

    elif mode == "出题":
        system_prompt = f"你是医学出题官。请根据以下文本内容，出3道选择题。{teacher_rule}请严格使用CSV格式输出，用竖线'|'分隔。"

    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {DEEPSEEK_API_KEY}"}
    data = {
        "model": "deepseek-chat",
        "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": combined_input}],
        "temperature": 0.8 if "临床趣味" in work_mode else 0.4, "stream": False, "max_tokens": 4000
    }
    try:
        response = requests.post(DEEPSEEK_URL, headers=headers, json=data, timeout=120)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        if "Connection" in str(e) or "Timeout" in str(e):
            return "⚠️ 网络连接异常，请检查您的网络后重试！"
        return f"❌ 调用出错啦：{str(e)}"


def get_diagnose_with_context(question_text, image_base64=None):
    context_text = f"【之前的PPT解构与考点分析】\n{st.session_state['last_analysis']}\n"
    system_prompt = f"""你是南方医科大学临床医学专业的资深诊断学导师。学生输入了一道错题。
请结合上述背景知识，进行“错题会诊”，严格按照以下结构输出：
1. 📝【错题识别】：提取或总结这道题的核心题干。
2. 🎯【考查核心】：这道题实际考查的核心医学概念是什么？（结合背景知识）
3. ❌【错因剖析】：分析学生的错误属于哪一类？
4. ✅【正确推导】：给出正确答案，并写出推演过程。
5. 🩺【处方建议】：给出3条精准的复习建议。
6. 🚀【举一反三】：基于这道题的知识点，出2道类似的变式题（附带答案）。
{context_text}
直接输出诊断报告，不要客套话。"""

    if image_base64:
        headers = {"Content-Type": "application/json", "Authorization": f"Bearer {ZHIPU_API_KEY}"}
        data = {"model": "glm-5.3-flash", "messages": [{"role": "user",
                                                        "content": [{"type": "text", "text": system_prompt},
                                                                    {"type": "image_url", "image_url": {
                                                                        "url": f"data:image/jpeg;base64,{image_base64}"}}]}],
                "temperature": 0.4}
        try:
            response = requests.post(ZHIPU_URL, headers=headers, json=data, timeout=120)
            response.raise_for_status()
            return response.json()["choices"][0]["message"]["content"]
        except Exception as e:
            return f"❌ 图片诊断出错：{str(e)}"
    else:
        headers = {"Content-Type": "application/json", "Authorization": f"Bearer {DEEPSEEK_API_KEY}"}
        data = {"model": "deepseek-chat",
                "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": question_text}],
                "temperature": 0.4, "stream": False, "max_tokens": 3000}
        try:
            response = requests.post(DEEPSEEK_URL, headers=headers, json=data, timeout=120)
            response.raise_for_status()
            return response.json()["choices"][0]["message"]["content"]
        except Exception as e:
            return f"❌ 文字诊断出错：{str(e)}"


def get_knowledge_graph(topic_text):
    combined_context = f"【第一部分：课程解构】\n{st.session_state['last_analysis']}\n\n【第二部分：错题诊断】\n{st.session_state['last_mistake_diagnosis']}"
    system_prompt = f"""你是知识图谱架构师。请根据以下学生学过的内容以及做错的题目，提取核心医学概念，并生成一段用于可视化知识网络的 Mermaid 代码。
【要求】：
1. 必须使用 Markdown 的代码块包裹 Mermaid 代码，语言标签为 mermaid。
2. 节点名尽量用简短的中文，逻辑关系必须符合医学逻辑（如：病因-->病理机制-->临床表现-->治疗）。
3. 可以将错题中暴露的薄弱知识点用特殊的颜色或形状标记出来（例如: A{{薄弱点}}）。
4. 代码示例如下：
graph TD
    A[钠钾泵] --> B(维持渗透压)
    A --> C(消耗ATP)
    A --> D(强心苷靶点)
{combined_context}
用户要求聚焦的主题：{topic_text}
直接输出 Mermaid 代码块，不要解释。"""
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {DEEPSEEK_API_KEY}"}
    data = {"model": "deepseek-chat", "messages": [{"role": "system", "content": system_prompt}], "temperature": 0.3,
            "stream": False}
    try:
        response = requests.post(DEEPSEEK_URL, headers=headers, json=data, timeout=120)
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"❌ 调用出错啦：{str(e)}"


# ==================== 7. 解析与渲染 ====================
def extract_text_from_file(uploaded_file):
    text = ""
    if uploaded_file.name.endswith('.pptx'):
        prs = Presentation(uploaded_file)
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"): text += shape.text + "\n"
    elif uploaded_file.name.endswith('.docx'):
        doc = Document(uploaded_file)
        for para in doc.paragraphs: text += para.text + "\n"
    elif uploaded_file.name.endswith('.pdf'):
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages: text += page.extract_text() + "\n"
    elif uploaded_file.name.endswith('.txt'):
        text = str(uploaded_file.read(), "utf-8")
    return text.strip()


def apply_highlight(text, hl_color, teacher_color):
    text = text.replace("[[",
                        f"<span style='background-color:{hl_color}; color:#000; padding:2px 4px; border-radius:4px; font-weight:bold;'>").replace(
        "]]", "</span>")
    text = text.replace("<teacher_important>",
                        f"<span style='background-color:{teacher_color}; color:#000; padding:2px 6px; border-radius:4px; font-weight:bold;'>")
    text = text.replace("</teacher_important>", "</span>")
    return text


# ==================== 8. 网页前端 ====================
st.title("🏥 南医大医学AI学习系统（云端内测版）")

with st.sidebar:
    st.header("📥 数据导入")
    uploaded_ppt = st.file_uploader("① 上传PPT/教材", type=['pptx', 'pdf', 'txt'])
    uploaded_audio = st.file_uploader("② 上传音频（待接入Whisper）", type=['mp3', 'wav', 'm4a'], disabled=True)
    st.caption("音频转文字：请先用飞书妙记/通义听悟导出为Word/TXT，再使用下方通道③上传。")
    uploaded_doc = st.file_uploader("③ 上传课堂笔记/逐字稿（Word/TXT/PDF）", type=['docx', 'txt', 'pdf'])
    st.divider()
    st.header("⚙️ 系统设置")
    course = st.selectbox("当前课程", ["细胞生物学", "系统解剖学", "病理学"])
    work_mode = st.radio("选择智能体工作模式", ["📖 微观伴读模式", "🏛️ 全篇融合模式", "🩺 临床趣味脑洞模式"])

    thinking_mode = st.toggle("🧠 开启深度思考模式（先思考后看答案）", value=True,
                              help="开启后，AI会先抛出一个临床情景问题，等您点击'我会了'或'我不会'后，再给出对应的深度解析或变式题。")
    enable_pubmed = st.checkbox("🔍 开启 PubMed 最新文献检索", value=False)

    st.markdown("**🎨 全局高亮色**")
    hl_core = st.color_picker("正文名词高亮色", "#FFEB3B")
    teacher_color = st.color_picker("老师强调内容-背景色", "#FFCDD2")

    st.markdown("**🏷️ 微观伴读专属标签色**")
    color_core = st.color_picker("必考核心-标签色", "#FFCDD2")
    color_explain = st.color_picker("白话拆解-标签色", "#FFE0B2")
    color_think = st.color_picker("思维发散-标签色", "#BBDEFB")
    st.success("系统状态：Ready")

if uploaded_ppt is not None:
    if 'last_uploaded_ppt' not in st.session_state or st.session_state.get('last_uploaded_ppt') != uploaded_ppt.name:
        with st.spinner("正在解析PPT..."):
            st.session_state['parsed_ppt_text'] = extract_text_from_file(uploaded_ppt)
            st.session_state['last_uploaded_ppt'] = uploaded_ppt.name
            st.sidebar.success("PPT解析成功！")

if uploaded_doc is not None:
    if 'last_uploaded_doc' not in st.session_state or st.session_state.get('last_uploaded_doc') != uploaded_doc.name:
        with st.spinner("正在解析笔记..."):
            st.session_state['parsed_note_text'] = extract_text_from_file(uploaded_doc)
            st.session_state['last_uploaded_doc'] = uploaded_doc.name
            st.sidebar.success("笔记解析成功！")

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📚 课程伴读", "🧠 错题诊断", "📇 卡片与图谱", "📕 我的错题本", "📂 我的分析库"])

with tab1:
    st.subheader("📥 输入区")
    col_input1, col_input2 = st.columns(2)
    with col_input1:
        ppt_text = st.text_area("PPT原文", value=st.session_state.get('parsed_ppt_text', ''), height=250,
                                placeholder="老师今天讲了...")
    with col_input2:
        audio_text = st.text_area("课堂笔记/录音转写", value=st.session_state.get('parsed_note_text', ''), height=250,
                                  placeholder="老师强调，这个必考...")

    btn_all = st.button("🚀 一键全自动完成（融合解构 + 生成Anki卡片）", use_container_width=True)
    st.divider()

    if st.session_state['thinking_stage'] == "idle":
        if btn_all:
            if not ppt_text and not audio_text:
                st.warning("请先输入内容！")
            else:
                with st.status("AI导师团队正在紧急处理中...", expanded=True) as status:
                    if thinking_mode and "微观伴读模式" in work_mode:
                        st.write("🧠 正在准备临床思考题...")
                        result_analyze = get_ai_response(ppt_text, audio_text, "解构", work_mode, thinking_mode,
                                                         enable_pubmed=enable_pubmed)
                        st.session_state['last_analysis'] = result_analyze
                        st.session_state['thinking_stage'] = "awaiting_feedback"
                        st.rerun()
                    else:
                        st.write("🔍 步骤 1：正在融合分析...")
                        result_analyze = get_ai_response(ppt_text, audio_text, "解构", work_mode, thinking_mode,
                                                         enable_pubmed=enable_pubmed)
                        st.write("📇 步骤 2：正在转化为Anki卡片...")
                        result_anki = get_ai_response(ppt_text, audio_text, "出题", work_mode, thinking_mode,
                                                      enable_pubmed=enable_pubmed)
                        st.session_state['last_analysis'] = result_analyze
                        st.session_state['last_anki'] = result_anki
                        st.session_state['thinking_stage'] = "done"
                        st.rerun()

    elif st.session_state['thinking_stage'] == "awaiting_feedback":
        st.subheader("🧠 临床思考题")
        st.info(st.session_state['last_analysis'])
        st.write("请先自己思考，然后选择你的情况：")
        col_fb1, col_fb2 = st.columns(2)
        if col_fb1.button("✅ 我会了（挑战高阶题）", use_container_width=True):
            st.session_state['user_feedback'] = "我会了"
            with st.spinner("正在生成高阶变式题..."):
                result_analyze = get_ai_response(ppt_text, audio_text, "解构", work_mode, thinking_mode, "我会了",
                                                 enable_pubmed)
                result_anki = get_ai_response(ppt_text, audio_text, "出题", work_mode, thinking_mode,
                                              enable_pubmed=enable_pubmed)
                st.session_state['last_analysis'] = result_analyze
                st.session_state['last_anki'] = result_anki
                st.session_state['thinking_stage'] = "done"
                st.rerun()
        if col_fb2.button("❌ 我不会（降维讲解）", use_container_width=True):
            st.session_state['user_feedback'] = "我不会"
            with st.spinner("正在生成基础解析..."):
                result_analyze = get_ai_response(ppt_text, audio_text, "解构", work_mode, thinking_mode, "我不会",
                                                 enable_pubmed)
                result_anki = get_ai_response(ppt_text, audio_text, "出题", work_mode, thinking_mode,
                                              enable_pubmed=enable_pubmed)
                st.session_state['last_analysis'] = result_analyze
                st.session_state['last_anki'] = result_anki
                st.session_state['thinking_stage'] = "done"
                st.rerun()

    if st.session_state['thinking_stage'] == "done":
        st.subheader("📤 AI 解构结果")
        result = st.session_state['last_analysis']

        if "我会了" in st.session_state.get('user_feedback', ''):
            if "=== 变式题 ===" in result:
                outline_part, questions_part = result.split("=== 变式题 ===", 1)
                st.markdown("### 📖 本章节知识点大纲")
                st.markdown(apply_highlight(outline_part.strip(), hl_core, teacher_color), unsafe_allow_html=True)
                st.markdown("### 📝 变式训练题")

                blocks = re.split(r'(?=第[一二三四五六七八九十0-9]+题[:：])', questions_part)
                for block in blocks:
                    if not block.strip(): continue
                    if "|" in block:
                        q, a = block.split("|", 1)
                        st.markdown(f"**{q.strip()}**")
                        with st.expander("🎯 点击查看答案解析（请先思考后再点开！）"):
                            st.markdown(apply_highlight(a.strip(), hl_core, teacher_color), unsafe_allow_html=True)
                    else:
                        st.markdown(block.strip())
            else:
                st.markdown(apply_highlight(result, hl_core, teacher_color), unsafe_allow_html=True)
        elif "全篇融合模式" in work_mode or "临床趣味脑洞模式" in work_mode:
            st.markdown(apply_highlight(result, hl_core, teacher_color), unsafe_allow_html=True)
        else:
            result = result.replace("🔴【", "🔴 ").replace("💡【", "💡 ").replace("🧠【", "🧠 ").replace("】", "")
            result = result.replace("必考核心", "\n🔴 必考核心\n").replace("白话拆解", "\n💡 白话拆解\n").replace(
                "思维发散", "\n🧠 思维发散\n")
            result = re.sub(r'(?<!\n)(\d+\.\s)', r'\n\1', result)
            if "🔴 必考核心" in result and "💡 白话拆解" in result and "🧠 思维发散" in result:
                parts = result.split("🔴 必考核心");
                rest = parts[1];
                parts = rest.split("💡 白话拆解")
                core_text = apply_highlight(parts[0], hl_core, teacher_color);
                rest = parts[1];
                parts = rest.split("🧠 思维发散")
                explain_text = apply_highlight(parts[0], hl_core, teacher_color);
                think_text = apply_highlight(parts[1], hl_core, teacher_color)
                html_core = f"<div style='margin-top:20px; border-top:1px dashed #ccc; padding-top:10px;'><span style='background-color: {color_core}; color: #333; padding: 4px 10px; border-radius: 4px; font-weight: bold; font-size: 18px;'>🔴 必考核心</span></div>"
                html_explain = f"<div style='margin-top:20px; border-top:1px dashed #ccc; padding-top:10px;'><span style='background-color: {color_explain}; color: #333; padding: 4px 10px; border-radius: 4px; font-weight: bold; font-size: 18px;'>💡 白话拆解</span></div>"
                html_think = f"<div style='margin-top:20px; border-top:1px dashed #ccc; padding-top:10px;'><span style='background-color: {color_think}; color: #333; padding: 4px 10px; border-radius: 4px; font-weight: bold; font-size: 18px;'>🧠 思维发散</span></div>"
                result = html_core + core_text + html_explain + explain_text + html_think + think_text
            st.markdown(f"<div style='font-size:17px; line-height:1.9;'>{result}</div>", unsafe_allow_html=True)

        if st.button("🔄 结束本次学习，返回初始状态"):
            st.session_state['thinking_stage'] = "idle"
            st.session_state['user_feedback'] = ""
            st.session_state['last_analysis'] = ""
            st.rerun()

        st.divider()
        st.subheader("🎯 Anki 卡片下载")
        if "last_anki" in st.session_state and "|" in st.session_state['last_anki']:
            apkg_file = generate_apkg(st.session_state['last_anki'])
            if apkg_file:
                st.download_button(
                    label="⬇️ 下载Anki卡片 (.apkg 双击自动导入)",
                    data=apkg_file,
                    file_name=f"{course}_Anki卡片.apkg",
                    mime="application/octet-stream",
                    use_container_width=True
                )
                st.caption("💡 下载后直接双击该 .apkg 文件，Anki 就会自动导入，无需手动设置分隔符！")
            else:
                st.warning("Anki卡片生成格式异常，请重试。")
        else:
            st.warning("Anki卡片内容为空或尚未生成。")

with tab2:
    st.subheader("🧠 错题诊断室")
    st.caption("已自动关联你刚才的PPT解构内容。上传错题，AI将结合背景进行精准诊断并出变式题！")
    diag_tab1, diag_tab2 = st.tabs(["📸 拍图诊断", "⌨️ 文字输入诊断"])
    with diag_tab1:
        uploaded_img = st.file_uploader("上传错题图片", type=['jpg', 'jpeg', 'png'])
        if st.button("🩺 开始图片会诊", use_container_width=True) and uploaded_img is not None:
            with st.spinner("导师正在看图诊断..."):
                img_base64 = base64.b64encode(uploaded_img.read()).decode('utf-8')
                res = get_diagnose_with_context("", img_base64)
                st.session_state['last_mistake_diagnosis'] = res
                with st.expander("🎯 点击查看AI诊断与正确答案（请先自己思考3秒再点开！）"):
                    st.markdown(apply_highlight(res, hl_core, teacher_color), unsafe_allow_html=True)
                st.success("✅ 云端内测期间，诊断结果暂不存档，刷新即清空。")
    with diag_tab2:
        wrong_input = st.text_area("粘贴错题文字", height=150)
        if st.button("🩺 开始文字会诊", use_container_width=True) and wrong_input:
            with st.spinner("导师正在分析..."):
                res = get_diagnose_with_context(wrong_input)
                st.session_state['last_mistake_diagnosis'] = res
                with st.expander("🎯 点击查看AI诊断与正确答案（请先自己思考3秒再点开！）"):
                    st.markdown(apply_highlight(res, hl_core, teacher_color), unsafe_allow_html=True)
                st.success("✅ 云端内测期间，诊断结果暂不存档，刷新即清空。")

with tab3:
    st.subheader("📇 知识网络与卡片图鉴")
    st.caption("系统已自动整合【课程解构】与【错题诊断】的内容，生成综合知识图谱。")
    graph_input = st.text_area("输入你想聚焦的知识点主题（例如：细胞膜）", height=100,
                               placeholder="例如：细胞膜的结构与功能")
    if st.button("🌐 生成整合知识网络图", use_container_width=True):
        with st.spinner("正在整合分析，绘制知识网络..."):
            graph_result = get_knowledge_graph(graph_input)
            st.subheader("🌐 综合知识网络图代码")
            st.info("💡 复制下方代码，粘贴到 mermaid.live 中查看高清图表！薄弱知识点已用不同形状标出。")
            st.code(graph_result, language="markdown")

with tab4:
    st.subheader("📕 我的错题本")
    st.info("📢 云端内测期间，错题本功能暂停开放，防止多人数据冲突。")
    st.caption("如需使用完整错题本功能，请等待后续版本更新。")

with tab5:
    st.subheader("📂 我的分析历史归档")
    st.info("📢 云端内测期间，分析历史功能暂停开放，防止多人数据冲突。")
    st.caption("如需使用完整分析历史功能，请等待后续版本更新。")