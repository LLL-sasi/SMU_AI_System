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
from pptx.enum.shapes import MSO_SHAPE_TYPE
import pdfplumber
from docx import Document

# 引入Mermaid渲染库
try:
    from streamlit_mermaid import st_mermaid

    HAS_MERMAID = True
except ImportError:
    HAS_MERMAID = False

# ==================== 1. 配置 API ====================
DEEPSEEK_API_KEY = st.secrets.get("DEEPSEEK_API_KEY", "sk-你的本地测试DeepSeek Key")
DEEPSEEK_URL = "https://api.deepseek.com/chat/completions"
ZHIPU_API_KEY = st.secrets.get("ZHIPU_API_KEY", "你的本地测试智谱Key")
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
if 'analysis_history' not in st.session_state: st.session_state['analysis_history'] = []
if 'chat_history' not in st.session_state: st.session_state['chat_history'] = []


# ==================== 3. 数据存储机制（会话隔离版） ====================
def load_mistakes(): return st.session_state['mistake_book']


def save_mistake(question, diagnosis, source="文字"):
    mistakes = load_mistakes()
    mistakes.append({"time": datetime.now().strftime("%Y-%m-%d %H:%M"), "source": source, "question": question,
                     "diagnosis": diagnosis})
    st.session_state['mistake_book'] = mistakes


def save_analysis_to_history(content, analysis_type):
    history = st.session_state['analysis_history']
    history.append({"time": datetime.now().strftime("%Y-%m-%d %H:%M"), "type": analysis_type, "content": content})
    if len(history) > 30: history = history[-30:]
    st.session_state['analysis_history'] = history


# ==================== 4. Anki 自动打包 ====================
def generate_apkg(csv_text):
    try:
        my_model = genanki.Model(1607392319, 'SMU_AI_Med_Model', fields=[{'name': 'Front'}, {'name': 'Back'}],
                                 templates=[{'name': 'Medical_Card', 'qfmt': '{{Front}}',
                                             'afmt': '{{FrontSide}}<hr id="answer">{{Back}}'}])
        my_deck = genanki.Deck(2059400110, 'SMU_AI_Medical_Deck')
        for line in csv_text.strip().split('\n'):
            parts = line.split('|')
            if len(parts) >= 2: my_deck.add_note(
                genanki.Note(model=my_model, fields=[parts[0].strip(), parts[1].strip()]))
        apkg_buffer = io.BytesIO()
        genanki.Package(my_deck).write_to_file(apkg_buffer)
        apkg_buffer.seek(0)
        return apkg_buffer
    except Exception:
        return None


# ==================== 5. 图片识别与文本提取 ====================
def describe_image_with_glm(base64_image):
    prompt = """你是一名专业的医学助教。请详细描述这张医学图片（如通路图、解剖图、机制图等）的核心内容，包括：1. 图中展示的主要结构和分子机制；2. 图中的关键节点和信号流向；3. 图中标注的文字和符号。请用专业、精炼的医学语言描述。"""
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {ZHIPU_API_KEY}"}
    data = {
        "model": "glm-4v-flash",
        "messages": [{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {
            "url": f"data:image/jpeg;base64,{base64_image}"}}]}],
        "temperature": 0.3, "max_tokens": 1500
    }
    try:
        response = requests.post(ZHIPU_URL, headers=headers, json=data, timeout=60)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"[图片识别失败: {str(e)}]"


def extract_ppt_with_images(uploaded_file):
    prs = Presentation(uploaded_file)
    full_text = ""
    image_count = 0
    for slide in prs.slides:
        slide_text = ""
        slide_images_desc = ""
        for shape in slide.shapes:
            if shape.has_text_frame: slide_text += shape.text_frame.text + "\n"
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    image_bytes = shape.image.blob
                    base64_image = base64.b64encode(image_bytes).decode('utf-8')
                    with st.spinner(f"正在识别第{image_count + 1}张图片..."):
                        img_desc = describe_image_with_glm(base64_image)
                    slide_images_desc += f"\n【图片{image_count + 1}内容描述】\n{img_desc}\n"
                    image_count += 1
                    time.sleep(0.5)
                except Exception as e:
                    slide_images_desc += f"\n[图片{image_count + 1}识别失败: {str(e)}]\n"
                    image_count += 1
        full_text += slide_text + slide_images_desc + "\n---\n"
    if image_count > 0: st.sidebar.success(f"PPT解析完成！识别了{image_count}张图片。")
    return full_text.strip()


def extract_text_from_file(uploaded_file):
    if uploaded_file.name.endswith('.pptx'):
        return extract_ppt_with_images(uploaded_file)
    elif uploaded_file.name.endswith('.docx'):
        doc = Document(uploaded_file);
        text = ""
        for para in doc.paragraphs: text += para.text + "\n"
        return text.strip()
    elif uploaded_file.name.endswith('.pdf'):
        with pdfplumber.open(uploaded_file) as pdf:
            text = ""
            for page in pdf.pages: text += page.extract_text() + "\n"
            return text.strip()
    elif uploaded_file.name.endswith('.txt'):
        return str(uploaded_file.read(), "utf-8")
    return ""


# ==================== 6. 本地知识库与 PubMed ====================
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
        if not id_list: return "未找到相关文献。"
        ids = ",".join(id_list)
        fetch_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&id={ids}&retmode=text&rettype=abstract"
        fetch_res = requests.get(fetch_url, timeout=15)
        time.sleep(1)
        return fetch_res.text
    except Exception as e:
        return f"PubMed检索失败：{str(e)}"


# ==================== 7. 核心 AI 处理 ====================
def get_ai_response(ppt_text, audio_text, mode, work_mode="微观伴读模式", thinking_mode=True, user_feedback="",
                    enable_pubmed=False):
    combined_input = f"【PPT文本】\n{ppt_text}\n\n【课堂录音/笔记】\n{audio_text}"
    local_knowledge = get_local_knowledge()
    pubmed_context = ""
    if enable_pubmed and ppt_text.strip():
        with st.spinner("🌐 正在检索 PubMed 最新文献..."):
            pubmed_context = f"\n【PubMed最新文献参考】\n{search_pubmed(ppt_text[:100])}\n"

    hallucination_rule = "\n【防幻觉与对比要求】1. 所有结论必须优先基于我提供的PPT文本和参考资料，若没有找到明确依据，必须加上“⚠️ 此内容为AI补充，请以教材为准”。2. 若遇到易混淆的医学名词（如渗出液vs漏出液、坏死vs凋亡），请主动使用Markdown表格进行对比分析。"
    teacher_rule = """【考点红绿灯】如果老师在录音里强调，请把核心内容按重要程度分为三级：
🔴【必考核心】：老师明确说“必考”、“期末考”的内容。
🟡【理解重点】：老师强调“注意”、“记住”但没有明确说必考的内容。
⚪【了解即可】：老师随口提到的背景知识。
格式要求：必须保留老师原意，去掉口语化废话，用 <teacher_important> 和 </teacher_important> 包裹起来。"""
    highlight_rule = "请用 [[ ]] 把最核心的【高频考点】、【必须掌握的结论】或【关键医学名词】包裹起来。"

    if mode == "解构":
        if "临床趣味脑洞模式" in work_mode:
            system_prompt = f"""你是三甲医院的一名资深主治医师。请根据【PPT文本】和【课堂笔记】，转化为一段风趣的“查房情景喜剧”。【创作要求】：1. 医学准确性；2. 结构化输出：🩺【临床查房情景剧】、💊【看病如破案】、🤣【记忆梗】；3. {highlight_rule}{hallucination_rule}\n{local_knowledge}{pubmed_context}\n直接输出段子！"""
        elif "全篇融合模式" in work_mode:
            system_prompt = f"""你是一名资深的医学学霸导师。请将【PPT文本】与【课堂录音/笔记】完美融合，整理出整章知识点复习大纲。【工作原则】：1. 全面覆盖；2. 重点凸显：{teacher_rule}；3. 深度串联：穿插“微观机制 -> 病理生理联系 -> 药理干预”；4. 保留高亮：{highlight_rule}{hallucination_rule}\n【输出格式】：多级标题和列表。若提供文献，务必加“前沿进展”。\n{local_knowledge}{pubmed_context}"""
        else:
            if thinking_mode and not user_feedback:
                system_prompt = f"""你是一位医学导师。请根据以下文本内容，抛出一个临床情景题。【绝对铁律】：只输出题目，不要输出解析！直接输出题目！\n{local_knowledge}{pubmed_context}"""
            elif thinking_mode and user_feedback == "我不会":
                system_prompt = f"""你是一位医学导师。学生表示暂时不会。请用大白话和生活化的比喻，详细拆解该知识点。\n{teacher_rule}{highlight_rule}{hallucination_rule}\n{local_knowledge}{pubmed_context}"""
            elif thinking_mode and user_feedback == "我会了":
                system_prompt = f"""你是一位医学导师。学生表示已经掌握该知识点。【死命令】：1. 【全局大纲】先输出【知识点大纲】！2. 【重点保留】必须保留老师强调的内容！用 <teacher_important> 标签包裹！3. 【题目分界线】大纲结束后，换行输出 === 变式题 === 。4. 【题目生成】下方直接生成 3 道变式题。第一句话必须是“第一题：...”5. 【格式要求】每道题必须是完整的选择题（含A、B、C、D四个选项）。6. 【切分规则】严格格式：'(题干+A、B、C、D)|(详细答案解析)'，用英文竖线分隔。\n{local_knowledge}{pubmed_context}"""
            else:
                system_prompt = f"""你是多智能体AI导师团队。请从文本中提取核心考点。\n{teacher_rule}\n【输出结构】：🔴 必考核心（提取3-8个核心考点，详尽写出微观机制 -> 病理生理联系 -> 药理干预。）；💡 白话拆解；🧠 思维发散。\n【铁律】：必须包含以上三段！{highlight_rule}{hallucination_rule}\n{local_knowledge}{pubmed_context}"""
    elif mode == "出题":
        system_prompt = f"你是医学出题官。请根据以下文本内容，出3道选择题。{teacher_rule}请严格使用CSV格式输出，用竖线'|'分隔。"

    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {DEEPSEEK_API_KEY}"}
    data = {"model": "deepseek-chat",
            "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": combined_input}],
            "temperature": 0.8 if "临床趣味" in work_mode else 0.4, "stream": False, "max_tokens": 4000}
    try:
        response = requests.post(DEEPSEEK_URL, headers=headers, json=data, timeout=120)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"⚠️ 系统繁忙或网络错误：{str(e)}"


def get_diagnose_with_context(question_text, image_base64=None):
    context_text = f"【之前的PPT解构与考点分析】\n{st.session_state['last_analysis']}\n"
    system_prompt = f"""你是南方医科大学临床医学专业的资深诊断学导师。学生输入了一道错题。
请结合上述背景知识，进行“错题会诊”，严格按照以下结构输出：
1. 📝【错题识别】：提取或总结这道题的核心题干。
2. 🎯【考查核心】：这道题实际考查的核心医学概念。
3. ❌【错因剖析】：分析学生的错误属于哪一类？
4. ✅【正确推导】：给出正确答案，并写出推演过程。
5. 🧠【临床推理路径（Clinical Reasoning）】：请用流程图的形式（如：患者主诉 -> 初步检查 -> 鉴别诊断 -> 确诊），展示这道题背后的临床思维推演逻辑。
6. 🩺【处方建议】：给出3条精准的复习建议。
7. 🚀【举一反三】：基于这道题的知识点，出2道类似的变式题（附带答案）。
{context_text}
直接输出诊断报告，不要客套话。"""
    if image_base64:
        headers = {"Content-Type": "application/json", "Authorization": f"Bearer {ZHIPU_API_KEY}"}
        data = {"model": "glm-4v-flash", "messages": [{"role": "user",
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
    system_prompt = f"""你是医学知识图谱架构师。请根据学生上传的课件内容，为其构建一张【全局认知地图】。
【核心构建规则】：
1. **全局总览优先**：从课件中提取出本章节的核心主干（一级节点），作为顶层骨架。
2. **分层展开**：将每一节的核心知识点（如：机制、病理、药理）作为子节点展开。
3. **痛点精准锚定**：如果“错题诊断”中有薄弱知识点，请必须使用菱形 A{{薄弱点}} 在全局图谱中将其标记出来！
4. **用户焦点过滤**：如果用户输入了主题词，请在全局骨架中高亮该分支。若无输入，则输出全局框架。
5. 代码必须严格包裹在 ```mermaid 和 ``` 之间。
【用户输入主题（可能为空）】：{topic_text}
【课件与错题背景】：\n{combined_context}
直接输出 Mermaid 代码块。"""
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {DEEPSEEK_API_KEY}"}
    data = {"model": "deepseek-chat", "messages": [{"role": "system", "content": system_prompt}], "temperature": 0.3,
            "stream": False}
    try:
        response = requests.post(DEEPSEEK_URL, headers=headers, json=data, timeout=120)
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"❌ 调用出错啦：{str(e)}"


def chat_with_ai(user_input, chat_history):
    current_context = st.session_state.get('parsed_ppt_text', '') + "\n" + st.session_state.get('parsed_note_text', '')
    system_prompt = f"""你是一名专业的医学AI助教。
【参考资料】：以下是学生本次课程上传的PPT和笔记内容：
{current_context[:4000]}
【回答规则】：
1. 如果提问与上述资料相关，优先基于资料回答，并指明“根据课件/笔记……”。
2. 如果资料中没有相关内容，再基于医学知识库回答，并标注“以下内容基于通用医学知识……”。
3. 绝对不要捏造医学事实！不知道请说明。
4. 请用 [[ ]] 把核心的医学名词或结论标出。"""
    messages = [{"role": "system", "content": system_prompt}]
    for chat in chat_history[-5:]:
        messages.append({"role": "user", "content": chat["user"]})
        messages.append({"role": "assistant", "content": chat["assistant"]})
    messages.append({"role": "user", "content": user_input})
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {DEEPSEEK_API_KEY}"}
    data = {"model": "deepseek-chat", "messages": messages, "temperature": 0.5, "stream": False, "max_tokens": 2500}
    try:
        response = requests.post(DEEPSEEK_URL, headers=headers, json=data, timeout=120)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"❌ 调用出错啦：{str(e)}"


# ==================== 8. 解析与渲染 ====================
def apply_highlight(text, hl_color, teacher_color):
    text = text.replace("[[",
                        f"<span style='background-color:{hl_color}; color:#000; padding:2px 4px; border-radius:4px; font-weight:bold;'>").replace(
        "]]", "</span>")
    text = text.replace("<teacher_important>",
                        f"<span style='background-color:{teacher_color}; color:#000; padding:2px 6px; border-radius:4px; font-weight:bold;'>")
    text = text.replace("</teacher_important>", "</span>")
    return text


# ==================== 9. 网页前端 ====================
st.title("🏥 南医大医学AI学习系统")
st.caption("基于多智能体协作与课堂语境融合的启发式学习平台 | 内测版")
st.info("👋 欢迎使用！请在左侧上传PPT或笔记，然后点击下方的‘一键全自动完成’按钮开始体验。新手请点左侧【使用指南】。")

with st.sidebar:
    st.header("📥 数据导入")
    uploaded_ppt = st.file_uploader("① 上传PPT/教材", type=['pptx', 'pdf', 'txt'])
    uploaded_audio = st.file_uploader("② 上传音频（待接入Whisper）", type=['mp3', 'wav', 'm4a'], disabled=True)
    st.caption("音频转文字：请先用飞书妙记/通义听悟导出为Word/TXT，再使用下方通道③上传。")
    uploaded_doc = st.file_uploader("③ 上传课堂笔记/逐字稿（Word/TXT/PDF）", type=['docx', 'txt', 'pdf'])
    st.divider()

    st.header("📊 学习仪表盘")
    col_stat1, col_stat2 = st.columns(2)
    col_stat1.metric("今日生成卡片", len(st.session_state['mistake_book']))
    col_stat2.metric("错题本累计", len(st.session_state['analysis_history']))

    st.header("⚙️ 系统设置")
    course = st.selectbox("当前课程", ["细胞生物学", "系统解剖学", "病理学"])
    work_mode = st.radio("选择智能体工作模式", ["📖 微观伴读模式", "🏛️ 全篇融合模式", "🩺 临床趣味脑洞模式"])
    thinking_mode = st.toggle("🧠 开启深度思考模式（先思考后看答案）", value=True)
    enable_pubmed = st.checkbox("🔍 开启 PubMed 最新文献检索", value=False)

    st.markdown("**🎨 全局高亮色**")
    hl_core = st.color_picker("正文名词高亮色", "#FFEB3B")
    teacher_color = st.color_picker("老师强调内容-背景色", "#FFCDD2")

    st.markdown("**🏷️ 微观伴读专属标签色**")
    color_core = st.color_picker("必考核心-标签色", "#FFCDD2")
    color_explain = st.color_picker("白话拆解-标签色", "#FFE0B2")
    color_think = st.color_picker("思维发散-标签色", "#BBDEFB")
    st.success("系统状态：Ready")

    with st.sidebar.expander("📖 点击查看使用指南"):
        st.markdown("""
        1. **课程伴读**：上传PPT和课堂笔记，点击一键生成，AI会自动提取重点并高亮老师强调的内容。
        2. **自适应学习**：开启“深度思考”后，AI会先提问，根据你的掌握情况出不同难度的题。
        3. **错题诊断**：直接拍图或粘贴错题，AI会给出错因剖析、临床推理路径，并支持一键生成Anki卡片。
        4. **自由问答**：基于上传的PPT和笔记，向AI提问任何医学问题。
        """)
    st.caption("⚠️ 声明：本系统生成内容仅供医学学习参考，不构成临床诊疗建议。")

if uploaded_ppt is not None:
    if 'last_uploaded_ppt' not in st.session_state or st.session_state.get('last_uploaded_ppt') != uploaded_ppt.name:
        with st.spinner("正在解析PPT（含图片识别，可能需要一点时间）..."):
            st.session_state['parsed_ppt_text'] = extract_text_from_file(uploaded_ppt)
            st.session_state['last_uploaded_ppt'] = uploaded_ppt.name

if uploaded_doc is not None:
    if 'last_uploaded_doc' not in st.session_state or st.session_state.get('last_uploaded_doc') != uploaded_doc.name:
        with st.spinner("正在解析笔记..."):
            st.session_state['parsed_note_text'] = extract_text_from_file(uploaded_doc)
            st.session_state['last_uploaded_doc'] = uploaded_doc.name
            st.sidebar.success("笔记解析成功！")

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["📚 课程伴读", "🧠 错题诊断", "📇 卡片与图谱", "📕 我的错题本", "📂 我的分析库", "💬 自由问答"])

with tab1:
    st.subheader("📥 输入区")
    col_input1, col_input2 = st.columns(2)
    with col_input1:
        ppt_text = st.text_area("PPT原文（含图片识别结果）", value=st.session_state.get('parsed_ppt_text', ''),
                                height=250, placeholder="老师今天讲了...")
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
                        save_analysis_to_history(result_analyze, "课程解构")
                        st.rerun()
    elif st.session_state['thinking_stage'] == "awaiting_feedback":
        st.subheader("🧠 临床思考题")
        st.info(st.session_state['last_analysis'])
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
                save_analysis_to_history(result_analyze, "课程解构-高阶")
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
                save_analysis_to_history(result_analyze, "课程解构-基础")
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

        st.download_button("📄 导出本次知识点大纲 (Markdown)", data=st.session_state['last_analysis'],
                           file_name=f"{course}_知识点大纲.md", mime="text/markdown", use_container_width=True)

        if st.button("🔄 结束本次学习，返回初始状态"):
            st.session_state['thinking_stage'] = "idle";
            st.session_state['user_feedback'] = "";
            st.session_state['last_analysis'] = "";
            st.rerun()
        st.divider();
        st.subheader("🎯 Anki 卡片下载")
        if "last_anki" in st.session_state and "|" in st.session_state['last_anki']:
            apkg_file = generate_apkg(st.session_state['last_anki'])
            if apkg_file:
                st.download_button("⬇️ 下载Anki卡片 (.apkg 双击自动导入)", data=apkg_file,
                                   file_name=f"{course}_Anki卡片.apkg", mime="application/octet-stream",
                                   use_container_width=True)
                st.caption("💡 下载后直接双击该 .apkg 文件，Anki 就会自动导入，无需手动设置分隔符！")
            else:
                st.warning("Anki卡片生成格式异常，请重试。")
        else:
            st.warning("Anki卡片内容为空或尚未生成。")

with tab2:
    st.subheader("🧠 错题诊断室")
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
                anki_csv = f"图片错题|{res}"
                apkg_data = generate_apkg(anki_csv)
                if apkg_data:
                    st.download_button("⬇️ 一键将本题生成Anki卡片 (.apkg)", data=apkg_data, file_name="错题卡片.apkg",
                                       mime="application/octet-stream", use_container_width=True)
                save_mistake("图片错题", res, "图片")
                st.success("✅ 已存入本次会话的错题本！")
    with diag_tab2:
        wrong_input = st.text_area("粘贴错题文字", height=150)
        if st.button("🩺 开始文字会诊", use_container_width=True) and wrong_input:
            with st.spinner("导师正在分析..."):
                res = get_diagnose_with_context(wrong_input)
                st.session_state['last_mistake_diagnosis'] = res
                with st.expander("🎯 点击查看AI诊断与正确答案（请先自己思考3秒再点开！）"):
                    st.markdown(apply_highlight(res, hl_core, teacher_color), unsafe_allow_html=True)
                anki_csv = f"{wrong_input}|{res}"
                apkg_data = generate_apkg(anki_csv)
                if apkg_data:
                    st.download_button("⬇️ 一键将本题生成Anki卡片 (.apkg)", data=apkg_data, file_name="错题卡片.apkg",
                                       mime="application/octet-stream", use_container_width=True)
                save_mistake(wrong_input, res, "文字")
                st.success("✅ 已存入本次会话的错题本！")

with tab3:
    st.subheader("📇 知识网络与卡片图鉴")
    st.caption("系统已自动整合刚才的【课程解构】与【错题诊断】，生成全局知识地图。")

    graph_choice = st.selectbox(
        "📌 请选择图谱生成模式",
        ["📖 生成整章全局知识框架（推荐，不管输入什么直接出图）", "🎯 针对某个特定知识点深入发散"]
    )

    if graph_choice == "📖 生成整章全局知识框架（推荐，不管输入什么直接出图）":
        user_topic = "输出整章的全局知识框架"
    else:
        user_topic = st.text_input("请输入你想深入探究的具体知识点（如：钠钾泵）")

    if st.button("🌐 生成整合知识网络图", use_container_width=True):
        with st.spinner("正在整合分析，绘制全局知识地图..."):
            graph_result = get_knowledge_graph(user_topic)
            st.subheader("🌐 综合知识网络图")

            if HAS_MERMAID:
                mermaid_code = graph_result.replace("```mermaid", "").replace("```", "").strip()
                st_mermaid(mermaid_code, height="500px")
            else:
                st.info("💡 复制下方代码，粘贴到 mermaid.live 中查看高清图表！薄弱知识点已用特殊形状标出。")
                st.code(graph_result, language="markdown")

with tab4:
    st.subheader("📕 我的错题本")
    st.caption("⚠️ 当前为会话临时存储。网页刷新后数据会重置，请及时导出备份。")
    mistakes = load_mistakes()
    if not mistakes:
        st.info("暂无错题记录。")
    else:
        for idx, item in enumerate(reversed(mistakes)):
            with st.expander(f"错题 {len(mistakes) - idx} | {item['time']} | 来源：{item['source']}"):
                st.markdown(f"**原题：**\n{item['question']}")
                st.markdown(f"**诊断：**\n{item['diagnosis']}")
        st.download_button("⬇️ 导出我的错题本 (JSON)", data=json.dumps(mistakes, ensure_ascii=False, indent=2),
                           file_name="my_mistake_book.json", mime="application/json", use_container_width=True)

    uploaded_json = st.file_uploader("导入你之前导出的错题本 (JSON)", type=['json'])
    if uploaded_json is not None:
        try:
            imported = json.load(uploaded_json)
            st.session_state['mistake_book'] = imported
            st.success(f"成功导入 {len(imported)} 条错题记录！")
            st.rerun()
        except Exception as e:
            st.error(f"导入失败：{e}")

with tab5:
    st.subheader("📂 我的分析历史归档")
    st.caption("⚠️ 当前为会话临时存储，刷新网页后数据会重置。")
    history = st.session_state['analysis_history']
    if not history:
        st.info("暂无历史记录。")
    else:
        for idx, item in enumerate(reversed(history)):
            with st.expander(f"【{item['type']}】 {item['time']}"):
                st.markdown(apply_highlight(item['content'], hl_core, teacher_color), unsafe_allow_html=True)

with tab6:
    st.subheader("💬 自由问答")
    st.caption("你可以直接向AI提问任何医学问题，它会优先基于你上传的PPT和笔记进行解答。")

    if st.sidebar.button("🗑️ 清空自由问答历史"):
        st.session_state.chat_history = []
        st.rerun()

    for chat in st.session_state.chat_history:
        with st.chat_message("user"): st.markdown(chat["user"])
        with st.chat_message("assistant"):
            st.markdown(apply_highlight(chat["assistant"], hl_core, teacher_color), unsafe_allow_html=True)

    if prompt := st.chat_input("输入你的医学问题..."):
        with st.chat_message("user"): st.markdown(prompt)
        with st.chat_message("assistant"):
            with st.spinner("AI正在思考中..."):
                response = chat_with_ai(prompt, st.session_state.chat_history)
                st.markdown(apply_highlight(response, hl_core, teacher_color), unsafe_allow_html=True)
        st.session_state.chat_history.append({"user": prompt, "assistant": response})