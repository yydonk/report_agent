import io
import json
import os
import tempfile

import streamlit as st

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="实验报告 Agent",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Theme CSS ─────────────────────────────────────────────────────────────────
_DARK = {
    "app_bg":         "#111111",
    "app_color":      "#f0f0f0",
    "sidebar_bg":     "#1c1c1c",
    "sidebar_border": "#2a2a2a",
    "sidebar_lbl":    "#aaaaaa",
    "card_bg":        "#1e1e1e",
    "card_border":    "#2d2d2d",
    "title_color":    "#ffffff",
    "status_bg":      "#1a2a3a",
    "status_border":  "#2a4a6a",
    "status_color":   "#c8dff0",
    "status_wait":    "#888888",
    "label_color":    "#cccccc",
    "hr_color":       "#2d2d2d",
    "dl_btn_bg":      "#2a5a2a",
    "input_bg":       "#2a2a2a",
    "input_color":    "#f0f0f0",
    "input_border":   "#444444",
    "deco_bg":        "#111111",   # top decoration bar matches background
}
_LIGHT = {
    "app_bg":         "#f5f6f8",
    "app_color":      "#1a1a1a",
    "sidebar_bg":     "#ffffff",
    "sidebar_border": "#e0e0e0",
    "sidebar_lbl":    "#666666",
    "card_bg":        "#ffffff",
    "card_border":    "#e2e4e8",
    "title_color":    "#111111",
    "status_bg":      "#eaf4fd",
    "status_border":  "#aed6f1",
    "status_color":   "#1a3a5a",
    "status_wait":    "#888888",
    "label_color":    "#444444",
    "hr_color":       "#e0e0e0",
    "dl_btn_bg":      "#276827",
    "input_bg":       "#ffffff",
    "input_color":    "#1a1a1a",
    "input_border":   "#cccccc",
    "deco_bg":        "#f5f6f8",
}

# Theme stored in session state (init before first CSS render)
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = True

def _inject_theme():
    t = _DARK if st.session_state.dark_mode else _LIGHT
    st.markdown(f"""
<style>
/* ── Remove top decoration bar ── */
[data-testid="stDecoration"] {{
    background: {t["deco_bg"]} !important;
    height: 2px !important;
}}
/* ── Remove top toolbar gap ── */
[data-testid="stHeader"] {{
    background: {t["deco_bg"]} !important;
}}

/* ── App background ── */
.stApp {{ background-color: {t["app_bg"]}; color: {t["app_color"]}; }}
[data-testid="stSidebar"] {{
    background-color: {t["sidebar_bg"]};
    border-right: 1px solid {t["sidebar_border"]};
}}

/* ── Sidebar labels ── */
.sidebar-section {{ font-size: 1.1rem; font-weight: 700; color: {t["title_color"]};
                    margin-top: 1.2rem; margin-bottom: 0.4rem; }}
.sidebar-sub {{ font-size: 0.82rem; color: {t["sidebar_lbl"]}; margin-bottom: 0.3rem; }}

/* ── Step cards ── */
.step-card {{
    background: {t["card_bg"]};
    border: 1px solid {t["card_border"]};
    border-radius: 10px;
    padding: 1.2rem 1.4rem;
    margin-bottom: 1.2rem;
    box-shadow: 0 1px 4px rgba(0,0,0,0.07);
}}
.step-title {{
    font-size: 1.25rem; font-weight: 700;
    color: {t["title_color"]}; margin-bottom: 0.6rem;
}}
.step-num {{
    display: inline-block;
    background: #e05252; color: #fff;
    border-radius: 50%; width: 1.6rem; height: 1.6rem;
    text-align: center; line-height: 1.6rem;
    font-size: 0.85rem; font-weight: 700;
    margin-right: 0.5rem;
}}

/* ── Status panel ── */
.status-panel {{
    background: {t["status_bg"]};
    border: 1px solid {t["status_border"]};
    border-radius: 10px;
    padding: 1rem 1.2rem;
    min-height: 120px;
    color: {t["status_color"]};
    font-size: 0.9rem;
}}
.status-ok   {{ color: #3a9a3a; font-weight: 600; }}
.status-wait {{ color: {t["status_wait"]}; }}
.status-err  {{ color: #e05252; font-weight: 600; }}

/* ── Primary action button ── */
div[data-testid="stButton"] > button[kind="primary"] {{
    background-color: #e05252 !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: 700 !important;
    width: 100% !important;
    color: #ffffff !important;
}}
div[data-testid="stButton"] > button[kind="primary"]:hover {{
    background-color: #c03030 !important;
}}

/* ── Download button ── */
div[data-testid="stDownloadButton"] > button {{
    background-color: {t["dl_btn_bg"]} !important;
    border: none !important;
    border-radius: 6px !important;
    color: #ffffff !important;
    font-weight: 600 !important;
    width: 100% !important;
}}

/* ── Input / textarea ── */
input, textarea, [data-baseweb="input"] input, [data-baseweb="textarea"] textarea {{
    background-color: {t["input_bg"]} !important;
    color: {t["input_color"]} !important;
    border-color: {t["input_border"]} !important;
}}
[data-baseweb="select"] > div {{
    background-color: {t["input_bg"]} !important;
    color: {t["input_color"]} !important;
    border-color: {t["input_border"]} !important;
}}

/* ── Labels & text ── */
label {{ color: {t["label_color"]} !important; }}
p, span, li {{ color: {t["app_color"]}; }}

/* ── Divider ── */
hr {{ border-color: {t["hr_color"]}; }}
</style>
""", unsafe_allow_html=True)

_inject_theme()


# ── Session state init ────────────────────────────────────────────────────────
for key, default in {
    "format_schema": None,
    "content_schema": None,
    "experiment": None,
    "draft_bytes": None,
    "final_bytes": None,
    "log": [],
}.items():
    if key not in st.session_state:
        st.session_state[key] = default


def _log(msg: str, kind: str = "info"):
    st.session_state.log.append((kind, msg))


def _clear_downstream(from_step: int):
    """Clear state for steps >= from_step to force re-generation."""
    if from_step <= 1:
        st.session_state.format_schema = None
        st.session_state.content_schema = None
    if from_step <= 2:
        st.session_state.experiment = None
    if from_step <= 3:
        st.session_state.draft_bytes = None
    if from_step <= 4:
        st.session_state.final_bytes = None


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    # Theme toggle row
    t_col1, t_col2 = st.columns([3, 1])
    t_col1.markdown('<div class="sidebar-section" style="margin-top:0.4rem">配置</div>', unsafe_allow_html=True)
    dark_toggle = t_col2.toggle("", value=st.session_state.dark_mode, key="theme_toggle",
                                help="切换暗色 / 亮色模式")
    if dark_toggle != st.session_state.dark_mode:
        st.session_state.dark_mode = dark_toggle
        st.rerun()

    # Re-inject theme after potential toggle (sidebar renders after main CSS)
    _inject_theme()

    st.markdown('<div class="sidebar-sub">API 密钥</div>', unsafe_allow_html=True)

    api_key = st.text_input(
        "DashScope API Key",
        value=os.getenv("DASHSCOPE_API_KEY", ""),
        type="password",
        label_visibility="collapsed",
        placeholder="sk-...",
    )
    base_url = st.text_input(
        "Base URL",
        value=os.getenv("DASHSCOPE_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"),
        label_visibility="visible",
    )

    st.markdown('<div class="sidebar-section">高级配置</div>', unsafe_allow_html=True)

    model_text = st.selectbox(
        "文本模型",
        ["qwen-plus", "qwen-max", "qwen-turbo", "deepseek-chat"],
        index=0,
    )
    model_vision = st.selectbox(
        "视觉模型",
        ["qwen-vl-plus", "qwen-vl-max"],
        index=0,
    )
    target_chars = st.slider(
        "目标生成字数",
        min_value=1000, max_value=10000, value=5000, step=500,
    )
    terminology = st.select_slider(
        "专业术语密度",
        options=["低", "中", "高"],
        value="中",
    )

    st.markdown('<div class="sidebar-section">报告信息</div>', unsafe_allow_html=True)
    stu_name     = st.text_input("姓名",     placeholder="张三",         key="stu_name")
    stu_id       = st.text_input("学号",     placeholder="2024000001",   key="stu_id")
    stu_course   = st.text_input("课程名称", placeholder="大学物理实验", key="stu_course")
    stu_teacher  = st.text_input("指导教师", placeholder="李老师",       key="stu_teacher")
    stu_date     = st.text_input("实验日期", placeholder="2025-03-10",   key="stu_date")
    add_fig_caps = st.toggle("自动添加图注编号", value=True, key="fig_caps")


# ── Helper: build DeepSeekClient ─────────────────────────────────────────────
def _make_client(model_name: str):
    from src.img_agent.deepseek_client import DeepSeekClient
    return DeepSeekClient(api_key=api_key, base_url=base_url, model=model_name)


# ── Main layout ───────────────────────────────────────────────────────────────
st.markdown("## 实验报告 Agent")
st.markdown("<span style='color:#888;font-size:0.95rem'>基于 AI 的实验报告自动生成系统</span>", unsafe_allow_html=True)
st.markdown("---")

col_main, col_status = st.columns([3, 1.2])

# ── Right: status panel ───────────────────────────────────────────────────────
with col_status:
    st.markdown("### 状态信息")
    status_box = st.empty()

    def _render_status():
        lines = st.session_state.log[-12:] if st.session_state.log else []
        if not lines:
            html = '<div class="status-panel"><span class="status-wait">尚未开始</span></div>'
        else:
            items = []
            for kind, msg in lines:
                css = {"ok": "status-ok", "err": "status-err"}.get(kind, "status-wait")
                items.append(f'<div class="{css}">{msg}</div>')
            html = '<div class="status-panel">' + "".join(items) + "</div>"
        status_box.markdown(html, unsafe_allow_html=True)

    _render_status()

# ── Left: steps ───────────────────────────────────────────────────────────────
with col_main:

    # ── Step 1: Template ──────────────────────────────────────────────────────
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown('<div class="step-title"><span class="step-num">1</span>模板输入</div>', unsafe_allow_html=True)
    st.markdown("上传实验报告 Word 模板（.docx），系统将自动解析格式规范与内容要求。")

    tpl_file = st.file_uploader("选择模板文件", type=["docx"], key="tpl_uploader", label_visibility="collapsed")

    if st.button("解析模板", type="primary", key="btn_parse_tpl"):
        if not tpl_file:
            st.warning("请先上传模板文件")
        elif not api_key:
            st.warning("请在左侧填写 API Key")
        else:
            _clear_downstream(1)
            st.session_state.log = []
            _log("正在解析模板...")
            _render_status()
            with st.spinner("解析模板中..."):
                try:
                    from src.doc_agent.word_rules import analyze_docx
                    from src.doc_agent.classify import refine_with_llm
                    from src.doc_agent.normalize import normalize_spec
                    from cli_build_schemas import build_format_schema, build_content_schema

                    tpl_bytes = tpl_file.read()
                    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
                        tmp.write(tpl_bytes)
                        tmp_path = tmp.name

                    base = analyze_docx(tmp_path)
                    client_txt = _make_client(model_text)
                    llm_obj = refine_with_llm(base, client_txt)
                    normalized = normalize_spec(base, llm_obj, tmp_path)
                    os.unlink(tmp_path)

                    fmt = build_format_schema([normalized])
                    cnt = build_content_schema([normalized])

                    # Serialize via json to strip non-serializable objects
                    st.session_state.format_schema = json.loads(json.dumps(fmt, ensure_ascii=False, default=str))
                    st.session_state.content_schema = json.loads(json.dumps(cnt, ensure_ascii=False, default=str))

                    n_chapters = len(cnt.get("chapters", {}))
                    _log(f"模板解析完成，识别 {n_chapters} 个章节", "ok")
                except Exception as e:
                    _log(f"解析失败：{e}", "err")
            _render_status()

    if st.session_state.format_schema:
        n = len(st.session_state.content_schema.get("chapters", {}))
        st.success(f"模板已就绪，共 {n} 个章节")

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Step 2: Experiment images ─────────────────────────────────────────────
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown('<div class="step-title"><span class="step-num">2</span>实验内容输入</div>', unsafe_allow_html=True)
    st.markdown("上传实验内容照片（支持多张），视觉模型将自动提取实验要素。")

    img_files = st.file_uploader(
        "选择实验图片", type=["png", "jpg", "jpeg"],
        accept_multiple_files=True, key="img_uploader",
        label_visibility="collapsed",
    )
    # Read bytes immediately — st.image() consumes the stream, so cache before display
    if img_files:
        img_cache = [(f.name, f.read()) for f in img_files]
        cols_prev = st.columns(min(len(img_cache), 4))
        for i, (name, data) in enumerate(img_cache[:4]):
            cols_prev[i].image(io.BytesIO(data), use_container_width=True)
    else:
        img_cache = []

    if st.button("提取实验内容", type="primary", key="btn_extract"):
        if not img_cache:
            st.warning("请先上传实验图片")
        elif not api_key:
            st.warning("请在左侧填写 API Key")
        else:
            _clear_downstream(2)
            _log("正在调用视觉模型提取实验内容...")
            _render_status()
            with st.spinner("提取实验内容中..."):
                try:
                    from src.img_agent.vision import _build_vision_messages, _img_to_data_url
                    import json as _json

                    tmp_paths = []
                    for name, data in img_cache:
                        suffix = os.path.splitext(name)[1] or ".jpg"
                        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
                            tmp.write(data)
                            tmp_paths.append(tmp.name)

                    # Verify image bytes are non-empty
                    for name, data in img_cache:
                        _log(f"图片 {name}：{len(data):,} bytes", "info")
                    for p in tmp_paths:
                        size = os.path.getsize(p)
                        _log(f"临时文件 {os.path.basename(p)}：{size:,} bytes", "info")

                    client_vis = _make_client(model_vision)
                    messages = _build_vision_messages(tmp_paths)
                    for p in tmp_paths:
                        os.unlink(p)

                    # Get raw response for diagnosis
                    raw_text = client_vis.chat(messages, temperature=0.1, max_tokens=4096).strip()
                    st.session_state["_vision_raw"] = raw_text  # store for display

                    # Try to extract JSON
                    s = raw_text.find("{")
                    e = raw_text.rfind("}")
                    json_str = raw_text[s:e+1] if s != -1 and e >= s else ""
                    try:
                        exp_dict = _json.loads(json_str) if json_str else {}
                    except Exception as je:
                        exp_dict = {}
                        _log(f"JSON解析失败：{je}", "err")
                        _log(f"原始响应前200字：{raw_text[:200]}", "err")

                    # Normalize to expected schema
                    if not isinstance(exp_dict, dict):
                        exp_dict = {}
                    if "steps" not in exp_dict:
                        exp_dict["steps"] = []
                    if "apparatus" not in exp_dict:
                        exp_dict["apparatus"] = []
                    if "data" not in exp_dict:
                        exp_dict["data"] = {}

                    st.session_state.experiment = exp_dict
                    n_steps = len(exp_dict.get("steps", []))
                    title_ok  = "✓" if exp_dict.get("title") else "✗"
                    obj_ok    = "✓" if exp_dict.get("objective") else "✗"
                    theory_ok = "✓" if exp_dict.get("theory") else "✗"
                    _log(f"提取完成 | 标题{title_ok} 目的{obj_ok} 原理{theory_ok} 步骤×{n_steps}", "ok")
                    if not any([exp_dict.get("title"), exp_dict.get("objective"),
                                exp_dict.get("theory"), n_steps]):
                        _log("警告：字段全空，见「原始响应」排查原因", "err")
                except Exception as ex:
                    _log(f"提取失败：{ex}", "err")
            _render_status()
            # Show raw API response for debugging when empty
            if st.session_state.get("_vision_raw"):
                with st.expander("原始 API 响应（调试用）"):
                    st.code(st.session_state["_vision_raw"][:2000], language="json")

    if st.session_state.experiment:
        exp = st.session_state.experiment
        has_content = any([exp.get("title"), exp.get("objective"), exp.get("theory"),
                           exp.get("steps"), exp.get("analysis")])
        if has_content:
            st.success(f"实验内容已就绪：{exp.get('title', '（无标题）')}")
        else:
            st.warning("视觉提取结果为空，请使用下方手动输入框补充实验内容")
        with st.expander("查看 / 编辑提取结果（JSON）"):
            edited = st.text_area("实验内容 JSON", value=json.dumps(exp, ensure_ascii=False, indent=2), height=260, key="exp_editor")
            if st.button("保存编辑", key="btn_save_exp"):
                try:
                    st.session_state.experiment = json.loads(edited)
                    st.session_state.draft_bytes = None
                    st.success("已保存")
                except Exception as e:
                    st.error(f"JSON 格式有误：{e}")

    # Manual input fallback when no images available
    with st.expander("手动输入实验内容（备用）"):
        manual_text = st.text_area(
            "粘贴或输入实验内容描述（自由格式）",
            height=160, key="manual_exp",
            placeholder="实验名称：\n实验目的：\n实验原理：\n实验步骤：\n..."
        )
        if st.button("使用手动输入", key="btn_manual_exp"):
            if manual_text.strip():
                st.session_state.experiment = {"title": "", "objective": "", "theory": "",
                                               "apparatus": [], "steps": [],
                                               "data": {}, "analysis": manual_text.strip()}
                st.session_state.draft_bytes = None
                _log("已使用手动输入的实验内容", "ok")
                _render_status()
                st.success("手动内容已保存")
            else:
                st.warning("请先输入内容")

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Step 3: Generate draft ────────────────────────────────────────────────
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown('<div class="step-title"><span class="step-num">3</span>生成草稿</div>', unsafe_allow_html=True)
    st.markdown("根据模板章节要求与实验内容生成纯文字草稿，下载后可在 Word 中插入图片并编辑。")

    # Chapter weight editor (only shown after schema is loaded)
    chapter_weights = {}
    if st.session_state.content_schema:
        from cli_generate_md_draft import _get_deduped_chapters
        all_chs = _get_deduped_chapters(st.session_state.content_schema)
        with st.expander("章节字数权重（可选，默认均等）"):
            st.caption("数值越大，该章节分配的字数越多。所有权重相对生效。")
            cols = st.columns(2)
            for i, ch in enumerate(all_chs):
                title = ch["title"]
                default = 3 if any(k in title for k in ["原理", "步骤", "数据", "分析", "误差"]) else 1
                w = cols[i % 2].slider(title, min_value=1, max_value=10, value=default, key=f"w_{i}")
                chapter_weights[title] = float(w)

    if st.button("生成草稿", type="primary", key="btn_draft"):
        if not st.session_state.content_schema:
            st.warning("请先完成步骤 1（解析模板）")
        elif not st.session_state.experiment:
            st.warning("请先完成步骤 2（提取实验内容）")
        elif not api_key:
            st.warning("请在左侧填写 API Key")
        else:
            _clear_downstream(3)
            _log(f"逐章节生成草稿（目标 {target_chars} 字）...")
            _render_status()
            progress = st.progress(0, text="准备中...")
            with st.spinner("逐章节生成中，请稍候..."):
                try:
                    from cli_generate_md_draft import (
                        generate_draft_chapters, write_draft_docx,
                        _get_deduped_chapters,
                    )

                    client_txt = _make_client(model_text)
                    experiment_text = json.dumps(st.session_state.experiment, ensure_ascii=False)
                    total_chapters = len(_get_deduped_chapters(st.session_state.content_schema))

                    def _progress_cb(i, total, title):
                        pct = int((i / total) * 95) + 2
                        progress.progress(pct, text=f"[{i+1}/{total}] {title}")
                        _log(f"[{i+1}/{total}] {title}", "info")
                        _render_status()

                    chapters = generate_draft_chapters(
                        st.session_state.content_schema,
                        experiment_text,
                        client_txt,
                        total_target=target_chars,
                        progress_cb=_progress_cb,
                        chapter_weights=chapter_weights or None,
                        terminology=terminology,
                    )

                    progress.progress(98, text="写入 Word 文件...")
                    buf = io.BytesIO()
                    write_draft_docx(chapters, buf)
                    st.session_state.draft_bytes = buf.getvalue()

                    total_actual = sum(len(c.get("content", "")) for c in chapters)
                    _log(f"草稿完成：{len(chapters)} 章节，实际约 {total_actual} 字", "ok")
                    progress.progress(100, text="完成")
                except Exception as e:
                    _log(f"草稿生成失败：{e}", "err")
                    progress.empty()
            _render_status()

    if st.session_state.draft_bytes:
        st.success("草稿已生成，点击下载后在 Word 中编辑并插入图片")
        st.download_button(
            label="下载草稿 draft.docx",
            data=st.session_state.draft_bytes,
            file_name="draft.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_draft",
        )

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Step 4: Final report ──────────────────────────────────────────────────
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown('<div class="step-title"><span class="step-num">4</span>生成最终实验报告</div>', unsafe_allow_html=True)
    st.markdown("上传编辑好的草稿（含插图），系统自动套用模板格式并调整图片大小，生成最终 Word 报告。")

    draft_upload = st.file_uploader("上传编辑后的草稿", type=["docx"], key="draft_uploader", label_visibility="collapsed")

    if st.button("生成最终报告", type="primary", key="btn_final"):
        if not st.session_state.format_schema:
            st.warning("请先完成步骤 1（解析模板）")
        elif not draft_upload:
            st.warning("请上传编辑好的草稿文件")
        else:
            _clear_downstream(4)
            _log("正在套用格式生成最终报告...")
            _render_status()
            with st.spinner("生成最终报告中..."):
                try:
                    from src.report_agent.renderer import render_report_from_draft

                    # Write draft to temp file
                    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_in:
                        tmp_in.write(draft_upload.read())
                        draft_path = tmp_in.name

                    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_out:
                        out_path = tmp_out.name

                    student_info = {
                        "name":       stu_name,
                        "student_id": stu_id,
                        "course":     stu_course,
                        "teacher":    stu_teacher,
                        "date":       stu_date,
                    }
                    render_report_from_draft(
                        draft_path,
                        st.session_state.format_schema,
                        out_path,
                        student_info=student_info,
                        add_figure_captions=add_fig_caps,
                    )

                    with open(out_path, "rb") as f:
                        st.session_state.final_bytes = f.read()

                    os.unlink(draft_path)
                    os.unlink(out_path)

                    _log("最终报告生成完成", "ok")
                except Exception as e:
                    _log(f"报告生成失败：{e}", "err")
            _render_status()

    if st.session_state.final_bytes:
        st.success("最终报告已生成")
        st.download_button(
            label="下载最终报告 final_report.docx",
            data=st.session_state.final_bytes,
            file_name="final_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_final",
        )

    st.markdown('</div>', unsafe_allow_html=True)

# Keep status in sync after reruns
with col_status:
    _render_status()
