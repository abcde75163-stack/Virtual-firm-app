import html
import os
import re
import tempfile
import traceback
from pathlib import Path

import streamlit as st

from core.patent_parser import extract_text_from_file
from core.report_generator import generate_vf_report_streaming
from core.docx_builder import build_docx


MAX_PATENT_FILES = 10
MAX_PATENT_TOTAL_BYTES = 50 * 1024 * 1024


def _get_secret(name: str, default: str = "") -> str:
    """환경변수를 우선하고 Streamlit Secrets를 보조 설정으로 사용한다."""
    value = os.environ.get(name)
    if value:
        return value
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


def _extract_uploaded_file(uploaded_file) -> str:
    """업로드 파일을 임시 저장해 파싱하고 예외 여부와 무관하게 삭제한다."""
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=Path(uploaded_file.name).suffix,
        ) as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name
        return extract_text_from_file(tmp_path)
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)


def _report_base_name(patent_files) -> str:
    first_stem = Path(patent_files[0].name).stem if patent_files else "보고서"
    first_stem = re.sub(r'[\\/:*?"<>|]+', "_", first_stem).strip(" ._") or "보고서"
    if len(patent_files) > 1:
        return f"{first_stem}_외{len(patent_files) - 1}건"
    return first_stem

# ── 페이지 설정 ──────────────────────────────────────────
st.set_page_config(
    page_title="VF TBD 보고서 생성기",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS ─────────────────────────────────────────────────
st.markdown("""
<style>
    /* 전체 폰트 */
    html, body, [class*="css"] { font-family: 'Malgun Gothic', sans-serif; }

    /* 헤더 */
    .main-header {
        background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%);
        padding: 2rem 2.5rem;
        border-radius: 12px;
        margin-bottom: 2rem;
        color: white;
    }
    .main-header h1 { color: white; margin: 0; font-size: 1.8rem; }
    .main-header p  { color: #BDD7EE; margin: 0.3rem 0 0; font-size: 0.95rem; }

    /* 섹션 카드 */
    .section-card {
        background: #f8f9fa;
        border: 1px solid #dee2e6;
        border-radius: 10px;
        padding: 1.5rem;
        margin-bottom: 1.2rem;
    }
    .section-title {
        font-size: 1rem;
        font-weight: 700;
        color: #1F3864;
        margin-bottom: 0.8rem;
        padding-bottom: 0.4rem;
        border-bottom: 2px solid #2E75B6;
    }
    .optional-badge {
        display: inline-block;
        background: #e9ecef;
        color: #6c757d;
        font-size: 0.72rem;
        padding: 1px 7px;
        border-radius: 10px;
        margin-left: 8px;
        vertical-align: middle;
    }

    /* 진행 상황 로그 */
    .progress-log {
        background: #0d1117;
        color: #58a6ff;
        border-radius: 8px;
        padding: 1.2rem 1.5rem;
        font-family: 'Courier New', monospace;
        font-size: 0.82rem;
        line-height: 1.7;
        max-height: 320px;
        overflow-y: auto;
        border: 1px solid #30363d;
    }
    .log-success { color: #3fb950; }
    .log-warn    { color: #d29922; }
    .log-error   { color: #f85149; }
    .log-info    { color: #58a6ff; }
    .log-section { color: #e3b341; font-weight: bold; }

    /* 에러 박스 */
    .error-box {
        background: #fff5f5;
        border: 1px solid #f5c6cb;
        border-left: 4px solid #dc3545;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        margin-top: 1rem;
    }
    .error-box pre {
        font-size: 0.78rem;
        color: #721c24;
        margin: 0.5rem 0 0;
        white-space: pre-wrap;
        word-break: break-all;
    }

    /* 다운로드 버튼 */
    .stDownloadButton > button {
        background: #1F3864 !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.7rem 2rem !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
        width: 100%;
        transition: background 0.2s;
    }
    .stDownloadButton > button:hover { background: #2E75B6 !important; }

    /* 생성 버튼 */
    .stButton > button {
        background: #2E75B6 !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.65rem 1.5rem !important;
        font-weight: 600 !important;
    }
    .stButton > button:hover { background: #1F3864 !important; }

    /* 성공 박스 */
    .success-box {
        background: #f0fff4;
        border: 1px solid #9be9a8;
        border-left: 4px solid #2ea043;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        margin-top: 1rem;
    }

    /* 파일 업로드 힌트 */
    .upload-hint {
        font-size: 0.78rem;
        color: #6c757d;
        margin-top: 0.3rem;
    }

    /* 스텝 인디케이터 */
    .step-bar {
        display: flex;
        gap: 0;
        margin-bottom: 1.5rem;
    }
    .step-item {
        flex: 1;
        text-align: center;
        padding: 0.5rem;
        font-size: 0.8rem;
        background: #e9ecef;
        color: #6c757d;
        border-right: 1px solid #dee2e6;
    }
    .step-item:first-child { border-radius: 8px 0 0 8px; }
    .step-item:last-child  { border-radius: 0 8px 8px 0; border-right: none; }
    .step-item.active  { background: #2E75B6; color: white; font-weight: 700; }
    .step-item.done    { background: #1F3864; color: #BDD7EE; }
</style>
""", unsafe_allow_html=True)


# ── 헤더 ────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📋 Virtual Firm TBD 보고서 생성기</h1>
    <p>부산대학교 산학협력단 TLO | 특허 명세서 기반 기술사업화 보고서 자동 생성 시스템 v1.0</p>
</div>
""", unsafe_allow_html=True)


# ── 세션 상태 초기화 ─────────────────────────────────────
if "logs"           not in st.session_state: st.session_state.logs = []
if "report_text"    not in st.session_state: st.session_state.report_text = ""
if "docx_bytes"     not in st.session_state: st.session_state.docx_bytes = None
if "error_detail"   not in st.session_state: st.session_state.error_detail = ""
if "generating"     not in st.session_state: st.session_state.generating = False
if "done"           not in st.session_state: st.session_state.done = False
if "download_filename" not in st.session_state: st.session_state.download_filename = "VF_TBD_보고서_v3.docx"


# ── API 키 ───────────────────────────────────────────────
api_key = _get_secret("OPENAI_API_KEY")
openai_model = _get_secret("OPENAI_MODEL", "gpt-6-astra")
if not api_key:
    st.warning("⚠️ OPENAI_API_KEY가 설정되지 않았습니다. Streamlit Cloud Secrets에서 설정해 주세요.")


# ═══════════════════════════════════════════════════════
# 입력 영역
# ═══════════════════════════════════════════════════════
col_left, col_right = st.columns([1, 1], gap="large")

with col_left:
    # ── STEP 1: 특허 명세서 ──
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">① 특허 명세서 업로드 <span style="color:#dc3545;font-size:0.8rem;">필수</span></div>', unsafe_allow_html=True)

    patent_files = st.file_uploader(
        "특허 명세서 파일을 한 개 이상 업로드하세요",
        type=["pdf", "docx"],
        key="patent_upload",
        accept_multiple_files=True,
        help=f"PDF 또는 Word(.docx) 형식의 특허 명세서를 최대 {MAX_PATENT_FILES}개까지 업로드하세요.",
    )
    st.markdown('<p class="upload-hint">지원 형식: PDF, DOCX · 최대 10개 · 전체 50MB</p>', unsafe_allow_html=True)

    patent_files = patent_files or []
    patent_total_bytes = sum(uploaded.size for uploaded in patent_files)
    patent_upload_error = ""
    if len(patent_files) > MAX_PATENT_FILES:
        patent_upload_error = f"특허 명세서는 최대 {MAX_PATENT_FILES}개까지 업로드할 수 있습니다."
    elif patent_total_bytes > MAX_PATENT_TOTAL_BYTES:
        patent_upload_error = "특허 명세서 전체 용량은 50MB 이하여야 합니다."

    if patent_upload_error:
        st.error(f"❌ {patent_upload_error}")
    elif patent_files:
        st.success(
            f"✅ 특허 명세서 {len(patent_files)}개 업로드 완료 "
            f"(총 {patent_total_bytes / 1024 / 1024:.1f}MB)"
        )
        for uploaded in patent_files:
            st.caption(f"• {uploaded.name} ({uploaded.size // 1024:,}KB)")

    st.markdown('</div>', unsafe_allow_html=True)

    # ── STEP 2: 보고서 유형 선택 ──
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">② 사업화 유형 설정</div>', unsafe_allow_html=True)

    biz_type = st.selectbox(
        "사업화 유형을 선택하세요",
        ["자동 판단 (AI 추론)", "기술이전형 우선", "기술창업형 우선", "기술이전 + 창업 모두 검토"],
        key="biz_type",
        help="미지정 시 특허 명세서를 분석하여 AI가 자동으로 판단합니다."
    )

    report_version = st.selectbox(
        "보고서 버전",
        ["VF TBD v3.0 (최신)"],
        key="report_version"
    )
    st.markdown('</div>', unsafe_allow_html=True)


with col_right:
    # ── STEP 3: 수요기업 정보 (선택) ──
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">③ 수요기업 정보 <span class="optional-badge">선택사항</span></div>',
        unsafe_allow_html=True
    )
    st.caption("수요기업 정보를 입력하면 해당 기업 맞춤형 보고서가 생성됩니다.")

    demand_company = st.text_input(
        "수요기업명",
        placeholder="예: (주)OO배터리, LG에너지솔루션 등",
        key="demand_company"
    )

    demand_ir_file = st.file_uploader(
        "수요기업 IR자료 또는 사업계획서",
        type=["pdf", "docx", "txt"],
        key="demand_ir",
        help="수요기업의 IR자료, 사업소개서, 제품 카탈로그 등을 업로드하면 보고서에 반영됩니다."
    )
    st.markdown('<p class="upload-hint">지원 형식: PDF, DOCX, TXT · 최대 20MB</p>', unsafe_allow_html=True)

    if demand_ir_file:
        st.success(f"✅ {demand_ir_file.name} ({demand_ir_file.size // 1024}KB) 업로드 완료")

    if demand_company or demand_ir_file:
        st.info(f"💡 수요기업 맞춤형 보고서로 생성됩니다: **{demand_company or '(이름 미입력)'}**")

    st.markdown('</div>', unsafe_allow_html=True)

    # ── STEP 4: 추가 메모 ──
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">④ 추가 지시사항 <span class="optional-badge">선택사항</span></div>',
        unsafe_allow_html=True
    )
    extra_note = st.text_area(
        "보고서 작성 시 강조하거나 추가할 내용",
        placeholder="예: 전고체 전지 응용 분야를 중점적으로 다뤄주세요.\n예: 중국 시장 기술이전 가능성을 반드시 포함해주세요.",
        height=100,
        key="extra_note"
    )
    st.markdown('</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════
# 생성 버튼
# ═══════════════════════════════════════════════════════
st.markdown("---")
btn_col1, btn_col2, btn_col3 = st.columns([1, 2, 1])

with btn_col2:
    generate_btn = st.button(
        "🚀 VF TBD 보고서 생성 시작",
        disabled=(not patent_files or bool(patent_upload_error) or not api_key or st.session_state.generating),
        use_container_width=True,
        key="generate_btn"
    )
    if not patent_files:
        st.markdown('<p style="text-align:center;color:#dc3545;font-size:0.82rem;margin-top:0.3rem;">특허 명세서를 한 개 이상 업로드해 주세요</p>', unsafe_allow_html=True)
    if not api_key:
        st.markdown('<p style="text-align:center;color:#dc3545;font-size:0.82rem;margin-top:0.3rem;">API 키가 설정되지 않았습니다</p>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════
# 생성 실행
# ═══════════════════════════════════════════════════════
if generate_btn and patent_files and not patent_upload_error and api_key:
    # 상태 초기화
    st.session_state.logs        = []
    st.session_state.report_text = ""
    st.session_state.docx_bytes  = None
    st.session_state.error_detail = ""
    st.session_state.generating  = True
    st.session_state.done        = False

    st.markdown("---")
    st.markdown("### ⚙️ 보고서 생성 진행 상황")

    progress_bar  = st.progress(0, text="준비 중...")
    log_container = st.empty()
    status_text   = st.empty()

    def add_log(msg: str, level: str = "info"):
        css_class = {
            "info":    "log-info",
            "success": "log-success",
            "warn":    "log-warn",
            "error":   "log-error",
            "section": "log-section",
        }.get(level, "log-info")
        safe_msg = html.escape(msg)
        st.session_state.logs.append(f'<span class="{css_class}">{safe_msg}</span>')
        log_html = "<br>".join(st.session_state.logs)
        log_container.markdown(
            f'<div class="progress-log">{log_html}</div>',
            unsafe_allow_html=True
        )

    try:
        # ── STEP 1: 파일 파싱 ──────────────────────────
        add_log("▶ [1/4] 파일 파싱 시작...", "section")
        progress_bar.progress(5, text="특허 명세서 파싱 중...")

        patent_documents = []
        for index, patent_file in enumerate(patent_files, start=1):
            add_log(f"  ▷ [{index}/{len(patent_files)}] {patent_file.name} 파싱 중...", "info")
            patent_text = _extract_uploaded_file(patent_file)
            patent_documents.append((patent_file.name, patent_text))
            char_count = len(patent_text)
            add_log(f"  ✓ {patent_file.name} ({char_count:,}자 추출)", "success")
            if char_count < 500:
                add_log(f"  ⚠ {patent_file.name}: 추출된 텍스트가 너무 짧습니다.", "warn")

        total_chars = sum(len(text) for _, text in patent_documents)
        add_log(
            f"  ✓ 특허 명세서 {len(patent_documents)}개 파싱 완료 (총 {total_chars:,}자)",
            "success",
        )

        # 수요기업 IR 파싱 (선택)
        demand_text = ""
        if demand_ir_file:
            add_log("  ▷ 수요기업 IR자료 파싱 중...", "info")
            demand_text = _extract_uploaded_file(demand_ir_file)
            add_log(f"  ✓ 수요기업 IR자료 파싱 완료 ({len(demand_text):,}자)", "success")

        progress_bar.progress(15, text="파싱 완료 — AI 분석 준비 중...")

        # ── STEP 2: AI 보고서 생성 ──────────────────────
        add_log(f"▶ [2/4] GPT 보고서 생성 시작 ({openai_model}, 4섹션 분할 생성)...", "section")
        add_log("  ℹ 분량에 따라 3~7분 소요될 수 있습니다.", "info")

        sections_done = 0
        full_report   = []

        section_labels = [
            ("Ⅰ. 기술 개요",        20, 40),
            ("Ⅱ. 문제점 및 해결방안", 40, 60),
            ("Ⅲ. Scale-up 전략",     60, 80),
            ("Ⅳ. Virtual Firm 활용", 80, 95),
        ]

        for section_name, prog_start, prog_end in section_labels:
            add_log(f"  ▷ {section_name} 생성 중...", "info")
            progress_bar.progress(prog_start, text=f"{section_name} 생성 중...")

            section_text = ""
            token_count  = 0

            for chunk in generate_vf_report_streaming(
                api_key        = api_key,
                patent_documents = patent_documents,
                section_name   = section_name,
                prev_sections  = "\n\n".join(full_report),
                biz_type       = biz_type,
                demand_company = demand_company,
                demand_text    = demand_text,
                extra_note     = extra_note,
                model          = openai_model,
            ):
                section_text += chunk
                token_count  += len(chunk)

                # 500자마다 진행 도트 표시
                if token_count % 500 < 10:
                    dots = "·" * ((token_count // 500) % 5 + 1)
                    status_text.markdown(f"*{section_name} 생성 중{dots}*")

            section_text = section_text.strip()
            expected_heading = f"# {section_name}"
            if not section_text.startswith(expected_heading):
                section_text = f"{expected_heading}\n\n{section_text}"
            full_report.append(section_text)
            sections_done += 1
            add_log(f"  ✓ {section_name} 완료 ({len(section_text):,}자)", "success")
            progress_bar.progress(prog_end, text=f"{section_name} 완료")

        st.session_state.report_text = "\n\n".join(full_report)
        add_log(f"  ✓ 전체 보고서 생성 완료 (총 {len(st.session_state.report_text):,}자)", "success")

        # ── STEP 3: docx 변환 ───────────────────────────
        add_log("▶ [3/4] Word 문서(.docx) 변환 중...", "section")
        progress_bar.progress(96, text="Word 문서 변환 중...")

        patent_names = [Path(uploaded.name).stem for uploaded in patent_files]
        patent_name = ", ".join(patent_names)
        docx_bytes = build_docx(
            report_text    = st.session_state.report_text,
            patent_name    = patent_name,
            demand_company = demand_company,
        )
        st.session_state.docx_bytes = docx_bytes
        fname = f"VF_TBD_{_report_base_name(patent_files)}"
        if demand_company:
            safe_company = re.sub(r'[\\/:*?"<>|]+', "_", demand_company).strip(" ._")
            if safe_company:
                fname += f"_{safe_company}"
        st.session_state.download_filename = f"{fname}_v3.docx"
        add_log(f"  ✓ Word 문서 변환 완료 ({len(docx_bytes)//1024}KB)", "success")

        # ── STEP 4: 완료 ────────────────────────────────
        add_log("▶ [4/4] 최종 검토 완료", "section")
        progress_bar.progress(100, text="✅ 보고서 생성 완료!")
        status_text.empty()

        st.session_state.done      = True
        st.session_state.generating = False

        add_log("════════════════════════════════════════", "section")
        add_log("✅ VF TBD 보고서 생성이 완료되었습니다!", "success")
        add_log("   아래 다운로드 버튼을 클릭하세요.", "info")

    except Exception as e:
        st.session_state.generating  = False
        st.session_state.error_detail = traceback.format_exc()
        add_log(f"❌ 오류 발생: {str(e)}", "error")
        add_log("   자세한 오류 정보는 아래를 확인하세요.", "error")
        progress_bar.progress(0, text="❌ 오류 발생")

        safe_error = html.escape(str(e))
        safe_detail = html.escape(st.session_state.error_detail)
        st.markdown(f"""
        <div class="error-box">
            <strong>❌ 오류가 발생했습니다</strong>
            <p style="margin:0.5rem 0;color:#721c24;">{safe_error}</p>
            <details>
                <summary style="cursor:pointer;color:#721c24;font-size:0.85rem;">상세 오류 정보 보기</summary>
                <pre>{safe_detail}</pre>
            </details>
        </div>
        """, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════
# 다운로드 영역
# ═══════════════════════════════════════════════════════
if st.session_state.done and st.session_state.docx_bytes:
    st.markdown("---")
    st.markdown("### 📥 보고서 다운로드")

    dl_col1, dl_col2, dl_col3 = st.columns([1, 2, 1])
    with dl_col2:
        fname = st.session_state.download_filename

        st.download_button(
            label     = "📄 Word 보고서 다운로드 (.docx)",
            data      = st.session_state.docx_bytes,
            file_name = fname,
            mime      = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width = True,
        )

        st.markdown(f"""
        <div class="success-box">
            <strong>✅ 보고서 생성 완료</strong><br>
            <span style="font-size:0.88rem;color:#155724;">
            파일명: <code>{fname}</code><br>
            총 {len(st.session_state.report_text):,}자 · {len(st.session_state.docx_bytes)//1024}KB<br>
            ※ 본 보고서는 참고용 자료입니다. 기술이전 협상 전 전문가 검토를 권장합니다.
            </span>
        </div>
        """, unsafe_allow_html=True)

    # 보고서 미리보기 (접기)
    with st.expander("📖 보고서 내용 미리보기 (텍스트)"):
        st.text_area(
            "생성된 보고서 내용",
            value  = st.session_state.report_text[:5000] + "\n\n... (전체 내용은 Word 파일을 확인하세요)",
            height = 400,
        )


# ── 푸터 ────────────────────────────────────────────────
st.markdown("""
<hr style="margin-top:3rem;">
<p style="text-align:center;color:#6c757d;font-size:0.8rem;">
    부산대학교 산학협력단 기술사업화팀 (TLO) | VF TBD Report Generator v1.0<br>
    본 시스템은 Claude AI 기반으로 생성된 참고용 보고서를 제공합니다.
    실제 기술이전 협상 전 반드시 전문가 검토를 받으시기 바랍니다.
</p>
""", unsafe_allow_html=True)
