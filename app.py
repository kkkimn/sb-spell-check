import os
import streamlit as st
import pandas as pd
from pptx import Presentation
import core
import io
from dotenv import load_dotenv

load_dotenv() # .env 파일이 존재하면 로컬 환경변수로 불러옴

# 기본 키 (금고 st.secrets 또는 .env 환경변수에서 우선 가져오기)
API_KEY_DEFAULT = ""
try:
    API_KEY_DEFAULT = st.secrets.get("OPENAI_API_KEY", "")
except Exception:
    pass

if not API_KEY_DEFAULT:
    API_KEY_DEFAULT = os.environ.get("OPENAI_API_KEY", "")

st.set_page_config(page_title="PPT 맞춤법 & 대본 AI", page_icon="✨", layout="wide")

st.title("✨ PPT 맞춤법 & 내레이션 자동 완성 솔루션 (Web 버전)")
st.markdown("압도적 성능의 **OpenAI (GPT-4o)** AI를 사용하여 PPT 문맥을 파악하고 맞춤법을 전수 검사합니다.")

# 사이드바
with st.sidebar:
    st.subheader("⚙️ AI 모델 선택")
    model_choice = st.radio(
        "정확도와 속도/비용 사이에서 선택하세요.",
        options=["꼼꼼 모드 (gpt-4o)", "빠른 모드 (gpt-4o-mini)"],
        index=0,
        help="꼼꼼 모드는 한국어 맞춤법·띄어쓰기·외래어 표기를 훨씬 정확하게 잡아냅니다. "
             "빠른 모드는 5~10배 저렴하지만 정확도가 떨어집니다."
    )
    selected_model = "gpt-4o" if "gpt-4o)" in model_choice else "gpt-4o-mini"
    
    st.divider()
    st.subheader("📖 사용자 맞춤법 사전")
    
    dict_file_path = "맞춤법사전.txt"
    default_dict_text = "챗지피티\nAI교수님\n한기대"
    
    if os.path.exists(dict_file_path):
        with open(dict_file_path, "r", encoding="utf-8") as f:
            default_dict_text = f.read()
            
    custom_dict_input = st.text_area(
        "AI가 절대 수정하면 안 되는 예외 단어를 쉼표(,)나 줄바꿈으로 적어주세요.",
        value=default_dict_text,
        height=150,
        help="여기에 적힌 단어는 맞춤법사전.txt 파일과 동기화됩니다."
    )
    
    if st.button("💾 사전 파일(`맞춤법사전.txt`)에 저장"):
        with open(dict_file_path, "w", encoding="utf-8") as f:
            f.write(custom_dict_input)
        st.success("✔ 성공적으로 파일에 저장되었습니다!")
    
    # 텍스트 에어리어 입력을 리스트로 변환
    custom_dict_list = []
    if custom_dict_input.strip():
        raw_words = custom_dict_input.replace('\n', ',').split(',')
        custom_dict_list = [w.strip() for w in raw_words if w.strip()]

# 메인 영역
st.subheader("📁 1. 파일 업로드 (PPTX / PDF 지원)")
uploaded_file = st.file_uploader("검사할 파워포인트 또는 PDF 파일을 올려주세요.", type=["pptx", "pdf"])

if uploaded_file is not None:
    st.success(f"'{uploaded_file.name}' 업로드 성공!")
    
    if 'corrections' not in st.session_state:
        st.session_state.corrections = None
    if 'script_text' not in st.session_state:
        st.session_state.script_text = None
        
    is_pdf = uploaded_file.name.lower().endswith('.pdf')
    
    # 업로드된 파일을 메모리 기반 객체로 로드
    if is_pdf:
        import fitz
        file_bytes = uploaded_file.read()
        doc_obj = fitz.open(stream=file_bytes, filetype="pdf")
    else:
        doc_obj = Presentation(uploaded_file)
        
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("🚀 AI 분석 및 텍스트 스캔 시작", use_container_width=True):
            st.session_state.corrections = None
            st.session_state.script_text = None
            
            with st.spinner("문서를 스캔하고 대본을 추출하는 중..."):
                if is_pdf:
                    script_text = core.extract_narrations_pdf(doc_obj)
                else:
                    script_text = core.extract_narrations(doc_obj)
                st.session_state.script_text = script_text
                
            st.success(f"대본 추출 완료! 이제 문서 검사에 진입합니다.")
            
            if not API_KEY_DEFAULT or not API_KEY_DEFAULT.startswith("sk-"):
                st.error("서버에 올바른 OpenAI API 환경변수 비밀키가 설정되어 있지 않습니다!")
                st.stop()
                
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            def update_progress(current, total):
                progress = int((current / total) * 100)
                progress_bar.progress(progress)
                status_text.markdown(f"**진행 상황:** {current}/{total} 페이지/슬라이드 스캔 완료... ({selected_model} 사용 중)")
            
            with st.spinner(f"OpenAI 맞춤법 스캔 중 ({selected_model})..."):
                if is_pdf:
                    corrections = core.get_openai_corrections_by_page_pdf(
                        doc_obj, 
                        API_KEY_DEFAULT, 
                        is_paid_tier=True,
                        custom_dict=custom_dict_list,
                        progress_callback=update_progress,
                        model=selected_model
                    )
                else:
                    corrections = core.get_openai_corrections_by_slide(
                        doc_obj, 
                        API_KEY_DEFAULT, 
                        is_paid_tier=True,
                        custom_dict=custom_dict_list,
                        progress_callback=update_progress,
                        model=selected_model
                    )
                st.session_state.corrections = corrections
                
            progress_bar.progress(100)
            status_text.markdown("**✅ AI 분석 완료!**")

    if st.session_state.corrections is not None:
        st.subheader("📋 2. 수정 전 / 수정 후 검토")
        
        c_dict = st.session_state.corrections
        if len(c_dict) == 0:
            st.info("AI가 변경할 곳을 찾지 못했습니다. 문장이 이미 완벽하거나 수정할 내용이 없습니다.")
        else:
            df = pd.DataFrame(list(c_dict.items()), columns=["수정 전(원본)", "수정 후(AI 제안)"])
            st.dataframe(df, use_container_width=True, hide_index=True)
            
            if is_pdf:
                st.warning("위 변경 사항들은 완성본 다운로드 시 '핫핑크색 형광펜 (메모 코멘트)' 형태로 PDF에 표시됩니다.")
            else:
                st.warning("위 변경 사항들은 완성본 다운로드 시 '핫핑크색' 서식으로 PPT에 일괄 덮어씌워집니다. "
                           "(부분 굵게/색상 등 일부 인라인 서식은 초기화될 수 있습니다.)")
            
        st.subheader("📥 3. 완성본 다운로드")
        
        with st.spinner("수정 및 덧그리기 작업 중입니다..."):
            out_stream = io.BytesIO()
            if is_pdf:
                core.apply_corrections_to_pdf(doc_obj, st.session_state.corrections)
                doc_obj.save(out_stream)
                doc_obj.close()
                mime_type = "application/pdf"
                btn_label = "💖 교정 하이라이트 PDF 다운로드"
                file_ext = "pdf"
            else:
                core.apply_corrections_to_ppt(doc_obj, st.session_state.corrections)
                doc_obj.save(out_stream)
                mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                btn_label = "💖 핑크색 교정 반영본 PPTX 다운로드"
                file_ext = "pptx"
                
            out_stream.seek(0)
            
        st.download_button(
            label=btn_label,
            data=out_stream,
            file_name=f"완료_{uploaded_file.name}",
            mime=mime_type,
            use_container_width=True
        )
            
        st.subheader("📝 4. 방송중고 대본 추출")
        st.markdown("대본 추출은 별도 도구에서 진행합니다. 아래 버튼을 누르면 새 탭에서 추출기가 열립니다.")
        
        html_source = "대본_추출기_통합.html"
        
        if not os.path.exists(html_source):
            st.error(f"'{html_source}' 파일을 찾을 수 없습니다. app.py와 같은 폴더에 두세요.")
        else:
            # HTML 파일 내용을 읽어서 JavaScript blob URL로 새 탭에 오픈
            # 이 방식은 로컬/배포(Streamlit Cloud, GitHub 등) 환경 모두에서 동일하게 작동함
            # (사용자 브라우저 측에서 직접 blob을 생성해 새 탭을 열기 때문)
            with open(html_source, "r", encoding="utf-8") as f:
                html_content = f.read()
            
            # JavaScript 문자열 리터럴 안전하게 인코딩 (백틱/역슬래시/줄바꿈 처리)
            import json
            html_js_string = json.dumps(html_content)
            
            # 버튼 HTML — 클릭 시 blob URL 생성 후 window.open()
            button_html = f"""
            <button id="open-script-extractor" style="
                background-color:#FF69B4;
                color:white;
                padding:0.9rem 1rem;
                border-radius:0.5rem;
                text-align:center;
                font-weight:bold;
                font-size:1.05rem;
                cursor:pointer;
                border:none;
                width:100%;
                font-family:inherit;
            ">
                🎙️ 방송중고 대본 추출 버튼
            </button>
            <script>
                (function() {{
                    const htmlContent = {html_js_string};
                    const btn = document.getElementById("open-script-extractor");
                    btn.addEventListener("click", function() {{
                        const blob = new Blob([htmlContent], {{ type: "text/html" }});
                        const blobUrl = URL.createObjectURL(blob);
                        const newWindow = window.open(blobUrl, "_blank");
                        if (!newWindow) {{
                            alert("팝업이 차단되었습니다. 브라우저의 팝업 차단을 해제해주세요.");
                        }}
                    }});
                }})();
            </script>
            """
            
            # HTML 컴포넌트로 렌더링 (iframe 안에서 실행되어 스크립트가 정상 동작)
            import streamlit.components.v1 as components
            components.html(button_html, height=70)
            
            st.caption("💡 버튼을 누르면 새 탭에서 대본 추출기가 열립니다. "
                      "만약 열리지 않으면 브라우저의 팝업 차단을 해제해주세요.")
