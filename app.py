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
    st.header("⚙️ 설정")
    
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
        # 쉼표와 줄바꿈 모두 처리
        raw_words = custom_dict_input.replace('\n', ',').split(',')
        custom_dict_list = [w.strip() for w in raw_words if w.strip()]

# 메인 영역
st.subheader("📁 1. PPTX 파일 업로드")
uploaded_file = st.file_uploader("검사할 파워포인트 파일을 올려주세요.", type=["pptx"])

if uploaded_file is not None:
    st.success(f"'{uploaded_file.name}' 업로드 성공!")
    
    if 'corrections' not in st.session_state:
        st.session_state.corrections = None
    if 'script_text' not in st.session_state:
        st.session_state.script_text = None
        
    prs = Presentation(uploaded_file)
        
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("🚀 AI 분석 및 텍스트 스캔 시작", use_container_width=True):
            # 이전 결과 초기화
            st.session_state.corrections = None
            st.session_state.script_text = None
            
            with st.spinner("PPT를 스캔하고 대본을 추출하는 중..."):
                # 1. 텍스트 추출 (내레이션)
                script_text = core.extract_narrations(prs)
                st.session_state.script_text = script_text
                
            st.success(f"화자별 대본 추출 완료! 이제 슬라이드 검사에 진입합니다.")
            
            # 2. OpenAI API 슬라이드별 개별 요청
            if not API_KEY_DEFAULT or not API_KEY_DEFAULT.startswith("sk-"):
                st.error("서버에 올바른 OpenAI API 환경변수 비밀키가 설정되어 있지 않습니다!")
                st.stop()
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            def update_progress(current, total):
                progress = int((current / total) * 100)
                progress_bar.progress(progress)
                status_text.markdown(f"**진행 상황:** {current}/{total} 장 스캔 완료... (최고 스피드 모드로 진행 중입니다)")
            
            with st.spinner("OpenAI 맞춤법 스캔 중 (슬라이드 한 장씩 꼼꼼히 검사합니다)..."):
                corrections = core.get_openai_corrections_by_slide(
                    prs, 
                    API_KEY_DEFAULT, 
                    is_paid_tier=True,
                    custom_dict=custom_dict_list,
                    progress_callback=update_progress
                )
                st.session_state.corrections = corrections
                
            progress_bar.progress(100)
            status_text.markdown("**✅ AI 분석 완료!**")

    # 결과물 표시
    if st.session_state.corrections is not None:
        st.subheader("📋 2. 수정 전 / 수정 후 검토")
        
        c_dict = st.session_state.corrections
        if len(c_dict) == 0:
            st.info("AI가 변경할 곳을 찾지 못했습니다. 문장이 이미 완벽하거나 수정할 내용이 없습니다.")
        else:
            df = pd.DataFrame(list(c_dict.items()), columns=["수정 전(원본)", "수정 후(AI 제안)"])
            st.dataframe(df, use_container_width=True, hide_index=True)
            
            st.warning("위 변경 사항들은 완성본 다운로드 시 '핫핑크색' 서식으로 PPT에 일괄 덮어씌워집니다.")
            
        st.subheader("📥 3. 완성본 다운로드")
        
        # 적용해서 메모리 스트림으로 저장
        with st.spinner("PPT 서식을 다시 그리는 중입니다..."):
            core.apply_corrections_to_ppt(prs, st.session_state.corrections)
            pptx_stream = io.BytesIO()
            prs.save(pptx_stream)
            pptx_stream.seek(0)
            
        # 대본 및 PPT 다운로드 버튼
        st.download_button(
            label="💖 색칠된 최종 PPTX 다운로드",
            data=pptx_stream,
            file_name=f"완료_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )
            
        st.subheader("📝 4. 화자별 대본 분리 다운로드 (교정 반영본)")
        
        # PPT 컨텍스트에 이미 반영된 텍스트를 재추출하여 가장 정확한 교정본 획득
        corrected_speakers = core.extract_narrations(prs)
        
        active_speakers = {k: v for k, v in corrected_speakers.items() if v}
        
        if not active_speakers:
            st.info("슬라이드 노트 표나 글상자에서 추출된 대본 텍스트가 없습니다.")
        else:
            # 화자 개수만큼 버튼을 가로로 정렬
            cols = st.columns(len(active_speakers))
            for idx, (speaker, lines) in enumerate(active_speakers.items()):
                with cols[idx]:
                    speaker_text = f"=== {speaker} 대본 ===\n\n" + "\n\n".join(lines)
                    st.download_button(
                        label=f"📥 [{speaker}] 대본 받기",
                        data=speaker_text.encode('utf-8'),
                        file_name=f"대본_{speaker}_{uploaded_file.name}.txt",
                        mime="text/plain",
                        use_container_width=True
                    )
