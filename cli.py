import os
import sys
from pptx import Presentation
import core
from dotenv import load_dotenv

load_dotenv()

# OpenAI 기본 키 (.env 또는 환경변수에서 로드)
API_KEY = os.environ.get("OPENAI_API_KEY", "")

def process_file_cli(input_path, is_paid_tier=True):
    print(f"\n[{input_path}] OpenAI GPT-4o-mini AI 스캐닝(문서별) 교정을 시작합니다...")
    if is_paid_tier:
        print("  ▶ [유료 요금제 모드] API 딜레이 대기 없이 최고 속도로 스캔합니다.")
    if not os.path.exists(input_path):
        print(f"에러: {input_path} 파일이 없습니다.")
        return
        
    is_pdf = input_path.lower().endswith(".pdf")
    base_name = os.path.splitext(input_path)[0]
    out_script = f"{base_name}_대본(통합).txt"
    
    if is_pdf:
        import fitz
        doc_obj = fitz.open(input_path)
        out_file = f"{base_name}_통합교정완료.pdf"
    else:
        doc_obj = Presentation(input_path)
        out_file = f"{base_name}_통합교정완료.pptx"
    
    # 1. 안내 메시지 (대본 추출은 최종 병합 후 수행)
    print(f" - PPT 내부 텍스트 스캔 및 슬라이드별 AI 맞춤법 교정안 도출 중...")

    # 2. OpenAI API에 슬라이드 단위 개별 요청
    print(f" - PPT 내부 텍스트 스캔 및 슬라이드별 AI 맞춤법 교정안 도출 중...")
    
    # 맞춤법 사전 파일 읽기
    dict_file_path = "맞춤법사전.txt"
    custom_dict_list = []
    if os.path.exists(dict_file_path):
        with open(dict_file_path, "r", encoding="utf-8") as f:
            dict_text = f.read()
            raw_words = dict_text.replace('\n', ',').split(',')
            custom_dict_list = [w.strip() for w in raw_words if w.strip()]
        if custom_dict_list:
            print(f"   > [맞춤법사전.txt] 파일에서 {len(custom_dict_list)}개의 예외 단어를 불러왔습니다.")
    
    def on_progress(current, total):
        print(f"   > 페이지/슬라이드 스캔 현황: {current}/{total} 장 완료...")
        
    if is_pdf:
        corrections_dict = core.get_openai_corrections_by_page_pdf(
            doc_obj, 
            API_KEY, 
            is_paid_tier=is_paid_tier,
            custom_dict=custom_dict_list,
            progress_callback=on_progress
        )
    else:
        corrections_dict = core.get_openai_corrections_by_slide(
            doc_obj, 
            API_KEY, 
            is_paid_tier=is_paid_tier,
            custom_dict=custom_dict_list,
            progress_callback=on_progress
        )
    
    print(f"\n   [AI 분석 결과] 총 {len(corrections_dict)}개의 문장이 교정 대상으로 식별되었습니다.")
    for k, v in corrections_dict.items():
        print(f"     - [수정 전]: {k}\n     - [수정 후]: {v}\n")
        
    # 3. 문서에 서식 덧씌우기
    print(f"\n - 문서에 교정 서식/코멘트 마커 반영 중...")
    if is_pdf:
        core.apply_corrections_to_pdf(doc_obj, corrections_dict)
    else:
        core.apply_corrections_to_ppt(doc_obj, corrections_dict)
    
    # 4. 교정된 대본 추출 및 저장
    print(f" - 교정이 반영된 최종 대본 분리 저장 중...")
    
    if is_pdf:
        raw_narrations = core.extract_narrations_pdf(doc_obj)
        corrected_narrations = {}
        for speaker, lines in raw_narrations.items():
            new_lines = []
            for line in lines:
                new_line = line
                for old_txt, new_txt in corrections_dict.items():
                    if old_txt in new_line:
                        new_line = new_line.replace(old_txt, new_txt)
                new_lines.append(new_line)
            corrected_narrations[speaker] = new_lines
    else:
        corrected_narrations = core.extract_narrations(doc_obj)
        
    for speaker, lines in corrected_narrations.items():
        if lines:
            speaker_script_path = f"{base_name}_대본({speaker}).txt"
            speaker_text = f"=== {speaker} 대본 ===\n\n" + "\n\n".join(lines)
            with open(speaker_script_path, 'w', encoding='utf-8') as f:
                f.write(speaker_text)
            print(f"   > [대본 저장 완료] {speaker_script_path}")
            
    doc_obj.save(out_file)
    if is_pdf:
        doc_obj.close()
    print(f"\n완료! AI 교정된 파일이 저장되었습니다: {out_file}\n")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("사용법: python cli.py [PPTX/PDF 파일경로]")
        print("예시: python cli.py \"한기대 보드.pptx\"")
        print("안내: 현재 유료 API 계정 모드가 기본값으로 적용되어 최고 속도로 동작합니다.")
    else:
        args = sys.argv[1:]
        # 기본값을 유료 모드(True)로 설정
        is_paid = True 
        if "--free" in args:
            is_paid = False
            args.remove("--free")
            
        # 기존 --paid가 들어와도 무시 (이미 True이므로)
        if "--paid" in args:
            args.remove("--paid")
            
        for f in args:
            process_file_cli(f, is_paid_tier=is_paid)
