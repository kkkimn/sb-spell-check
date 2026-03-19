import re
import json
import time
import difflib
from pptx import Presentation
from pptx.dml.color import RGBColor
from openai import OpenAI

def extract_narrations(prs):
    """
    PPT의 표와 슬라이드 노트를 분석하여 
    교수, 성우, 선생님, 기타 4가지 화자로 대본을 분리하여 텍스트 덩어리로 반환합니다.
    """
    narrations = {"교수": [], "성우": [], "선생님": [], "기타": []}
    slide_height = prs.slide_height
    
    for i, slide in enumerate(prs.slides):
        # 1) 표와 일반 텍스트 상자에서 내레이션 찾기 (하단 배치 여부 확인)
        for shape in slide.shapes:
            # 대본은 보통 슬라이드 하단 40% 영역에 배치된다고 가정
            is_bottom_shape = hasattr(shape, "top") and shape.top > slide_height * 0.6
            
            if shape.has_table:
                for row in shape.table.rows:
                    if len(row.cells) >= 3:
                        speaker_raw = row.cells[1].text.strip()
                        narration_text = row.cells[2].text.strip()
                        
                        speaker_found = None
                        if "교수" in speaker_raw: speaker_found = "교수"
                        elif "선생님" in speaker_raw: speaker_found = "선생님"
                        elif "성우" in speaker_raw: speaker_found = "성우"
                        
                        if speaker_found:
                            narrations[speaker_found].append(f"[슬라이드 {i+1}] {speaker_raw} :\n{narration_text}")
                        elif narration_text and (is_bottom_shape or len(speaker_raw) < 15):
                            # 화자를 특정 못해도 표가 하단에 있거나 화자명 길이가 짧으면 대본일 확률이 매우 높음
                            narrations["기타"].append(f"[슬라이드 {i+1}] {speaker_raw if speaker_raw else '텍스트'} :\n{narration_text}")
                    
                    elif len(row.cells) == 2:
                        speaker_raw = row.cells[0].text.strip()
                        narration_text = row.cells[1].text.strip()
                        
                        speaker_found = None
                        if "교수" in speaker_raw: speaker_found = "교수"
                        elif "선생님" in speaker_raw: speaker_found = "선생님"
                        elif "성우" in speaker_raw: speaker_found = "성우"
                        
                        if speaker_found:
                            narrations[speaker_found].append(f"[슬라이드 {i+1}] {speaker_raw} :\n{narration_text}")
                        elif narration_text and is_bottom_shape:
                            narrations["기타"].append(f"[슬라이드 {i+1}] 텍스트 :\n{narration_text}")
                            
            elif shape.has_text_frame:
                text = shape.text.strip()
                if not text: continue
                
                # '교수: ...' 형태 매칭
                match = re.match(r'^\s*(교수|성우|선생님|기타)님?\s*[:]\s*(.*)', text, flags=re.DOTALL)
                if match:
                    speaker_found = match.group(1)
                    narration_text = match.group(2).strip()
                    narrations[speaker_found].append(f"[슬라이드 {i+1}] {speaker_found} :\n{narration_text}")
                else:
                    # 매칭 안되더라도 하단에 있는 텍스트 상자면 기타 대본으로 수집
                    if is_bottom_shape and len(text) > 5 and text not in [line.split(':\n')[-1].strip() for lines in narrations.values() for line in lines]:
                        narrations["기타"].append(f"[슬라이드 {i+1}] 글상자 :\n{text}")
                            
        # 2) 슬라이드 노트에서 내레이션 찾기 (기본 내레이션 창구)
        if slide.has_notes_slide:
            notes_text = slide.notes_slide.notes_text_frame.text.strip()
            if notes_text:
                speaker_found = "기타"
                if "교수" in notes_text: speaker_found = "교수"
                elif "선생님" in notes_text: speaker_found = "선생님"
                elif "성우" in notes_text: speaker_found = "성우"
                
                narrations[speaker_found].append(f"[슬라이드 {i+1}] {speaker_found} (노트) :\n{notes_text}")

    return narrations


def get_openai_corrections_by_slide(prs, api_key, is_paid_tier=True, custom_dict=None, progress_callback=None):
    """
    슬라이드를 하나씩 읽어가면서 문맥을 바탕으로 OpenAI 교정안을 확보합니다.
    """
    client = OpenAI(api_key=api_key)
    global_corrections = {}
    
    # 사용자 정의 사전 프롬프트 설정
    custom_dict_prompt = ""
    if custom_dict and len(custom_dict) > 0:
        custom_dict_prompt = "\n\n[중요 예외 처리 규칙(사용자 맞춤법 사전)]\n다음 단어들은 사용자의 의도적인 예외 단어입니다. 이 단어들이 원문에 등장하면 절대로 띄어쓰기나 맞춤법을 수정하지 말고 원형 그대로 보존하세요:\n"
        custom_dict_prompt += ", ".join(custom_dict)
    
    total_slides = len(prs.slides)
    
    for i, slide in enumerate(prs.slides):
        # 해당 슬라이드의 모든 텍스트 긁기
        slide_texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                slide_texts.append(shape.text.strip())
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            slide_texts.append(cell.text.strip())
                            
        # 하단 슬라이드 노트(내레이션)도 검사 대상에 포함
        if slide.has_notes_slide:
            notes_text = slide.notes_slide.notes_text_frame.text.strip()
            if notes_text:
                slide_texts.append(notes_text)
                            
        full_text = "\n".join([t for t in slide_texts if len(t) > 1])
        
        if not full_text.strip():
            if progress_callback: progress_callback(i + 1, total_slides)
            continue
            
        system_prompt = "너는 파워포인트 슬라이드 텍스트의 맞춤법, 오탈자, 다중 띄어쓰기를 교정하는 한국어 AI야. 문맥을 바탕으로 자연스럽게 고치되, 반드시 원본과 수정본이 한 쌍을 이루는 '순수 JSON' 객체 형태로 응답해. 교정할 것이 없으면 {} 만 반환해.\n"
        system_prompt += "[매우 중요한 규칙]\n"
        system_prompt += "1. 화면 설명, UI 텍스트, 제목 등 개조식이나 명사형으로 끝나는 문장을 억지로 '~니다' 나 '~합니다' 모양의 완성형 문장으로 바꾸지 마세요. 오직 기본 어투를 유지한 채 '맞춤법'과 '띄어쓰기'만 교정하세요.\n"
        system_prompt += "2. PPT 하단의 내레이션이나 대본 텍스트도 포함되어 있으니 똑같이 오탈자를 점검하세요.\n"
        system_prompt += custom_dict_prompt
        user_prompt = f'형식 예시: {{"틀린원문1": "고친문장1"}}\n\n=== 슬라이드 {i+1} 텍스트 ===\n{full_text}'

        success = False
        for attempt in range(5): # OpenAI는 안정적이니 5번만
            try:
                response = client.chat.completions.create(
                    model="gpt-4o-mini", # 가성비가 매우 뛰어나고 빠름
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    response_format={"type": "json_object"}, # 완벽한 JSON 보장
                    temperature=0.1
                )
                
                res_text = response.choices[0].message.content.strip()
                slide_dict = json.loads(res_text)
                
                for k, v in slide_dict.items():
                    if k != v:  
                        # 이중 안전장치: 결과물(v)이 사용자 사전에 의해 교정되어서는 안 되는 단어인데 교정되었다면 기각
                        skip_correction = False
                        if custom_dict:
                            for word in custom_dict:
                                if word in k and word not in v:
                                    skip_correction = True
                                    break
                        
                        if not skip_correction:
                            global_corrections[k] = v
                success = True
                break
                
            except Exception as e:
                err_msg = str(e)
                if "rate limit" in err_msg.lower() or "429" in err_msg:
                    print(f"   [API 한도 초과] 5초간 대기 후 슬라이드 {i+1} 재시도... ({attempt+1}/5)")
                    time.sleep(5) 
                else:
                    print(f"   [API 오류] 재시도 중... 사유: {e}")
                    time.sleep(2)
                    
        # OpenAI는 유료 API 기준 Rate Limit이 매우 널널하므로 is_paid_tier가 True면 딜레이 없음
        # Free Tier(Tier 1)일 경우에만 미세한 1초 딜레이
        if success and not is_paid_tier:
            time.sleep(1)
            
        if progress_callback:
            progress_callback(i + 1, total_slides)
            
    return global_corrections



def apply_corrections_to_ppt(prs, corrections_dict):
    """
    맞춤법 교정 딕셔너리를 돌며 PPT 내부에 수정된 텍스트를 핫핑크색 서식으로 적용합니다.
    """
    # 원문과 정확히 일치하는 단락 통째 교체 (가장 안전)
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    _apply_to_paragraph(paragraph, corrections_dict)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text_frame:
                            for paragraph in cell.text_frame.paragraphs:
                                _apply_to_paragraph(paragraph, corrections_dict)
                                
        # 하단 슬라이드 노트(내레이션) 서식도 동일하게 핫핑크 적용
        if slide.has_notes_slide:
            for paragraph in slide.notes_slide.notes_text_frame.paragraphs:
                _apply_to_paragraph(paragraph, corrections_dict)
                                
def _apply_to_paragraph(paragraph, corrections_dict):
    original_text = paragraph.text.strip()
    if not original_text:
        return
        
    corrected_text = original_text
    is_changed = False
    
    # 딕셔너리 중 일치하는게 있는지 확인
    for old_txt, new_txt in corrections_dict.items():
        if old_txt in corrected_text:
            corrected_text = corrected_text.replace(old_txt, new_txt)
            is_changed = True
            
    # 다중 띄어쓰기를 여기서 기계적으로 추가 반영 (AI가 놓칠 경우 대비)
    spaced_fixed = re.sub(r' {2,}', ' ', corrected_text)
    if spaced_fixed != corrected_text:
        corrected_text = spaced_fixed
        is_changed = True
        
    if not is_changed:
        return
        
    # 기존 서식 복사 준비
    if paragraph.runs:
        font_ref = paragraph.runs[0].font
    else:
        font_ref = None
        
    paragraph.clear()
    
    # 단어(공백 포함) 단위로 토큰화하여 세밀한 부분 비교 수행 (difflib)
    tokens_orig = re.split(r'(\s+)', original_text)
    tokens_corr = re.split(r'(\s+)', corrected_text)
    
    matcher = difflib.SequenceMatcher(None, tokens_orig, tokens_corr)
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'delete':
            continue  # 삭제된 텍스트는 넘어감
            
        chunk_text = "".join(tokens_corr[j1:j2])
        if not chunk_text:
            continue
            
        new_run = paragraph.add_run()
        new_run.text = chunk_text
        
        # 기본 폰트 속성 상속
        if font_ref:
            for attr in ['name', 'size', 'bold', 'italic', 'underline']:
                try: setattr(new_run.font, attr, getattr(font_ref, attr))
                except: pass
                
        # 수정된 부분에만 부분적으로 핫핑크색 적용!
        if tag in ('replace', 'insert'):
            new_run.font.color.rgb = RGBColor(255, 105, 180)
        else:
            # 원본 유지 부분은 기존 색상 복구 시도
            if font_ref and hasattr(font_ref, 'color') and hasattr(font_ref.color, 'rgb') and font_ref.color.rgb:
                try: new_run.font.color.rgb = font_ref.color.rgb
                except: pass
            elif font_ref and hasattr(font_ref, 'color') and hasattr(font_ref.color, 'theme_color') and font_ref.color.theme_color:
                try: new_run.font.color.theme_color = font_ref.color.theme_color
                except: pass
