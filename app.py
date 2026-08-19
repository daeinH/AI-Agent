import os
import sys
import time
import io

# 🚨 Windows 한글 인코딩 충돌 방지
os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["PYTHONUTF8"] = "1"
if sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
if sys.stderr.encoding.lower() != 'utf-8':
    sys.stderr.reconfigure(encoding='utf-8')

import streamlit as st
import math
import pandas as pd
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# 구글 공식 SDK
import google.generativeai as genai

# ==========================================
# 🔐 API Key 연동 (웹 배포용)
# ==========================================
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
except:
    API_KEY = None

# ==========================================
# 1. Core Engineering Logic (동적 엔진)
# ==========================================
class UtilityEngineeringEngine:
    def __init__(self):
        self.price_db = {
            "MCCB_50AF/30AT": 45000, "MCCB_125AF/100AT": 120000, "MCCB_250AF/200AT": 250000,
            "F-CV_16sq": 5000, "F-CV_35sq": 12000, "F-CV_70sq": 24000, "Cable_Tray_W300": 25000
        }

    def step1_sizing(self, power_kw: float) -> dict:
        voltage_v = 380
        power_factor = 0.9
        current = (power_kw * 1000) / (math.sqrt(3) * voltage_v * power_factor)
        design_current = current * 1.25
        
        if design_current <= 30: breaker, cable, cable_sq = "MCCB_50AF/30AT", "F-CV_16sq", 16
        elif design_current <= 100: breaker, cable, cable_sq = "MCCB_125AF/100AT", "F-CV_35sq", 35
        else: breaker, cable, cable_sq = "MCCB_250AF/200AT", "F-CV_70sq", 70
        return {"rated_current_A": round(current, 2), "breaker": breaker, "cable": cable, "cable_sq": cable_sq}

    def step2_evaluate_capacity(self, tr_capacity_kva: float, current_load_kw: float, add_power_kw: float) -> dict:
        expected_load = current_load_kw + add_power_kw
        expected_rate = (expected_load / tr_capacity_kva) * 100
        is_safe = expected_rate <= 80
        analysis = "적합" if is_safe else f"부하율 {expected_rate:.1f}%로 위험."
        solution = "문제없음" if is_safe else "타 변압기 연계 요망."
        return {"tr_capacity_kva": tr_capacity_kva, "expected_load_rate": round(expected_rate, 2), "is_safe": is_safe, "analysis": analysis, "solution": solution}

    def step3_kec_verification(self, current_A: float, cable_sq: float, length_m: float) -> dict:
        drop_rate = ((30.8 * length_m * current_A) / (1000 * cable_sq) / 380) * 100
        return {"voltage_drop_rate": round(drop_rate, 2), "is_passed": drop_rate <= 3.0}

    def step4_generate_boq(self, breaker_name: str, cable_name: str, length_m: float) -> dict:
        total = self.price_db.get(breaker_name, 0) + (self.price_db.get(cable_name, 0) * length_m) + (self.price_db["Cable_Tray_W300"] * length_m)
        return {"total_estimated_cost": total}

engine = UtilityEngineeringEngine()
tools_list = [engine.step1_sizing, engine.step2_evaluate_capacity, engine.step3_kec_verification, engine.step4_generate_boq]

# ==========================================
# 2. 문서 자동생성 로직 (엑셀)
# ==========================================
def generate_excel_document(prompt_text, ai_response_text):
    """엔지니어링 검토 결과를 바탕으로 기안서 및 안전작업허가서(PTW) 엑셀 생성"""
    wb = Workbook()
    
    # 헤더 스타일
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4F81BD")
    align_center = Alignment(horizontal="center", vertical="center")

    # Sheet 1: 발주 기안서
    ws1 = wb.active
    ws1.title = "1. 증설 공사 기안 및 발주서"
    headers1 = ["항목", "내용"]
    ws1.append(headers1)
    for cell in ws1[1]:
        cell.font, cell.fill, cell.alignment = header_font, header_fill, align_center
        
    ws1.append(["공사명", "Utility 설비 Hook-up 증설 공사"])
    ws1.append(["요청 부서", "생산기술팀"])
    ws1.append(["요청 요약", prompt_text])
    ws1.append(["AI 검토 내용 (BoQ 및 규격)", ai_response_text[:3000]]) # 엑셀 셀 제한 방지
    ws1.column_dimensions['A'].width = 25
    ws1.column_dimensions['B'].width = 100

    # Sheet 2: 안전작업허가서 (PTW) & LOTO
    ws2 = wb.create_sheet(title="2. 안전작업허가서(PTW)")
    headers2 = ["구분", "안전 확보 지침 (LOTO)"]
    ws2.append(headers2)
    for cell in ws2[1]:
        cell.font, cell.fill, cell.alignment = header_font, header_fill, align_center
        
    ws2.append(["작업 위험성 평가", "감전, 아크 플래시, 단락 사고 위험"])
    ws2.append(["LOTO (차단 절차)", "1. 메인 판넬 차단기(MCCB) Open\n2. 잠금장치(Lock) 체결 및 위험 Tag 부착"])
    ws2.append(["안전 점검", "1. 검전기로 무전압 확인\n2. 잔류 전하 방전 및 접지 용구 설치"])
    ws2.column_dimensions['A'].width = 25
    ws2.column_dimensions['B'].width = 80

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# 3. Streamlit UI
# ==========================================
st.set_page_config(page_title="신입 엔지니어 필수 AI Agent", page_icon="⚡", layout="wide")

st.title("⚡ Utility 전기 엔지니어링 AI Agent")

# 사이드바: 업무 모드 및 데이터 업로드
with st.sidebar:
    st.header("🛠️ 업무 모드 선택")
    app_mode = st.radio("수행할 업무를 선택하세요:", 
                        ["🏗️ 증설 엔지니어링 (발주/안전)", "👨‍🏫 신입사원 SLD 튜터링"])
    st.markdown("---")
    
    st.header("📂 현장 데이터 업로드")
    sld_file = st.file_uploader("단선도(SLD) 파일 업로드", type=["jpg", "png", "pdf"])
    if sld_file and sld_file.type != "application/pdf":
        st.image(Image.open(sld_file), caption="SLD 도면", use_container_width=True)

    scada_file = st.file_uploader("SCADA 부하 데이터 업로드", type=["csv", "xlsx"])
    scada_context = ""
    if scada_file:
        df = pd.read_csv(scada_file) if scada_file.name.endswith('.csv') else pd.read_excel(scada_file)
        scada_context = f"\n\n### [SCADA 데이터]\n{df.to_string(index=False)}"
        st.success("✅ SCADA 연동됨")

# 세션 상태 분리 (모드별 채팅 기록)
if "eng_msg" not in st.session_state: st.session_state.eng_msg = []
if "tutor_msg" not in st.session_state: st.session_state.tutor_msg = []
if "latest_eng_result" not in st.session_state: st.session_state.latest_eng_result = None
if "latest_eng_prompt" not in st.session_state: st.session_state.latest_eng_prompt = None

# 모드에 따른 화면 구성
if app_mode == "🏗️ 증설 엔지니어링 (발주/안전)":
    st.subheader("📊 부하 증설 검토 및 공사 발주 자동화")
    st.info("증설 용량과 거리를 입력하면 Sizing, KEC 검증 후 **품의서 및 안전허가서(PTW)를 엑셀로 자동 생성**해 줍니다.")
    
    for msg in st.session_state.eng_msg:
        with st.chat_message(msg["role"]): st.markdown(msg["content"])
        
    if prompt := st.chat_input("증설 계획을 입력하세요 (예: TR-1에 50kW 장비 80m 증설)"):
        st.session_state.eng_msg.append({"role": "user", "content": prompt})
        with st.chat_message("user"): st.markdown(prompt)

        if API_KEY:
            with st.chat_message("assistant"):
                with st.spinner("엔지니어링 검토 및 문서 초안 작성 중..."):
                    genai.configure(api_key=API_KEY)
                    sys_instruct = """당신은 Hook-Up 공사를 총괄하는 전기 엔지니어입니다.
                    Tool을 사용해 도면과 SCADA를 분석하고 다음 목차로 보고서를 쓰세요:
                    1. 분석 데이터
                    2. 케이블/차단기 Sizing
                    3. 여유용량 및 KEC 검증
                    4. 🚨문제 진단 및 경제적 대안(VE)
                    5. 안전작업 계획 (PTW 및 LOTO 절차 요약) <- 필수 추가
                    6. 발주 예상 공사비"""
                    
                    model = genai.GenerativeModel(model_name='gemini-1.5-flash-latest', system_instruction=sys_instruct, tools=tools_list)
                    chat = model.start_chat(enable_automatic_function_calling=True)
                    
                    contents = []
                    if sld_file:
                        if sld_file.type == "application/pdf": contents.append({"mime_type": "application/pdf", "data": sld_file.getvalue()})
                        else: contents.append(Image.open(sld_file))
                    contents.append(prompt + scada_context)

                    # Backoff 재시도 로직
                    retry_delay = 5
                    for attempt in range(3):
                        try:
                            res = chat.send_message(contents)
                            final_text = res.text
                            break
                        except Exception as e:
                            if "503" in str(e) or "429" in str(e):
                                time.sleep(retry_delay); retry_delay *= 2
                            else: raise e
                    
                    st.markdown(final_text)
                    st.session_state.eng_msg.append({"role": "assistant", "content": final_text})
                    st.session_state.latest_eng_result = final_text
                    st.session_state.latest_eng_prompt = prompt

    # 검토가 완료되면 엑셀 다운로드 버튼 노출
    if st.session_state.latest_eng_result:
        st.markdown("---")
        st.subheader("📑 결재용 자동 생성 문서")
        excel_data = generate_excel_document(st.session_state.latest_eng_prompt, st.session_state.latest_eng_result)
        st.download_button(
            label="📥 공사 발주 품의서 및 안전작업허가서(PTW) 엑셀 다운로드",
            data=excel_data,
            file_name="Hook-Up_기안_및_안전계획.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    # 모드 2: 신입사원 튜터링
    st.subheader("👨‍🏫 SLD 도면 스터디 (친절한 사수 AI)")
    st.info("신입사원 교육 모드입니다. 업로드된 단선도(SLD)에 대해 궁금한 점을 무엇이든 물어보세요!")
    
    for msg in st.session_state.tutor_msg:
        with st.chat_message(msg["role"]): st.markdown(msg["content"])
        
    if prompt := st.chat_input("도면에 대해 질문하세요 (예: VCB 차단기의 역할이 뭐야?)"):
        st.session_state.tutor_msg.append({"role": "user", "content": prompt})
        with st.chat_message("user"): st.markdown(prompt)

        if API_KEY:
            with st.chat_message("assistant"):
                with st.spinner("사수 AI가 도면을 확인하고 있습니다..."):
                    genai.configure(api_key=API_KEY)
                    sys_instruct = """당신은 갓 입사한 전기 직무 신입사원에게 도면(SLD)을 친절하게 가르쳐주는 10년 차 사수입니다. 
                    계산 Tool은 사용하지 마세요. 도면 이미지를 보고 설비의 역할, 전기의 흐름, 현장 용어를 아주 알기 쉽게 설명해주세요."""
                    
                    model = genai.GenerativeModel(model_name='gemini-1.5-flash-latest', system_instruction=sys_instruct)
                    
                    contents = []
                    if sld_file:
                        if sld_file.type == "application/pdf": contents.append({"mime_type": "application/pdf", "data": sld_file.getvalue()})
                        else: contents.append(Image.open(sld_file))
                    contents.append(prompt)

                    retry_delay = 5
                    for attempt in range(3):
                        try:
                            res = model.generate_content(contents)
                            final_text = res.text
                            break
                        except Exception as e:
                            if "503" in str(e) or "429" in str(e):
                                time.sleep(retry_delay); retry_delay *= 2
                            else: raise e
                    
                    st.markdown(final_text)
                    st.session_state.tutor_msg.append({"role": "assistant", "content": final_text})