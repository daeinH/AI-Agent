import os
import sys
import time
import io
import math

# Windows 한글 인코딩 강제 설정
os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["PYTHONUTF8"] = "1"
if sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
if sys.stderr.encoding.lower() != 'utf-8':
    sys.stderr.reconfigure(encoding='utf-8')

import streamlit as st
import pandas as pd
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import google.generativeai as genai

try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
except:
    API_KEY = None

# ==========================================
# 1. KEC/IEC 정밀 설계 엔진 (AI Tool 연동용)
# ==========================================
class UtilityEngineeringEngine:
    def __init__(self):
        # 차단기 단가표
        self.price_db = {
            "MCCB_50AF": 45000, "MCCB_125AF": 120000, "MCCB_250AF": 250000,
            "MCCB_400AF": 450000, "MCCB_630AF": 750000, "MCCB_800AF": 1100000, "MCCB_1000AF": 1500000,
            "Cable_Tray_W300": 25000
        }
        
        # KEC 전선 규격 DB: (SQ, E 노출, A1 단열벽, C 콘크리트, R[Ω/km], X[Ω/km])
        self.cable_data = [
            (2.5, 31, 18.5, 24, 8.91, 0.106), (4, 42, 25, 33, 5.57, 0.101),
            (6, 54, 32, 43, 3.71, 0.096), (10, 75, 44, 58, 2.24, 0.090),
            (16, 100, 59, 79, 1.41, 0.086), (25, 127, 77, 105, 0.889, 0.083),
            (35, 158, 96, 130, 0.641, 0.081), (50, 192, 117, 161, 0.473, 0.080),
            (70, 246, 149, 204, 0.328, 0.077), (95, 298, 180, 246, 0.236, 0.076),
            (120, 346, 208, 285, 0.187, 0.074), (150, 399, 236, 328, 0.153, 0.074),
            (185, 456, 268, 372, 0.122, 0.074), (240, 538, 315, 434, 0.093, 0.073),
            (300, 621, 363, 500, 0.074, 0.073)
        ]

    def advanced_sizing(self, load_kw: float, distance_m: float, power_factor: float = 0.9, demand_factor: float = 1.0, is_continuous: bool = True, power_type: str = "3상4선", install_method: str = "E", temp_c: float = 30.0) -> dict:
        """장비 용량, 거리, 역률, 공사 방법 등을 반영하여 정밀 Sizing 및 전압강하를 산출합니다."""
        # 1. 전원 방식 세팅
        if power_type == "3상4선":
            v_line, v_base, phase_type, b_coeff = 380, 220, 3, math.sqrt(3)
        elif power_type == "단상":
            v_line, v_base, phase_type, b_coeff = 220, 220, 1, 2.0
        else:
            v_line, v_base, phase_type, b_coeff = 380, 380, 3, math.sqrt(3)

        # 2. 수용률 반영 부하 및 정격 전류 산출
        applied_kw = load_kw * demand_factor
        current = (applied_kw * 1000) / (math.sqrt(3) * v_line * power_factor) if phase_type == 3 else (applied_kw * 1000) / (v_line * power_factor)
        
        # 3. 차단기 선정 (연속 부하 여유율 반영)
        margin = 1.25 if is_continuous else 1.0
        target_cb = current * margin
        cb_standard = [30, 50, 100, 200, 250, 400, 600, 800, 1000]
        selected_cb = next((cb for cb in cb_standard if cb >= target_cb), 1000)

        # 4. 환경 계수 및 공사 방법 세팅
        temp_factor = 1.0 if temp_c <= 30 else (0.96 if temp_c <= 35 else (0.91 if temp_c <= 40 else 0.82))
        method_idx = 2 if install_method == "A1" else (3 if install_method == "C" else 1)
        sin_theta = math.sqrt(1 - power_factor**2)

        # 5. KEC 정밀 전압강하 기반 최적 케이블 선정
        optimal_sq, final_v_drop = None, 999.0
        for row in self.cable_data:
            sq, base_amp, r_val, x_val = row[0], row[method_idx], row[4], row[5]
            adjusted_amp = base_amp * temp_factor
            
            if adjusted_amp >= selected_cb:
                e_drop = (b_coeff * current * distance_m * (r_val * power_factor + x_val * sin_theta)) / 1000
                v_drop_p = (e_drop / v_base) * 100
                
                if v_drop_p <= 3.0:
                    optimal_sq = sq
                    final_v_drop = v_drop_p
                    break

        return {
            "load_current_A": round(current, 1),
            "selected_breaker_AT": selected_cb,
            "optimal_cable_SQ": optimal_sq if optimal_sq else "규격 초과 (다조 포설 요망)",
            "voltage_drop_percent": round(final_v_drop, 2)
        }

engine = UtilityEngineeringEngine()

# Gemini Tool 바인딩
def precision_design_tool(load_kw: float, distance_m: float, power_factor: float, demand_factor: float, is_continuous: bool, power_type: str, install_method: str, temp_c: float) -> dict:
    """모든 변수를 입력받아 KEC 정밀 기반 차단기/케이블 규격과 전압강하를 도출합니다."""
    return engine.advanced_sizing(load_kw, distance_m, power_factor, demand_factor, is_continuous, power_type, install_method, temp_c)

tools_list = [precision_design_tool]

# ==========================================
# 2. 문서 자동생성 (엑셀 및 화면 표)
# ==========================================
def generate_excel_document(prompt_text, ai_response_text):
    wb = Workbook()
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4F81BD")
    align_center = Alignment(horizontal="center", vertical="center")

    ws1 = wb.active
    ws1.title = "1. 증설 공사 기안 및 발주서"
    ws1.append(["항목", "내용"])
    for cell in ws1[1]: cell.font, cell.fill, cell.alignment = header_font, header_fill, align_center
    ws1.append(["공사명", "Utility 설비 Hook-up 증설 공사"])
    ws1.append(["요청 요약", prompt_text])
    ws1.append(["AI 정밀 검토", ai_response_text[:3000]])
    ws1.column_dimensions['A'].width = 25; ws1.column_dimensions['B'].width = 100

    ws2 = wb.create_sheet(title="2. 안전작업허가서(PTW)")
    ws2.append(["구분", "안전 확보 지침 (LOTO)"])
    for cell in ws2[1]: cell.font, cell.fill, cell.alignment = header_font, header_fill, align_center
    ws2.append(["위험성 평가", "감전 및 단락 사고 위험"])
    ws2.append(["LOTO 절차", "1. 판넬 차단기 Open\n2. 잠금장치 및 Tag 부착"])
    ws2.column_dimensions['A'].width = 25; ws2.column_dimensions['B'].width = 80

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# 3. Streamlit UI
# ==========================================
st.set_page_config(page_title="유틸리티 AI Agent", page_icon="⚡", layout="wide")
st.title("⚡ Utility 전기 AI Agent (정밀 KEC 설계 탑재)")

with st.sidebar:
    st.header("🛠️ 업무 모드 선택")
    app_mode = st.radio("모드:", ["🏗️ 증설 엔지니어링 (발주/안전)", "👨‍🏫 신입사원 SLD 튜터링"])
    st.markdown("---")
    sld_file = st.file_uploader("단선도(SLD) 업로드", type=["jpg", "png", "pdf"])
    if sld_file and sld_file.type != "application/pdf": st.image(Image.open(sld_file))

if "eng_msg" not in st.session_state: st.session_state.eng_msg = []
if "latest_eng_result" not in st.session_state: st.session_state.latest_eng_result = None
if "latest_eng_prompt" not in st.session_state: st.session_state.latest_eng_prompt = None

if app_mode == "🏗️ 증설 엔지니어링 (발주/안전)":
    st.info("💡 장비 용량, 거리, 온도, 공사 방법(E, A1, C) 등을 입력하면 AI가 KEC 교류 정밀 임피던스를 반영해 최적화합니다.")
    
    for msg in st.session_state.eng_msg:
        with st.chat_message(msg["role"]): st.markdown(msg["content"])
        
    if prompt := st.chat_input("입력 예: 3상4선식 50kW 장비 120m 증설. 수용률 0.8, 연속부하, E트레이, 35도 환경"):
        st.session_state.eng_msg.append({"role": "user", "content": prompt})
        with st.chat_message("user"): st.markdown(prompt)

        if API_KEY:
            with st.chat_message("assistant"):
                with st.spinner("AI가 KEC 정밀 교류 전압강하를 계산 중입니다..."):
                    genai.configure(api_key=API_KEY)
                    sys_instruct = """당신은 KEC 설계 규정을 완벽히 숙지한 전기 엔지니어입니다.
                    사용자 요청에서 다음 변수를 추출해 precision_design_tool을 호출하세요:
                    - load_kw, distance_m, power_factor(기본 0.9), demand_factor(기본 1.0)
                    - is_continuous(기본 True), power_type(3상4선, 단상 등), install_method(E/A1/C), temp_c(기본 30)
                    
                    결과를 바탕으로 아래 목차로 브리핑하세요:
                    1. 적용된 설계 기준 (온도, 공사방법, 수용률 등 명시)
                    2. 정밀 케이블/차단기 Sizing 및 전압강하율
                    3. 안전작업 계획 요약 (PTW)"""
                    
                    model = genai.GenerativeModel(model_name='gemini-3.6-flash', system_instruction=sys_instruct, tools=tools_list)
                    chat = model.start_chat(enable_automatic_function_calling=True)
                    
                    contents = [{"mime_type": "application/pdf", "data": sld_file.getvalue()}] if sld_file and sld_file.type == "application/pdf" else ([Image.open(sld_file)] if sld_file else [])
                    contents.append(prompt)

                    delay = 5
                    for attempt in range(3):
                        try:
                            final_text = chat.send_message(contents).text
                            break
                        except Exception as e:
                            if "503" in str(e) or "429" in str(e):
                                if attempt < 2: time.sleep(delay); delay *= 2
                                else: raise Exception("서버 혼잡.")
                            else: raise e
                    
                    st.markdown(final_text)
                    st.session_state.eng_msg.append({"role": "assistant", "content": final_text})
                    st.session_state.latest_eng_result, st.session_state.latest_eng_prompt = final_text, prompt

    if st.session_state.latest_eng_result:
        st.markdown("---")
        st.subheader("📑 결재용 문서 미리보기 및 다운로드")
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### 📄 1. 증설 기안서 요약")
            st.table(pd.DataFrame([["공사명", "Utility Hook-up 증설"], ["요청 요약", st.session_state.latest_eng_prompt]], columns=["항목", "내용"]))
        with col2:
            st.markdown("#### 📄 2. 안전작업허가서(PTW)")
            st.table(pd.DataFrame([["위험성 평가", "감전, 단락 사고 위험"], ["LOTO 절차", "1. 차단기 Open 2. Tag 부착"]], columns=["구분", "안전 지침"]))

        st.download_button("📥 엑셀 원본 다운로드", data=generate_excel_document(st.session_state.latest_eng_prompt, st.session_state.latest_eng_result), file_name="Hook-Up_정밀설계.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.subheader("👨‍🏫 SLD 도면 스터디 (친절한 사수 AI)")
    # (튜터링 로직은 기존과 완전히 동일하게 유지됩니다. 지면 관계상 모델명만 gemini-3.6-flash 로 맞추어 배포하시면 됩니다.)