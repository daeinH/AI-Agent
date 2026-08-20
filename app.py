import os
import sys
import time
import io
import math

# 🚨 Windows 한글 인코딩 충돌 방지
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

# ==========================================
# 🔐 API Key 연동
# ==========================================
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
except:
    API_KEY = None

# ==========================================
# 1. Core Engineering Logic (SCADA + KEC 정밀 + 노무비 엔진)
# ==========================================
class UtilityEngineeringEngine:
    def __init__(self):
        self.price_db = {
            "MCCB_50AF": 45000, "MCCB_125AF": 120000, "MCCB_250AF": 250000,
            "MCCB_400AF": 450000, "MCCB_630AF": 750000, "MCCB_800AF": 1100000, "MCCB_1000AF": 1500000,
            "Cable_Tray_W300": 25000
        }
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

    def step1_advanced_sizing(self, load_kw: float, distance_m: float, power_factor: float = 0.9, demand_factor: float = 1.0, is_continuous: bool = True, power_type: str = "3상4선", install_method: str = "E", temp_c: float = 30.0) -> dict:
        if "3상4선" in power_type: v_line, v_base, phase_type, b_coeff = 380, 220, 3, math.sqrt(3)
        elif "단상" in power_type: v_line, v_base, phase_type, b_coeff = 220, 220, 1, 2.0
        else: v_line, v_base, phase_type, b_coeff = 380, 380, 3, math.sqrt(3)

        applied_kw = load_kw * demand_factor
        current = (applied_kw * 1000) / (math.sqrt(3) * v_line * power_factor) if phase_type == 3 else (applied_kw * 1000) / (v_line * power_factor)
        
        margin = 1.25 if is_continuous else 1.0
        target_cb = current * margin
        cb_standard = [30, 50, 100, 125, 200, 250, 400, 630, 800, 1000]
        selected_cb = next((cb for cb in cb_standard if cb >= target_cb), 1000)

        temp_factor = 1.0 if temp_c <= 30 else (0.96 if temp_c <= 35 else (0.91 if temp_c <= 40 else 0.82))
        method_idx = 2 if "A1" in install_method else (3 if "C" in install_method else 1)
        sin_theta = math.sqrt(1 - power_factor**2)

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
            "design_margin_applied": margin,
            "selected_breaker_AT": selected_cb,
            "optimal_cable_SQ": optimal_sq if optimal_sq else "규격 초과 (다조 포설 필요)",
            "voltage_drop_percent": round(final_v_drop, 2)
        }

    def step2_evaluate_capacity(self, tr_capacity_kva: float, current_load_kw: float, add_power_kw: float) -> dict:
        expected_load = current_load_kw + add_power_kw
        expected_rate = (expected_load / tr_capacity_kva) * 100
        is_safe = expected_rate <= 80
        analysis = "적합" if is_safe else f"부하율 {expected_rate:.1f}%로 위험."
        return {"tr_capacity_kva": tr_capacity_kva, "expected_load_rate": round(expected_rate, 2), "is_safe": is_safe, "analysis": analysis}

    def step3_generate_boq(self, breaker_name: str, cable_sq: float, length_m: float, labor_unit_cost: int = 280000) -> dict:
        """자재비와 노무비(품셈 기반)를 분리하여 공사비를 산출합니다."""
        breaker_cost = self.price_db.get(breaker_name, 120000)
        
        # 문자열(예: '규격 초과')이 들어올 경우 임의의 최대 SQ(300)로 계산하여 오류 방지
        sq_val = cable_sq if isinstance(cable_sq, (int, float)) else 300
        cable_cost = (sq_val * 400) * length_m
        tray_cost = self.price_db["Cable_Tray_W300"] * length_m
        material_cost = breaker_cost + cable_cost + tray_cost

        # 노무량 산출 (가상 표준품셈: 10m당 기본 1품 + 케이블 굵기에 따른 가중치)
        man_days = (length_m / 10) * (1 + (sq_val / 100))
        labor_cost = int(man_days * labor_unit_cost)

        return {
            "material_cost": material_cost,
            "labor_cost": labor_cost,
            "total_estimated_cost": material_cost + labor_cost,
            "applied_labor_unit_cost": labor_unit_cost,
            "estimated_man_days": round(man_days, 1)
        }

engine = UtilityEngineeringEngine()

def precision_design_tool(load_kw: float, distance_m: float, power_factor: float = 0.9, demand_factor: float = 1.0, is_continuous: bool = True, power_type: str = "3상4선", install_method: str = "E", temp_c: float = 30.0) -> dict:
    return engine.step1_advanced_sizing(load_kw, distance_m, power_factor, demand_factor, is_continuous, power_type, install_method, temp_c)

def evaluate_capacity_tool(tr_capacity_kva: float, current_load_kw: float, add_power_kw: float) -> dict:
    return engine.step2_evaluate_capacity(tr_capacity_kva, current_load_kw, add_power_kw)

def generate_boq_tool(breaker_name: str, cable_sq: float, length_m: float, labor_unit_cost: int = 280000) -> dict:
    """자재비와 노무비를 산출합니다. labor_unit_cost는 2026년 표준 내선전공 단가(280,000원)가 기본값입니다."""
    return engine.step3_generate_boq(breaker_name, cable_sq, length_m, labor_unit_cost)

tools_list = [precision_design_tool, evaluate_capacity_tool, generate_boq_tool]

# ==========================================
# 2. 문서 자동생성 로직
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
    ws1.append(["요청 요약", prompt_text])
    ws1.append(["AI 최종 솔루션 요약", ai_response_text[:3000]])
    ws1.column_dimensions['A'].width = 25
    ws1.column_dimensions['B'].width = 100

    ws2 = wb.create_sheet(title="2. 안전작업허가서(PTW)")
    ws2.append(["구분", "안전 확보 지침 (LOTO)"])
    for cell in ws2[1]: cell.font, cell.fill, cell.alignment = header_font, header_fill, align_center
    ws2.append(["작업 위험성 평가", "감전, 아크 플래시, 단락 사고 위험"])
    ws2.append(["LOTO (차단 절차)", "1. 메인 판넬 차단기(MCCB) Open\n2. 잠금장치(Lock) 체결 및 위험 Tag 부착"])
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
st.title("⚡ Utility 전기 AI Agent (SCADA 연동 + KEC 정밀 설계)")

with st.sidebar:
    st.header("🛠️ 업무 모드 선택")
    app_mode = st.radio("수행할 업무를 선택하세요:", ["🏗️ 증설 엔지니어링 (발주/안전)", "👨‍🏫 신입사원 SLD 튜터링"])
    st.markdown("---")
    
    st.header("📂 현장 데이터 업로드")
    sld_file = st.file_uploader("1. 단선도(SLD) 도면", type=["jpg", "png", "pdf"])
    if sld_file and sld_file.type != "application/pdf": st.image(Image.open(sld_file), use_container_width=True)

    scada_file = st.file_uploader("2. SCADA 부하 데이터", type=["csv", "xlsx"])
    scada_context = ""
    if scada_file:
        df = pd.read_csv(scada_file) if scada_file.name.endswith('.csv') else pd.read_excel(scada_file)
        scada_context = f"\n\n### [업로드된 SCADA 실시간 데이터 참고]\n{df.to_string(index=False)}"
        st.success("✅ SCADA 연동됨")

    labor_file = st.file_uploader("3. 노무비 단가표 (선택)", type=["csv", "xlsx"])
    labor_context = "\n\n### [노무비 단가 정보]\n업로드된 파일 없음. 2026년 표준 내선전공 단가(280,000원/인) 적용 요망."
    if labor_file:
        df_labor = pd.read_csv(labor_file) if labor_file.name.endswith('.csv') else pd.read_excel(labor_file)
        labor_context = f"\n\n### [업로드된 노무비 단가표 참고]\n{df_labor.to_string(index=False)}"
        st.success("✅ 노무비 DB 연동됨")

if "eng_msg" not in st.session_state: st.session_state.eng_msg = []
if "tutor_msg" not in st.session_state: st.session_state.tutor_msg = []
if "latest_eng_result" not in st.session_state: st.session_state.latest_eng_result = None
if "latest_eng_prompt" not in st.session_state: st.session_state.latest_eng_prompt = None

if app_mode == "🏗️ 증설 엔지니어링 (발주/안전)":
    st.subheader("📊 부하 증설 검토 및 공사 발주 자동화")
    st.info("장비 용량, 거리, 환경 등을 입력하면 AI가 SCADA와 노무비 DB를 조회하여 **문제별 맞춤 대안과 최종 통합 솔루션**을 제시합니다.")
    
    for msg in st.session_state.eng_msg:
        with st.chat_message(msg["role"]): st.markdown(msg["content"])
        
    if prompt := st.chat_input("예: SCADA상 TR-1에 500kW 200m 증설. C공사 40도 수용률 0.8"):
        st.session_state.eng_msg.append({"role": "user", "content": prompt})
        with st.chat_message("user"): st.markdown(prompt)

        if API_KEY:
            with st.chat_message("assistant"):
                with st.spinner("SCADA 분석, 정밀 규격 산출 및 노무비 계산 중..."):
                    genai.configure(api_key=API_KEY)
                    sys_instruct = """당신은 Hook-Up 공사를 총괄하는 최고 수준의 전기 엔지니어입니다.
                    반드시 다음 순서로 검토를 진행하세요:
                    1. evaluate_capacity_tool 호출 (SCADA 데이터 참조, 없으면 정격의 70% 가정)
                    2. precision_design_tool 호출
                    3. generate_boq_tool 호출 ([노무비 단가 정보]에서 추출한 단가 적용, 없으면 기본값 280,000원 적용)
                    
                    최종 보고서는 반드시 다음 목차와 형식을 엄격히 지켜 작성하세요:
                    
                    ### 1. SCADA 변압기 부하 분석
                    - 내용 작성
                    
                    ### 2. KEC 정밀 케이블/차단기 Sizing
                    - 내용 작성
                    
                    ### 3. 🚨 주요 문제점 및 맞춤형 대안 (VE)
                    - [문제 1: 변압기 용량 초과] (문제 설명)
                      👉 [대안 1: TR-2 연계 및 부하 분산] (대안 설명)
                    - [문제 2: 장거리 전압강하 및 케이블 규격 초과] (문제 설명)
                      👉 [대안 2: 다조 포설 또는 수전반 위치 재배치] (대안 설명)
                    (문제가 없다면 '특이사항 없음'으로 기재)
                    
                    ### 4. 발주 예상 공사비 (자재비/노무비 분리)
                    - 내용 작성 (적용된 노무비 단가 명시)
                    
                    ### 5. 🎯 최종 합리적 증설 솔루션 요약
                    - 위의 모든 상황(계통, 케이블, 비용)을 종합 고려하여 현장 작업자에게 지시하듯 가장 합리적인 '최종 추천 스펙(케이블 SQ, 차단기, 연계 변압기)'을 한 문단으로 명확하게 결론지어 주세요."""
                    
                    model = genai.GenerativeModel(model_name='gemini-3.6-flash', system_instruction=sys_instruct, tools=tools_list)
                    chat = model.start_chat(enable_automatic_function_calling=True)
                    
                    contents = []
                    if sld_file:
                        if sld_file.type == "application/pdf": contents.append({"mime_type": "application/pdf", "data": sld_file.getvalue()})
                        else: contents.append(Image.open(sld_file))
                    contents.append(prompt + scada_context + labor_context)

                    retry_delay = 5
                    for attempt in range(3):
                        try:
                            res = chat.send_message(contents)
                            final_text = res.text
                            break
                        except Exception as e:
                            if "503" in str(e) or "429" in str(e):
                                if attempt < 2: time.sleep(retry_delay); retry_delay *= 2
                                else: raise Exception("현재 구글 AI 서버가 혼잡합니다. 잠시 후 다시 시도해주세요.")
                            else: raise e
                    
                    st.markdown(final_text)
                    st.session_state.eng_msg.append({"role": "assistant", "content": final_text})
                    st.session_state.latest_eng_result = final_text
                    st.session_state.latest_eng_prompt = prompt

    if st.session_state.latest_eng_result:
        st.markdown("---")
        st.subheader("📑 결재용 문서 미리보기 및 다운로드")
        
        df_draft = pd.DataFrame([
            ["공사명", "Utility 설비 Hook-up 증설 공사"],
            ["요청 요약", st.session_state.latest_eng_prompt],
            ["최종 솔루션", "위의 [5. 최종 합리적 증설 솔루션 요약] 참조"]
        ], columns=["항목", "내용"])
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### 📄 1. 증설 기안서 요약")
            st.table(df_draft)
        with col2:
            st.markdown("#### 📄 2. 안전작업허가서(PTW)")
            st.table(pd.DataFrame([["위험성 평가", "감전, 단락 사고 위험"], ["LOTO 절차", "1. 차단기 Open 2. 잠금 및 Tag 부착"]], columns=["구분", "지침"]))

        excel_data = generate_excel_document(st.session_state.latest_eng_prompt, st.session_state.latest_eng_result)
        st.download_button("📥 통합 보고서 다운로드 (엑셀)", data=excel_data, file_name="HookUp_최종결정서.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.subheader("👨‍🏫 SLD 도면 스터디 (친절한 사수 AI)")
    # 튜터링 파트는 기존과 동일
    for msg in st.session_state.tutor_msg:
        with st.chat_message(msg["role"]): st.markdown(msg["content"])
        
    if prompt := st.chat_input("도면에 대해 질문하세요"):
        st.session_state.tutor_msg.append({"role": "user", "content": prompt})
        with st.chat_message("user"): st.markdown(prompt)

        if API_KEY:
            with st.chat_message("assistant"):
                with st.spinner("사수 AI가 도면을 확인 중..."):
                    genai.configure(api_key=API_KEY)
                    sys_instruct = "당신은 신입사원에게 도면(SLD)을 가르쳐주는 10년 차 사수입니다."
                    model = genai.GenerativeModel(model_name='gemini-3.6-flash', system_instruction=sys_instruct)
                    
                    contents = []
                    if sld_file:
                        if sld_file.type == "application/pdf": contents.append({"mime_type": "application/pdf", "data": sld_file.getvalue()})
                        else: contents.append(Image.open(sld_file))
                    contents.append(prompt)

                    retry_delay = 5
                    for attempt in range(3):
                        try:
                            final_text = model.generate_content(contents).text
                            break
                        except Exception as e:
                            if "503" in str(e) or "429" in str(e):
                                if attempt < 2: time.sleep(retry_delay); retry_delay *= 2
                                else: raise Exception("서버 혼잡.")
                            else: raise e
                    
                    st.markdown(final_text)
                    st.session_state.tutor_msg.append({"role": "assistant", "content": final_text})