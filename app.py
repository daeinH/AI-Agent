import os
import sys
import time

# ==========================================
# 🚨 Windows 한글 인코딩(ASCII) 충돌 강제 방지 설정
# ==========================================
os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["PYTHONUTF8"] = "1"
if sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
if sys.stderr.encoding.lower() != 'utf-8':
    sys.stderr.reconfigure(encoding='utf-8')

import streamlit as st
import math
import pandas as pd
from typing import Dict
from PIL import Image

# 🚨 Streamlit Cloud 충돌 방지를 위해 안정화된 구글 공식 SDK 사용
import google.generativeai as genai

# ==========================================
# 🔐 보안된 API Key 연동 (웹 배포용)
# ==========================================
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
except:
    API_KEY = None

# ==========================================
# 1. Core Engineering Logic (동적 엔지니어링 및 진단 엔진)
# ==========================================
class UtilityEngineeringEngine:
    def __init__(self):
        self.price_db = {
            "MCCB_50AF/30AT": 45000, "MCCB_125AF/100AT": 120000, "MCCB_250AF/200AT": 250000,
            "F-CV_16sq": 5000, "F-CV_35sq": 12000, "F-CV_70sq": 24000, "Cable_Tray_W300": 25000
        }

    def step1_sizing(self, power_kw: float) -> Dict:
        voltage_v = 380
        power_factor = 0.9
        current = (power_kw * 1000) / (math.sqrt(3) * voltage_v * power_factor)
        design_current = current * 1.25
        
        if design_current <= 30: breaker, cable, cable_sq = "MCCB_50AF/30AT", "F-CV_16sq", 16
        elif design_current <= 100: breaker, cable, cable_sq = "MCCB_125AF/100AT", "F-CV_35sq", 35
        else: breaker, cable, cable_sq = "MCCB_250AF/200AT", "F-CV_70sq", 70
            
        return {"rated_current_A": round(current, 2), "breaker": breaker, "cable": cable, "cable_sq": cable_sq}

    def step2_evaluate_capacity(self, tr_capacity_kva: float, current_load_kw: float, add_power_kw: float) -> Dict:
        expected_load = current_load_kw + add_power_kw
        expected_rate = (expected_load / tr_capacity_kva) * 100
        is_safe = expected_rate <= 80
        
        if is_safe:
            analysis = "적합 (안정권)"
            solution = "현재 변압기 계통 연계에 문제없음."
        else:
            analysis = f"예상 부하율이 {expected_rate:.1f}%로 안전 기준치(80%)를 초과하여 변압기 소손 및 블랙아웃 위험이 있습니다."
            solution = "경제적 최선책 제안: 막대한 비용이 드는 변압기 신설 대신, 현재 부하율이 낮은 인근 타 변압기(TR)로 전원 인입 지점을 변경하는 것이 가장 경제적입니다."

        return {
            "tr_capacity_kva": tr_capacity_kva,
            "current_load_kw": current_load_kw,
            "expected_load_rate": round(expected_rate, 2),
            "is_safe": is_safe,
            "problem_analysis": analysis,
            "economical_solution": solution
        }

    def step3_kec_verification(self, current_A: float, cable_sq: float, length_m: float) -> Dict:
        voltage_drop = (30.8 * length_m * current_A) / (1000 * cable_sq)
        drop_rate = (voltage_drop / 380) * 100
        is_passed = drop_rate <= 3.0
        
        if is_passed:
            analysis = "적합"
            solution = "KEC 전압강하 규정(3% 이내) 만족."
        else:
            required_sq = (30.8 * length_m * current_A) / (1000 * (380 * 0.03))
            analysis = f"계산된 전압강하율이 {drop_rate:.1f}%로 KEC 허용 기준(3.0%)을 초과하여 장비 오작동 위험이 있습니다."
            solution = f"경제적 최선책 제안: 케이블 포설 거리를 단축할 수 없다면, 케이블 굵기를 최소 {math.ceil(required_sq)}sq 이상의 상위 규격으로 업그레이드하여 재시공 비용을 방지해야 합니다."

        return {
            "voltage_drop_rate": round(drop_rate, 2),
            "is_passed": is_passed,
            "problem_analysis": analysis,
            "economical_solution": solution
        }

    def step4_generate_boq(self, breaker_name: str, cable_name: str, length_m: float) -> Dict:
        breaker_cost = self.price_db.get(breaker_name, 0)
        cable_cost = self.price_db.get(cable_name, 0) * length_m
        tray_cost = self.price_db["Cable_Tray_W300"] * length_m
        return {"total_estimated_cost": breaker_cost + cable_cost + tray_cost}

engine = UtilityEngineeringEngine()

# Gemini 파이썬 Tool 바인딩
def calculate_sizing_tool(power_kw: float) -> dict:
    """1. 장비의 증설 용량(power_kw)을 입력받아 적합한 차단기와 케이블 규격을 계산합니다."""
    return engine.step1_sizing(power_kw)

def evaluate_capacity_tool(tr_capacity_kva: float, current_load_kw: float, add_power_kw: float) -> dict:
    """2. 변압기 용량, 현재 부하, 증설 용량을 입력받아 여유 용량을 평가합니다."""
    return engine.step2_evaluate_capacity(tr_capacity_kva, current_load_kw, add_power_kw)

def verify_kec_tool(current_A: float, cable_sq: float, length_m: float) -> dict:
    """3. 전류, 케이블굵기, 거리를 입력받아 KEC 전압강하 규정을 검증합니다."""
    return engine.step3_kec_verification(current_A, cable_sq, length_m)

def generate_boq_tool(breaker_name: str, cable_name: str, length_m: float) -> dict:
    """4. 차단기, 케이블 명칭, 거리를 입력받아 최종 공사비를 산출합니다."""
    return engine.step4_generate_boq(breaker_name, cable_name, length_m)

tools_list = [calculate_sizing_tool, evaluate_capacity_tool, verify_kec_tool, generate_boq_tool]

# ==========================================
# 2. Streamlit UI (웹 대시보드)
# ==========================================
st.set_page_config(page_title="고급 엔지니어링 AI Agent", page_icon="⚡", layout="wide")

st.title("⚡ 통합 증설 엔지니어링 AI Agent")
st.markdown("도면(SLD)과 실시간 SCADA 데이터를 분석하며, **설계 부적합 시 경제적인 최적 대안을 자동 제안**합니다.")

# 사이드바 설정
with st.sidebar:
    st.header("⚙️ 서비스 상태")
    if API_KEY:
        st.success("✅ AI 서버 정상 연결됨")
    else:
        st.error("❌ 서버 설정 오류: 웹 호스팅 환경에 API Key가 등록되지 않았습니다.")
    
    st.markdown("---")
    st.header("📂 현장 데이터 업로드")
    
    st.subheader("1. 단선도(SLD) 도면")
    sld_file = st.file_uploader("단선도 파일을 업로드하세요.", type=["jpg", "jpeg", "png", "pdf"])
    
    if sld_file:
        if sld_file.type == "application/pdf":
            st.success("✅ PDF 도면 업로드 완료")
        else:
            sld_image = Image.open(sld_file)
            st.image(sld_image, caption="업로드된 SLD 도면", use_container_width=True)

    st.markdown("---")
    
    st.subheader("2. SCADA 실시간 데이터 (선택)")
    scada_file = st.file_uploader("SCADA 운영 데이터(.csv, .xlsx)를 업로드하세요.", type=["csv", "xlsx"])
    
    scada_context = ""
    if scada_file:
        try:
            if scada_file.name.endswith('.csv'):
                df = pd.read_csv(scada_file, encoding='utf-8-sig')
            else:
                df = pd.read_excel(scada_file)
            
            scada_context = f"\n\n### [업로드된 SCADA 실시간 데이터 참고]\n{df.to_string(index=False)}"
            st.success(f"✅ SCADA 데이터 연동 완료 (총 {len(df)}행)")
            with st.expander("데이터 미리보기"):
                st.dataframe(df)
        except Exception as e:
            st.error(f"SCADA 데이터를 읽는 중 오류가 발생했습니다: {e}")

if "messages" not in st.session_state:
    st.session_state.messages = [{"role": "assistant", "content": "안녕하세요! 도면 및 SCADA 데이터(선택)를 업로드하신 후 증설 계획을 말씀해주세요. 만약 설계 기준에 미달하는 경우, 원인 분석과 경제적 대안까지 함께 검토해 드립니다."}]

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

if prompt := st.chat_input("증설 계획 및 요청사항을 자유롭게 입력하세요."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    if not API_KEY:
        st.warning("⚠️ 웹앱 서버에 API 키 설정이 필요합니다. 관리자에게 문의하세요.")
    else:
        with st.chat_message("assistant"):
            with st.spinner("AI Agent가 현장 데이터 분석 및 경제적 타당성(VE)을 검토 중입니다..."):
                try:
                    # API Key 세팅
                    genai.configure(api_key=API_KEY)
                    
                    sys_instruction = """당신은 시각적 도면(SLD)과 실시간 SCADA 데이터를 융합 분석하는 최고 수준의 전기 엔지니어링 AI Agent입니다.
                    다음 절차에 따라 완벽한 엔지니어링 보고서를 작성하세요:
                    1. 업로드된 SLD 도면에서 타겟 변압기(TR)의 정격 용량(kVA)을 추출하세요.
                    2. 사용자가 제시한 프롬프트에 [업로드된 SCADA 실시간 데이터 참고] 표가 있다면, 도면의 TR 명칭과 매칭하여 해당 TR의 '현재 부하량'을 추출하여 계산에 활용하세요. 만약 데이터가 없으면 정격 용량의 70%를 현재 부하량으로 가정하세요.
                    3. calculate_sizing_tool -> evaluate_capacity_tool -> verify_kec_tool -> generate_boq_tool 순서로 도구를 호출하세요.
                    4. 최종 보고서는 다음 목차로 구성하세요:
                       - [1. 분석 기준 데이터 (도면 및 SCADA)]
                       - [2. 부하 Sizing 결과]
                       - [3. 계통 여유용량 및 KEC 검증]
                       - [🚨 문제 진단 및 경제적 최선책 (VE 제안)] <- Tool의 결과에 문제가 발견된 경우 반드시 포함. 문제가 없다면 '특이사항 없음' 기재.
                       - [4. 예상 공사비 (BoQ)]"""

                    # 모델 설정 (서버가 쾌적하고 도면 판독률이 높은 gemini-1.5-flash 모델 적용)
                    model = genai.GenerativeModel(
                        model_name='gemini-1.5-flash',
                        system_instruction=sys_instruction,
                        tools=tools_list
                    )
                    
                    # Tool 자동 호출을 위한 Chat 세션 열기
                    chat = model.start_chat(enable_automatic_function_calling=True)

                    # 프롬프트와 데이터 합치기
                    prompt_with_context = prompt + scada_context
                    contents_to_send = []
                    
                    if sld_file:
                        if sld_file.type == "application/pdf":
                            # PDF 파일 처리
                            contents_to_send.append({"mime_type": "application/pdf", "data": sld_file.getvalue()})
                        else:
                            # 이미지 파일 처리
                            sld_image = Image.open(sld_file)
                            contents_to_send.append(sld_image)
                            
                    contents_to_send.append(prompt_with_context)

                    # 서버 과부하 대응 (Exponential Backoff) 자동 재시도 로직
                    max_retries = 3
                    retry_delay = 5
                    for attempt in range(max_retries):
                        try:
                            # AI에게 데이터 전송 및 답변 받기
                            response = chat.send_message(contents_to_send)
                            final_text = response.text
                            break # 성공하면 반복문 탈출
                            
                        except Exception as api_e:
                            error_msg = str(api_e)
                            if "503" in error_msg or "429" in error_msg:
                                if attempt < max_retries - 1:
                                    st.warning(f"⚠️ 구글 서버 일시적 혼잡. {retry_delay}초 후 자동으로 재시도합니다...")
                                    time.sleep(retry_delay)
                                    retry_delay *= 2
                                else:
                                    raise Exception("현재 구글 AI 서버가 너무 혼잡합니다. 잠시 후 화면을 새로고침하고 다시 시도해주세요.")
                            else:
                                raise api_e

                    # 최종 결과 출력
                    st.markdown(final_text)
                    st.session_state.messages.append({"role": "assistant", "content": final_text})

                except Exception as e:
                    st.error(f"에러가 발생했습니다: {e}")