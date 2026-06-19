import streamlit as st
import google.generativeai as genai
from PIL import Image
import json
import os
import time
import pandas as pd

# 1. 세션 상태 및 외부 클라우드 비밀 금고 자동 로드 로직 추가
if "api_key" not in st.session_state:
    try:
        st.session_state.api_key = st.secrets["GEMINI_API_KEY"]
    except KeyError:
        st.session_state.api_key = ""

if "report_data" not in st.session_state: st.session_state.report_data = None

st.set_page_config(page_title="기계설비 성능점검 검토자문 시스템", page_icon="📝", layout="wide")

with st.sidebar:
    st.header("⚙️ 설정 및 파일 업로드")
    
    # 2. 키가 이미 로드된 경우 성공 메시지를 보여주고, 없는 경우에만 입력창을 띄웁니다.
    if st.session_state.api_key:
        st.success("🔐 API 키가 클라우드에서 안전하게 자동 연결되었습니다.")
        genai.configure(api_key=st.session_state.api_key)
        api_key_to_use = st.session_state.api_key
    else:
        api_key_to_use = st.text_input("Google Gemini API Key", type="password")
        if api_key_to_use:
            st.session_state.api_key = api_key_to_use
            genai.configure(api_key=api_key_to_use)
            
    selected_model = None
    if api_key_to_use:
        try:
            available_models = [m.name.replace("models/", "") for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            if available_models:
                default_idx = 0
                for i, name in enumerate(available_models):
                    if "pro" in name: default_idx = i; break
                selected_model = st.selectbox("🤖 사용할 AI 모델", available_models, index=default_idx)
            else: st.error("⚠️ 사용 가능한 모델이 없습니다.")
        except Exception: st.error("⚠️ API 키가 유효하지 않습니다.")
            
    st.markdown("---")
    uploaded_files = st.file_uploader("보고서 업로드 (사진 여러 장 또는 PDF)", type=["jpg", "jpeg", "png", "pdf"], accept_multiple_files=True)

# 메인 화면
st.title("📝 기계설비 성능점검 보고서 검토자문 시스템")
st.markdown("점검업체의 **점검결과**가 적정하게 작성되었는지 AI가 검토하고, **[적정]**과 **[보완필요]**로 구분하여 결과를 제공합니다.")

if not st.session_state.api_key:
    st.info("💡 좌측 사이드바에 Google Gemini API Key를 입력해 주세요.")
else:
    if uploaded_files and selected_model:
        col1, col2 = st.columns([1, 1.2])
        
        pdfs = [f for f in uploaded_files if f.type == "application/pdf"]
        imgs = [f for f in uploaded_files if "image" in f.type]
        
        with col1:
            st.subheader("📁 업로드된 자료")
            if pdfs: st.info(f"📄 PDF 문서: {pdfs[0].name}")
            if imgs:
                for img_file in imgs: st.image(Image.open(img_file), caption=img_file.name, use_column_width=True)
            
        with col2:
            st.subheader("🤖 검토자문 의견 생성")
            
            if st.button("🔍 검토자문서 생성 (전체 문서 100% 분석)", type="primary"):
                with st.spinner(f"'{selected_model}' 모델이 문서 내의 **모든 항목**을 하나도 빠짐없이 추출 중입니다..."):
                    try:
                        model = genai.GenerativeModel(selected_model)
                        
                        system_prompt = """
                        당신은 기계설비 성능점검 보고서 검토 전문가입니다.
                        성능점검업체가 '점검 기준'에 맞게 적절히 점검하고 보고서를 작성했는지 판단하세요.

                        [🚨 핵심 지시사항 - 절대 누락 금지 🚨]
                        1. 업로드된 문서(PDF 또는 사진들)에 존재하는 **모든 장비와 모든 점검 항목을 100% 단 하나도 빠짐없이** 추출하여 배열(items)에 추가하세요.
                        2. 문제가 없는 완벽한(적정한) 항목이라고 해서 임의로 생략하거나 요약해서 건너뛰면 절대 안 됩니다. 문서의 처음부터 끝까지 모든 항목을 반복해서 추출하세요.

                        [분석 및 판단 기준]
                        1. 페이지 번호: 해당 항목이 보고서의 몇 페이지에 있는지 파악하세요. (모르면 '확인 불가')
                        2. 업체 점검결과 추출: 보고서에 점검업체가 표기한 '점검결과'([적합], [조치필요], [부적합], [해당없음] 등)를 그대로 추출하세요.
                        3. 사진 속 수치 판독: 첨부된 계측기 사진 등에서 측정된 수치를 정확하게 읽으세요.
                        4. 검토자문의견 작성: 사진에서 확인된 수치와 '점검 기준'을 비교하여, 업체의 '점검결과'가 적정하게 작성되었는지 서술하세요. 
                           - 사진 수치가 안 보이면 "판독 불가"로 기재하고 보완필요 판정을 내리세요.

                        [JSON 출력 양식]
                        {
                          "items": [
                            {
                              "page_number": "[페이지 번호 (예: 12p)]",
                              "item_name": "[점검 항목명]",
                              "criteria": "[점검 기준]",
                              "reported_result": "[점검업체의 점검결과 (예: 적합)]",
                              "photo_value": "[사진에서 인식된 측정값]",
                              "report_value": "[보고서 텍스트 기재 내용]",
                              "ai_judgment": "[적정 또는 보완필요 (둘 중 하나만 기재)]",
                              "expert_detailed_opinion": "[검토자문의견: 점검결과가 적정한지/보완이 필요한지 사진 수치를 근거로 상세 서술]"
                            }
                          ]
                        }
                        
                        순수 JSON 형식으로만 응답하세요.
                        """
                        
                        contents_to_send = [system_prompt]
                        if pdfs:
                            temp_pdf_path = "temp_report.pdf"
                            with open(temp_pdf_path, "wb") as f: f.write(pdfs[0].read())
                            uploaded_pdf = genai.upload_file(temp_pdf_path)
                            while uploaded_pdf.state.name == "PROCESSING": time.sleep(2); uploaded_pdf = genai.get_file(uploaded_pdf.name)
                            contents_to_send.append(uploaded_pdf)
                        
                        if imgs:
                            pil_images = [Image.open(img) for img in imgs]
                            contents_to_send.extend(pil_images)
                        
                        response = model.generate_content(contents_to_send)
                        
                        if pdfs:
                            genai.delete_file(uploaded_pdf.name)
                            if os.path.exists(temp_pdf_path): os.remove(temp_pdf_path)
                        
                        raw_text = response.text.replace("```json", "").replace("```", "").strip()
                        st.session_state.report_data = json.loads(raw_text)
                        
                    except Exception as e:
                        st.error(f"❌ 분석 오류: {str(e)}")

            if st.session_state.report_data:
                data = st.session_state.report_data
                items = data.get("items", [])
                
                if items:
                    df = pd.DataFrame(items)
                    df_export = df.rename(columns={
                        "page_number": "페이지",
                        "item_name": "점검 항목",
                        "criteria": "점검 기준",
                        "reported_result": "업체 점검결과",
                        "photo_value": "사진 판독값",
                        "report_value": "보고서 기재값",
                        "ai_judgment": "AI 판정",
                        "expert_detailed_opinion": "검토자문의견"
                    })
                    
                    csv = df_export.to_csv(index=False, encoding='utf-8-sig')
                    
                    st.download_button(
                        label="📥 한눈에 보기 (엑셀 파일로 다운로드)",
                        data=csv,
                        file_name="기계설비_검토자문의견.csv",
                        mime="text/csv",
                    )
                st.markdown("---")

                adequate_indices = [i for i, item in enumerate(items) if item.get("ai_judgment", "") == "적정"]
                inadequate_indices = [i for i, item in enumerate(items) if item.get("ai_judgment", "") != "적정"]
                
                tab1, tab2 = st.tabs([f"🟢 적정 항목 ({len(adequate_indices)}건)", f"🔴 보완필요 항목 ({len(inadequate_indices)}건)"])
                
                def display_item(index, item_data):
                    st.markdown(f"#### 📌 [Page: {item_data.get('page_number', '확인불가')}] {item_data.get('item_name')}")
                    
                    col_a, col_b = st.columns(2)
                    with col_a: st.markdown(f"- **점검 기준:** {item_data.get('criteria')}")
                    with col_b: st.markdown(f"- **업체 점검결과:** `{item_data.get('reported_result')}`")
                    
                    st.markdown(f"- **사진 판독값:** `{item_data.get('photo_value')}`")
                    st.markdown(f"- **보고서 기재내용:** `{item_data.get('report_value')}`")
                    
                    if item_data.get("ai_judgment") == "적정":
                        st.success(f"🟢 **[적정] 검토자문의견:**\n\n{item_data.get('expert_detailed_opinion')}")
                    else:
                        st.error(f"🔴 **[{item_data.get('ai_judgment')}] 검토자문의견:**\n\n{item_data.get('expert_detailed_opinion')}")
                        
                    with st.expander("💬 자문의견 수정 지시하기"):
                        feedback = st.text_input("수정 지시사항을 입력하세요:", key=f"fb_{index}")
                        if st.button("AI 의견 수정 적용", key=f"btn_{index}"):
                            if feedback:
                                with st.spinner("수정 중..."):
                                    edit_model = genai.GenerativeModel(selected_model)
                                    edit_prompt = f"기존 의견: {item_data.get('expert_detailed_opinion')}\n지시사항: {feedback}\n지시를 반영해 기존 의견을 고쳐서 한 문장으로 답변해."
                                    edit_response = edit_model.generate_content(edit_prompt)
                                    st.session_state.report_data["items"][index]["expert_detailed_opinion"] = edit_response.text.strip()
                                    st.rerun()
                    st.markdown("<hr style='margin-top: 10px; margin-bottom: 20px;'>", unsafe_allow_html=True)

                with tab1:
                    if not adequate_indices: st.info("적정 판정을 받은 항목이 없습니다.")
                    for i in adequate_indices: display_item(i, items[i])
                        
                with tab2:
                    if not inadequate_indices: st.info("보완이 필요한 항목이 없습니다.")
                    for i in inadequate_indices: display_item(i, items[i])
