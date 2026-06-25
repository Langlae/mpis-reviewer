import streamlit as st
import google.generativeai as genai
from PIL import Image
import json
import os
import time
import pandas as pd
import io
import urllib.parse
from datetime import datetime
from openpyxl.styles import Alignment
import streamlit.components.v1 as components
from pptx_maker import create_review_pptx


def _parse_json_with_repair(text: str) -> dict:
    """JSON 파싱 시도 후 실패하면 잘린 응답을 복구하여 재시도합니다."""
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        # 마지막으로 완전히 닫힌 items 항목을 찾아 JSON을 재조립합니다.
        search_key = '"expert_detailed_opinion"'
        pos = 0
        last_item_end = None

        while True:
            found = text.find(search_key, pos)
            if found == -1:
                break
            colon = text.find(":", found + len(search_key))
            if colon == -1:
                break
            q_start = text.find('"', colon + 1)
            if q_start == -1:
                break
            i = q_start + 1
            while i < len(text):
                if text[i] == "\\":
                    i += 2
                elif text[i] == '"':
                    break
                else:
                    i += 1
            close = text.find("}", i)
            if close == -1:
                break
            last_item_end = close
            pos = close + 1

        if last_item_end is None:
            raise ValueError(
                "AI 응답이 너무 길어 JSON을 복구할 수 없습니다. "
                "더 적은 페이지로 나눠 업로드해 보세요."
            )

        truncated = text[: last_item_end + 1].rstrip().rstrip(",")
        repaired = truncated + "\n  ]\n}"
        return json.loads(repaired)


# --- 세션 상태 초기화 ---
if "api_key" not in st.session_state:
    try:
        st.session_state.api_key = st.secrets["GEMINI_API_KEY"]
    except KeyError:
        st.session_state.api_key = ""

if "history" not in st.session_state:
    st.session_state.history = []  # [{name, timestamp, data}, ...]

if "selected_idx" not in st.session_state:
    st.session_state.selected_idx = None

st.set_page_config(page_title="기계설비 성능점검 검토자문 시스템", page_icon="📝", layout="wide")

# --- 사이드바 ---
with st.sidebar:
    st.header("⚙️ 설정 및 파일 업로드")

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
            available_models = [
                m.name.replace("models/", "")
                for m in genai.list_models()
                if "generateContent" in m.supported_generation_methods
            ]
            if available_models:
                default_idx = next(
                    (i for i, n in enumerate(available_models) if "pro" in n), 0
                )
                selected_model = st.selectbox("🤖 사용할 AI 모델", available_models, index=default_idx)
            else:
                st.error("⚠️ 사용 가능한 모델이 없습니다.")
        except Exception:
            st.error("⚠️ API 키가 유효하지 않습니다.")

    st.markdown("---")
    uploaded_files = st.file_uploader(
        "보고서 업로드 (사진 여러 장 또는 PDF)",
        type=["jpg", "jpeg", "png", "pdf"],
        accept_multiple_files=True,
    )

    # --- 검토 기록 (히스토리) ---
    if st.session_state.history:
        st.markdown("---")
        st.subheader("📋 검토 기록")
        for i, record in enumerate(st.session_state.history):
            is_selected = i == st.session_state.selected_idx
            label = f"{'▶ ' if is_selected else ''}{record['name'][:22]} ({record['timestamp']})"
            if st.button(
                label,
                key=f"hist_{i}",
                use_container_width=True,
                type="primary" if is_selected else "secondary",
            ):
                st.session_state.selected_idx = i
                st.rerun()

# --- 메인 화면 ---
st.title("📝 기계설비 성능점검 보고서 검토자문 시스템")
st.markdown(
    "점검업체의 **점검결과**가 적정하게 작성되었는지 AI가 검토하고, "
    "**[적정]**과 **[보완필요]**로 구분하여 결과를 제공합니다."
)

if not st.session_state.api_key:
    st.info("💡 좌측 사이드바에 Google Gemini API Key를 입력해 주세요.")
else:
    # 파일 업로드 + 생성 섹션
    if uploaded_files and selected_model:
        pdfs = [f for f in uploaded_files if f.type == "application/pdf"]
        imgs = [f for f in uploaded_files if "image" in f.type]

        col1, col2 = st.columns([1, 1.2])

        with col1:
            st.subheader("📁 업로드된 자료")
            if pdfs:
                st.info(f"📄 PDF 문서: {pdfs[0].name}")
            for img_file in imgs:
                img_file.seek(0)
                st.image(Image.open(img_file), caption=img_file.name, use_container_width=True)

        with col2:
            st.subheader("🤖 검토자문 의견 생성")
            if st.button("🔍 검토자문서 생성 (전체 문서 100% 분석)", type="primary"):
                with st.spinner(
                    f"'{selected_model}' 모델이 문서 내의 **모든 항목**을 하나도 빠짐없이 추출 중입니다..."
                ):
                    try:
                        max_tokens = 65536 if "2.5" in selected_model else 8192
                        model = genai.GenerativeModel(
                            selected_model,
                            generation_config=genai.GenerationConfig(max_output_tokens=max_tokens),
                        )

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
                        4. 검토자문의견 작성 (아래 [검토자문의견 작성 원칙]을 반드시 따르세요)

                        [🔴 검토자문의견 작성 원칙 - 핵심 품질 기준 🔴]
                        검토자문의견(expert_detailed_opinion)은 반드시 다음 원칙에 따라 **정량적·기술적**으로 작성하세요.
                        절대로 "양호", "적합", "이상없음" 등의 정성적 표현만으로 작성하지 마세요.

                        ▶ 원칙 1: 설계값 대비 측정값 비율을 반드시 계산하여 명시
                          - 보고서에 설계값(명판, 계산서, 사양서 등)이 기재된 경우, 측정값이 설계값의 몇 %인지 계산하세요.
                          - 작성 예시: "유량 측정값이 341.74m³/h로 설계유량 370m³/h의 92.4%이며, 점검기준(±10% 이내)을 충족하므로 적합함."
                          - 설계값이 보고서에 없으면 "설계값 미기재로 정량 비교 불가 — 보완 필요"라고 명시하세요.

                        ▶ 원칙 2: 단위 환산이 필요한 경우 직접 계산하여 서술
                          - 점검기준이 '풍량(CMH)' 기준인데 사진에 '풍속(m/s)'만 있는 경우:
                            풍량(CMH) = 풍속(m/s) × 덕트단면적(m²) × 3600 으로 환산하세요.
                            덕트 사이즈가 보고서에 있으면 직접 계산하고, 없으면 "풍속만 기재되어 있고 덕트 사이즈 미확인으로 풍량 환산 불가 — 보완 필요"라고 서술하세요.
                          - 브라인 온도, 전력, 압력 등 다른 단위도 동일하게 적용하세요.

                        ▶ 원칙 3: 운전 조건(정격/일반) 명시 및 비교 기준 확인
                          - 점검기준이 '정격운전 기준'인 경우, 측정이 정격운전 상태에서 이루어졌는지 확인하세요.
                          - 정격운전 상태가 아닌 일반운전 상태에서 측정한 값을 정격값과 비교한 경우: "정격운전 조건이 아닌 일반운전(인버터 42Hz) 상태에서 측정한 값을 정격값(60Hz)과 비교하여 기준 적용이 부적절함 — 보완 필요"라고 명시하세요.
                          - 인버터 운전 시에는 인버터 출력 주파수 대비 측정값 비율로 비교 가능함을 안내하세요.

                        ▶ 원칙 4: 점검결과 판단 근거를 구체적으로 서술
                          - 단순히 "사유: 양호" 또는 "적합" 한 단어로 끝나는 경우 → 반드시 "보완필요" 판정
                          - 반드시 "무엇을 측정했고(측정값), 설계값/기준값은 얼마이며, 그 비율은 얼마이고, 따라서 기준을 충족/불충족한다"는 논리 구조로 서술하세요.
                          - 사진 수치가 판독 불가인 경우: "측정 사진의 수치 판독 불가 — 보완 필요"로 기재하고 보완필요 판정을 내리세요.

                        ▶ 원칙 5: 보완이 필요한 경우 구체적인 개선 방향 제시
                          - 단순히 "보완필요"가 아니라, 무엇을 어떻게 보완해야 하는지 기술적으로 서술하세요.
                          - 예: "설계유량 대비 측정유량 비율(%)을 계산하여 기재하고, 기준 충족 여부를 정량적으로 서술할 것."

                        [JSON 출력 양식]
                        {
                          "basic_info": {
                            "building_name": "[건축물명]",
                            "inspection_company": "[성능점검업체명]",
                            "company_representative": "[업체 대표자명]",
                            "building_address": "[건축물 주소 (전체 주소)]",
                            "management_entity": "[관리주체(건물주) 이름]",
                            "management_address": "[관리주체 주소 (전체 주소)]"
                          },
                          "items": [
                            {
                              "page_number": "[페이지 번호 (예: 12p)]",
                              "item_name": "[점검 항목명]",
                              "criteria": "[점검 기준]",
                              "reported_result": "[점검업체의 점검결과 (예: 적합)]",
                              "photo_value": "[사진에서 인식된 측정값 (단위 포함)]",
                              "report_value": "[보고서 텍스트 기재 내용]",
                              "ai_judgment": "[적정 또는 보완필요 (둘 중 하나만 기재)]",
                              "expert_detailed_opinion": "[검토자문의견: 측정값·설계값·비율 계산을 포함한 정량적·기술적 서술. 보완필요 시 구체적 개선방향 제시]"
                            }
                          ]
                        }

                        basic_info의 모든 필드는 보고서 표지 또는 본문에서 반드시 찾아 기재하세요. 확인 불가 시 "확인 불가"로 기재하세요.
                        순수 JSON 형식으로만 응답하세요.
                        """

                        contents_to_send = [system_prompt]

                        if pdfs:
                            temp_pdf_path = "temp_report.pdf"
                            with open(temp_pdf_path, "wb") as f:
                                f.write(pdfs[0].read())
                            # SDK 버전에 따라 upload_file 위치가 다름
                            _upload = getattr(genai, "upload_file", None) or getattr(genai.files, "upload", None)
                            uploaded_pdf = _upload(temp_pdf_path)
                            _get = getattr(genai, "get_file", None) or getattr(genai.files, "get", None)
                            while uploaded_pdf.state.name == "PROCESSING":
                                time.sleep(2)
                                uploaded_pdf = _get(uploaded_pdf.name)
                            contents_to_send.append(uploaded_pdf)

                        if imgs:
                            for img_file in imgs:
                                img_file.seek(0)
                            contents_to_send.extend([Image.open(img) for img in imgs])

                        response = model.generate_content(contents_to_send)

                        if pdfs:
                            _delete = getattr(genai, "delete_file", None) or getattr(genai.files, "delete", None)
                            try:
                                _delete(uploaded_pdf.name)
                            except Exception:
                                pass
                            if os.path.exists(temp_pdf_path):
                                os.remove(temp_pdf_path)

                        raw_text = response.text.replace("```json", "").replace("```", "").strip()
                        new_data = _parse_json_with_repair(raw_text)

                        file_name = (
                            pdfs[0].name if pdfs else (imgs[0].name if imgs else "보고서")
                        )
                        timestamp = datetime.now().strftime("%m/%d %H:%M")
                        st.session_state.history.insert(
                            0, {"name": file_name, "timestamp": timestamp, "data": new_data}
                        )
                        st.session_state.selected_idx = 0
                        st.rerun()

                    except Exception as e:
                        st.error(f"❌ 분석 오류: {str(e)}")

        st.markdown("---")

    elif not st.session_state.history:
        st.info("💡 좌측에서 보고서를 업로드하고 '검토자문서 생성' 버튼을 눌러주세요.")

    # --- 결과 표시 (히스토리 기반, 파일 업로드 여부와 무관하게 유지) ---
    if st.session_state.selected_idx is not None and st.session_state.history:
        idx = st.session_state.selected_idx
        record = st.session_state.history[idx]
        data = record["data"]
        items = data.get("items", [])

        st.markdown(f"### 📄 {record['name']} — {record['timestamp']}")

        if items:
            df_export = pd.DataFrame(items).rename(
                columns={
                    "page_number": "페이지",
                    "item_name": "점검 항목",
                    "criteria": "점검 기준",
                    "reported_result": "업체 점검결과",
                    "photo_value": "사진 판독값",
                    "report_value": "보고서 기재값",
                    "ai_judgment": "AI 판정",
                    "expert_detailed_opinion": "검토자문의견",
                }
            )

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_export.to_excel(writer, index=False, sheet_name="검토결과")
                ws = writer.sheets["검토결과"]
                col_widths = [8, 30, 35, 14, 14, 20, 10, 60]
                for col_idx, width in enumerate(col_widths, 1):
                    ws.column_dimensions[ws.cell(1, col_idx).column_letter].width = width
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        cell.alignment = Alignment(wrap_text=True, vertical="top")
            buffer.seek(0)

            file_stem = record["name"].rsplit(".", 1)[0]

            dl_col1, dl_col2 = st.columns(2)
            with dl_col1:
                st.download_button(
                    label="📥 엑셀 다운로드",
                    data=buffer,
                    file_name=f"{file_stem}_검토자문의견.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with dl_col2:
                with st.spinner("PPT 생성 중..."):
                    pptx_bytes = create_review_pptx(data)
                st.download_button(
                    label="📊 PPT 다운로드 (Pretendard)",
                    data=pptx_bytes,
                    file_name=f"{file_stem}_검토자문의견.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )

        st.markdown("---")

        adequate_indices = [i for i, item in enumerate(items) if item.get("ai_judgment") == "적정"]
        inadequate_indices = [i for i, item in enumerate(items) if item.get("ai_judgment") != "적정"]

        tab0, tab1, tab2 = st.tabs(
            ["📋 기본정보", f"🟢 적정 항목 ({len(adequate_indices)}건)", f"🔴 보완필요 항목 ({len(inadequate_indices)}건)"]
        )

        def display_item(item_idx, item_data):
            st.markdown(
                f"#### 📌 [Page: {item_data.get('page_number', '확인불가')}] {item_data.get('item_name')}"
            )
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown(f"- **점검 기준:** {item_data.get('criteria')}")
            with col_b:
                st.markdown(f"- **업체 점검결과:** `{item_data.get('reported_result')}`")
            st.markdown(f"- **사진 판독값:** `{item_data.get('photo_value')}`")
            st.markdown(f"- **보고서 기재내용:** `{item_data.get('report_value')}`")

            if item_data.get("ai_judgment") == "적정":
                st.success(f"🟢 **[적정] 검토자문의견:**\n\n{item_data.get('expert_detailed_opinion')}")
            else:
                st.error(
                    f"🔴 **[{item_data.get('ai_judgment')}] 검토자문의견:**\n\n{item_data.get('expert_detailed_opinion')}"
                )

            with st.expander("💬 검토의견 재작성 요청 (AI가 직접 판단)"):
                feedback = st.text_area(
                    "검토 방향을 입력하세요 (AI가 항목 데이터를 직접 재분석하여 의견을 새로 작성합니다):",
                    key=f"fb_{idx}_{item_idx}",
                    height=80,
                    placeholder="예) 측정값이 기준 내에 있으므로 적정으로 재검토해주세요 / 풍량 환산 계산을 포함하여 다시 작성해주세요",
                )
                if st.button("🔄 AI 검토의견 재작성", key=f"btn_{idx}_{item_idx}"):
                    if feedback and selected_model:
                        with st.spinner("AI가 항목을 재분석하여 의견을 새로 작성 중..."):
                            edit_model = genai.GenerativeModel(selected_model)
                            edit_prompt = f"""당신은 기계설비 성능점검 보고서 검토 전문가입니다.
아래 점검 항목의 데이터를 직접 재분석하여 전문적인 검토자문의견을 새로 작성하고, 적정/보완필요 판정을 내려주세요.
사용자의 지시 방향을 참고하되, 의견 내용은 당신이 전문가로서 직접 판단하여 작성하세요.

[점검 항목 데이터]
- 점검 기준: {item_data.get('criteria', '')}
- 업체 점검결과: {item_data.get('reported_result', '')}
- 사진 판독값: {item_data.get('photo_value', '')}
- 보고서 기재값: {item_data.get('report_value', '')}
- 기존 AI 검토의견: {item_data.get('expert_detailed_opinion', '')}

[검토자 지시 방향]
{feedback}

[작성 원칙]
- 설계값 대비 측정값 비율(%)을 계산하여 정량적으로 서술
- 단순 "양호", "적합" 같은 정성적 표현만으로 끝내지 말 것
- 보완필요 시 구체적 개선 방향 제시

순수 JSON 형식으로만 응답하세요 (다른 텍스트 없이):
{{"ai_judgment": "적정 또는 보완필요 중 하나만", "expert_detailed_opinion": "재작성한 검토자문의견"}}"""

                            edit_response = edit_model.generate_content(edit_prompt)
                            raw = edit_response.text.replace("```json", "").replace("```", "").strip()
                            try:
                                result = json.loads(raw)
                                new_judgment = result.get("ai_judgment", item_data.get("ai_judgment"))
                                new_opinion  = result.get("expert_detailed_opinion", raw)
                                # 판정값 정규화
                                if "적정" in new_judgment and "보완" not in new_judgment:
                                    new_judgment = "적정"
                                else:
                                    new_judgment = "보완필요"
                            except Exception:
                                new_judgment = item_data.get("ai_judgment")
                                new_opinion  = raw

                            st.session_state.history[idx]["data"]["items"][item_idx]["ai_judgment"] = new_judgment
                            st.session_state.history[idx]["data"]["items"][item_idx]["expert_detailed_opinion"] = new_opinion
                            st.rerun()

            st.markdown(
                "<hr style='margin-top: 10px; margin-bottom: 20px;'>", unsafe_allow_html=True
            )

        def render_map(address: str):
            if not address or address == "확인 불가":
                st.caption("주소 정보가 없어 지도를 표시할 수 없습니다.")
                return
            encoded = urllib.parse.quote(address)
            map_html = (
                f'<iframe width="100%" height="280" frameborder="0" style="border:0; border-radius:8px;" '
                f'src="https://maps.google.com/maps?q={encoded}&output=embed" allowfullscreen></iframe>'
            )
            components.html(map_html, height=290)
            st.link_button("🗺️ Google Maps에서 열기", f"https://maps.google.com/maps?q={encoded}")

        with tab0:
            basic = data.get("basic_info", {})

            st.markdown("### 🏢 건축물 정보")
            col_l, col_r = st.columns(2)
            with col_l:
                st.markdown(f"**건축물명:** {basic.get('building_name', '확인 불가')}")
                st.markdown(f"**주소:** {basic.get('building_address', '확인 불가')}")
            with col_r:
                st.markdown(f"**성능점검업체명:** {basic.get('inspection_company', '확인 불가')}")
                st.markdown(f"**업체 대표자명:** {basic.get('company_representative', '확인 불가')}")
            render_map(basic.get("building_address", ""))

            st.markdown("---")
            st.markdown("### 👤 관리주체 정보")
            col_l2, col_r2 = st.columns(2)
            with col_l2:
                st.markdown(f"**관리주체(건물주):** {basic.get('management_entity', '확인 불가')}")
            with col_r2:
                st.markdown(f"**주소:** {basic.get('management_address', '확인 불가')}")
            render_map(basic.get("management_address", ""))

        with tab1:
            if not adequate_indices:
                st.info("적정 판정을 받은 항목이 없습니다.")
            for i in adequate_indices:
                display_item(i, items[i])

        with tab2:
            if not inadequate_indices:
                st.info("보완이 필요한 항목이 없습니다.")
            for i in inadequate_indices:
                display_item(i, items[i])
