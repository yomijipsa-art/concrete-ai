import streamlit as st
import google.generativeai as genai
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import io
import os

# --- [설정 구간] ---
# 1. Gemini API 키 설정 (Streamlit Secrets 사용)
try:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
except Exception:
    st.error("Streamlit Cloud 설정에서 'GEMINI_API_KEY'를 등록해주세요.")

# 모델 설정 (가장 빠르고 효율적인 1.5 Flash 모델)
model = genai.GenerativeModel('gemini-3-flash-preview')

# 2. 고정된 엑셀 양식 파일 이름
TEMPLATE_FILE = "template.xlsx" 

# 3. 텍스트 데이터가 들어갈 셀 위치 (Sheet1)
text_cell_map = {
    "공사명": "D3",
    "타설위치": "D6",
    "타설규격": "D4",
    "슬럼프": "D11",
    "공기량": "D16",
    "염화물": "D29",
    "온도": "D34",
    "단위수량": "D23",
    "타설일자": "D5",
    "업체명": "L6"
}

# 4. 사진이 들어갈 셀 위치 (Sheet2)
image_cell_map = {
    "사진1": "B3",  # 1번 사진 위치
    "사진2": "B7"   # 2번 사진 위치
}

# 사진 고정 크기 설정 (픽셀 단위: 가로 406px, 세로 614px)
# 이는 약 10.74cm * 16.25cm 크기입니다.
FIXED_WIDTH = 614
FIXED_HEIGHT = 406
# ------------------

st.title("🏗️ 스마트 건설 현장 보고서 생성기")
st.write("사진 2장을 올리면 AI가 분석하여 엑셀 보고서를 만들어줍니다.")

# 사진 업로드 (2장 제한)
uploaded_images = st.file_uploader("사진 2장을 한꺼번에 선택해 주세요.", type=["jpg", "png", "jpeg", "mpo"], accept_multiple_files=True)

if len(uploaded_images) == 2:
    st.success("✅ 사진 2장이 업로드되었습니다.")
    
    if st.button("🚀 보고서 생성 시작"):
        try:
            with st.spinner("AI 분석 및 엑셀 생성 중..."):
                # --- [STEP 1: 2번 사진에서 텍스트 추출] ---
                # 2번째로 업로드된 사진을 분석 대상으로 삼습니다.
                img2_pill = PILImage.open(uploaded_images[1])
                
                prompt = """
                이 사진에서 공사명, 타설위치, 타설규격, 슬럼프, 공기량, 염화물, 온도, 단위수량, 타설일자, 업체명을 찾아줘. 
                결과는 반드시 '항목: 값' 형식으로 한 줄씩 써줘. 다른 설명은 하지 마.
                """
                response = model.generate_content([prompt, img2_pill])
                extracted_text = response.text
                
                # 데이터 파싱 로직
                parsed_data = {}
                for line in extracted_text.split('\n'):
                    if ":" in line:
                        k, v = line.split(":", 1)
                        # AI 답변에 포함될 수 있는 * 기호 등 제거
                        clean_k = k.replace("*", "").strip()
                        clean_v = v.replace("*", "").strip()
                        parsed_data[clean_k] = clean_v

                # --- [STEP 2: 엑셀 파일 작업] ---
                if not os.path.exists(TEMPLATE_FILE):
                    st.error(f"폴더 내에 '{TEMPLATE_FILE}' 파일이 없습니다.")
                    st.stop()

                wb = load_workbook(TEMPLATE_FILE)
                
                # Sheet1: 데이터 입력
                ws1 = wb.worksheets[0]
                for key, cell in text_cell_map.items():
                    val = parsed_data.get(key, "데이터 없음")
                    ws1[cell] = val

                # Sheet2: 사진 삽입 (MPO 및 호환성 오류 수정 버전)
                ws2 = wb.worksheets[1]
                
                for i, upload_file in enumerate(uploaded_images):
                    label = f"사진{i+1}"
                    cell_pos = image_cell_map[label]
                    
                    # [핵심 수정] 모든 이미지를 표준 RGB JPEG로 변환하여 처리
                    img = PILImage.open(upload_file)
                    if img.mode != 'RGB':
                        img = img.convert('RGB')
                    
                    img_path = f"temp_img_{i}.jpg"
                    img.save(img_path, format="JPEG")
                    
                    # 엑셀 이미지 객체 생성 및 크기 고정
                    xl_img = Image(img_path)
                    xl_img.width = FIXED_WIDTH
                    xl_img.height = FIXED_HEIGHT
                    
                    ws2.add_image(xl_img, cell_pos)

                # --- [STEP 3: 파일 저장 및 다운로드] ---
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)

                st.success("🎉 보고서 생성이 완료되었습니다!")
                
                st.download_button(
                    label="📥 완성된 엑셀 다운로드",
                    data=output,
                    file_name=f"현장보고서_{parsed_data.get('타설위치', '결과')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # 작업 종료 후 임시 이미지 파일 삭제
                for i in range(len(uploaded_images)):
                    tmp_file = f"temp_img_{i}.jpg"
                    if os.path.exists(tmp_file):
                        os.remove(tmp_file)

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")

elif len(uploaded_images) > 0:
    st.warning("사진을 반드시 **2장** 선택해야 합니다.")

