import streamlit as st
import google.generativeai as genai
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import io
import os

# --- [설정 구간] ---
# 1. Gemini API 키 설정
genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
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
    "사진2": "B7"  # 2번 사진 위치
}
# ------------------

st.title("🏗️ 스마트 건설 현장 보고서 생성기")

# 사진 업로드 (2장 제한)
uploaded_images = st.file_uploader("사진 2장을 한꺼번에 선택해 주세요.", type=["jpg", "png", "jpeg"], accept_multiple_files=True)

if len(uploaded_images) == 2:
    st.success("사진 2장이 준비되었습니다.")
    
    if st.button("🚀 보고서 생성 시작"):
        try:
            with st.spinner("AI가 2번 사진에서 데이터를 추출하고 엑셀을 만드는 중..."):
                # --- [STEP 1: 2번 사진에서 텍스트 추출] ---
                img2_data = PILImage.open(uploaded_images[1]) # 2번째 사진
                prompt = """
                이 사진에서 공사명, 타설위치, 타설규격, 슬럼프, 공기량, 염화물, 온도, 단위수량, 타설일자, 업체명을 찾아줘. 
                결과는 반드시 '항목: 값' 형식으로 한 줄씩 써줘.
                """
                response = model.generate_content([prompt, img2_data])
                extracted_text = response.text
                
                # 추출된 텍스트 파싱 (간단한 예시용)
                # 실제로는 AI 결과에 따라 보정이 필요할 수 있습니다.
                parsed_data = {}
                for line in extracted_text.split('\n'):
                    if ":" in line:
                        k, v = line.split(":", 1)
                        parsed_data[k.strip()] = v.strip()

                # --- [STEP 2: 엑셀 파일 작업] ---
                wb = load_workbook(TEMPLATE_FILE)
                
                # 1번째 시트에 텍스트 데이터 입력
                ws1 = wb.worksheets[0] # 첫 번째 시트
                for key, cell in text_cell_map.items():
                    # AI가 추출한 데이터가 parsed_data에 있다면 입력 (없으면 기본값)
                    val = parsed_data.get(key, "데이터 없음")
                    ws1[cell] = val

                # 2번째 시트에 사진 삽입
                ws2 = wb.worksheets[1] # 두 번째 시트
                
              for i, upload_file in enumerate(uploaded_images):
                    label = f"사진{i+1}"
                    cell_pos = image_cell_map[label]
                    
                    # 1. 파일을 Pillow 이미지 객체로 읽기
                    img = PILImage.open(upload_file)
                    
                    # 2. 만약 MPO나 PNG(투명도) 등 특이한 형식이면 표준 RGB로 변환
                    if img.mode != 'RGB':
                        img = img.convert('RGB')
                    
                    # 3. 표준 JPEG 파일로 임시 저장
                    img_path = f"temp_img_{i}.jpg"
                    img.save(img_path, format="JPEG")
                    
                    # 4. 엑셀에 삽입
                    xl_img = Image(img_path)
                    xl_img.width = 406  # 가로 10.74cm 근사치
                    xl_img.height = 614 # 세로 16.25cm 근사치
                    
                    ws2.add_image(xl_img, cell_pos)

                # --- [STEP 3: 파일 저장 및 다운로드] ---
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)

                st.success("✅ 보고서가 완성되었습니다!")
                
                st.download_button(
                    label="📥 완성된 엑셀 다운로드",
                    data=output,
                    file_name="현장보고서_완성본.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # 임시 이미지 파일 삭제
                for i in range(len(uploaded_images)):
                    if os.path.exists(f"temp_img_{i}.jpg"):
                        os.remove(f"temp_img_{i}.jpg")

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")

elif len(uploaded_images) > 0:

    st.warning("사진을 반드시 **2장** 선택해야 합니다.")

