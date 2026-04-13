import streamlit as st
from streamlit_drawable_canvas import st_canvas
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from PIL import Image
import io
import base64
import xlsxwriter

# --- 앱 설정 ---
st.set_page_config(page_title="재구조화 동의서 관리 시스템", layout="centered")

# 구글 시트 연결
conn = st.connection("gsheets", type=GSheetsConnection)

# --- [함수] 엑셀 생성 로직 ---
def generate_excel_with_images(df):
    output = io.BytesIO()
    # 엑셀 엔진으로 XlsxWriter 사용
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='동의서_결과')
        workbook  = writer.book
        worksheet = writer.sheets['동의서_결과']

        # 엑셀 서식 설정 (가운데 정렬, 테두리)
        format_center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        
        # 열 너비 및 행 높이 조절 (사인이 잘 보이도록)
        worksheet.set_column('B:B', 15) # 이름 열
        worksheet.set_column('C:C', 25) # 서명 열
        
        for idx, row in df.iterrows():
            row_num = idx + 1 # 헤더 제외한 시작 행
            worksheet.set_row(row_num, 60) # 행 높이를 60으로 확대
            
            sig_data = str(row['서명'])
            if sig_data.startswith("'"): sig_data = sig_data[1:]
            
            if sig_data and sig_data != "":
                try:
                    # Base64를 이미지로 복구
                    img_data = base64.b64decode(sig_data)
                    img_io = io.BytesIO(img_data)
                    
                    # 엑셀 셀에 이미지 삽입 (C열 = index 2)
                    worksheet.insert_image(row_num, 2, 'signature.png', 
                                         {'image_data': img_io, 
                                          'x_scale': 0.5, 'y_scale': 0.5, # 크기 조절
                                          'x_offset': 5, 'y_offset': 5})
                    # 기존 문자열 데이터는 가리기 위해 빈 값으로 덮어쓰기 시도 (선택사항)
                    worksheet.write(row_num, 2, "", format_center)
                except:
                    pass
    return output.getvalue()

# --- 화면 구성 ---
st.title("2026 경기기공 재구조화 동의서")

# 탭 메뉴 구성 (서명하기 / 관리자용)
menu_tab1, menu_tab2 = st.tabs(["✍️ 서명하기", "📊 관리자 메뉴"])

# --- 탭 1: 서명하기 ---
with menu_tab1:
    with st.container(border=True):
        st.subheader("📢 재구조화 사업 안내")
        st.info("[안내 문구를 이곳에 입력하세요. 교육청 양식의 취지 등...]")

    try:
        df = conn.read(worksheet="동의서양식", ttl=0).fillna("")
        staff_list = df['성함'].tolist()
    except:
        st.error("시트를 불러올 수 없습니다. '동의서양식' 탭을 확인해 주세요.")
        st.stop()

    selected_name = st.selectbox("성함을 선택하세요", ["선택하세요"] + staff_list)

    if selected_name != "선택하세요":
        row_idx = df.index[df['성함'] == selected_name].tolist()[0]
        if df.at[row_idx, '서명']:
            st.warning("⚠️ 이미 서명하셨습니다. 다시 서명하면 덮어씁니다.")
        
        canvas_result = st_canvas(
            fill_color="rgba(255, 255, 255, 0)", stroke_width=3,
            stroke_color="#000000", background_color="#f0f2f6",
            height=150, width=300, drawing_mode="freedraw", key="sig"
        )

        if st.button("서명 제출", use_container_width=True):
            if canvas_result.image_data is not None:
                img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
                img.thumbnail((200, 100))
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                sig_b64 = base64.b64encode(buf.getvalue()).decode()

                df.at[row_idx, '서명'] = f"'{sig_b64}"
                conn.update(worksheet="동의서양식", data=df)
                st.success("제출 완료!")
                st.balloons()

# --- 탭 2: 관리자 메뉴 (엑셀 다운로드) ---
with menu_tab2:
    st.subheader("📥 최종 결과물 생성")
    st.write("아래 버튼을 누르면 모든 서명이 이미지로 포함된 엑셀 파일이 생성됩니다.")
    
    # 관리자 비밀번호 확인 (선택사항)
    pw = st.text_input("관리자 비밀번호", type="password")
    if pw == st.secrets.get("auth", {}).get("admin_password", "1234"):
        if st.button("🖼️ 이미지 포함 엑셀 다운로드", use_container_width=True):
            with st.spinner("엑셀 파일을 생성 중입니다..."):
                excel_data = generate_excel_with_images(df)
                st.download_button(
                    label="💾 엑셀 파일 받기",
                    data=excel_data,
                    file_name=f"재구조화_동의서_결과_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    elif pw != "":
        st.error("비밀번호가 틀렸습니다.")
